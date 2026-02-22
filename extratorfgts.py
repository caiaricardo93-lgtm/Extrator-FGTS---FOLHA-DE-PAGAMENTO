import re
from datetime import datetime
from pathlib import Path
import tempfile

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st


# ====== CONFIGURAÇÃO ======
EVENTOS = ["900", "902", "903", "908", "916", "917"]

# número no formato brasileiro: 1.234,56 ou 123,45
RE_MONEY_LINE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")

# Período: 01/03/2025 a 31/03/2025
RE_PERIODO = re.compile(r"Período:\s*(\d{2}/\d{2}/\d{4})")

# Header do funcionário:
# aceita nome colado em Adm (tipo SILVAAdm)
RE_HEADER = re.compile(
    r"Func:\s*(?P<mat>\d+)\s+(?P<nome>.+?)\s*Adm\s*(?P<adm>\d{2}/\d{2}/\d{4})\s*Dem:\s*(?P<dem>\d{2}/\d{2}/\d{4})?",
    re.DOTALL
)


def br_money_to_str(value: str) -> str:
    """Normaliza para '1234,56' (mantém vírgula)."""
    n = float(value.replace(".", "").replace(",", "."))
    return f"{n:.2f}".replace(".", ",")


def extrair_periodo_mm_aaaa(texto: str) -> str:
    m = RE_PERIODO.search(texto)
    if not m:
        return ""
    d = datetime.strptime(m.group(1), "%d/%m/%Y")
    return d.strftime("%m/%Y")


def pegar_valor_evento_por_linhas(linhas: list[str], codigo_evento: str) -> str:
    """
    Ao achar a linha == '900', procura o número monetário mais próximo para cima.
    """
    for i, ln in enumerate(linhas):
        if ln.strip() == codigo_evento:
            for j in range(i - 1, max(-1, i - 13), -1):
                v = linhas[j].strip()
                if RE_MONEY_LINE.match(v):
                    return br_money_to_str(v)
            return ""
    return ""


def extrair_pdf(path_pdf: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    doc = fitz.open(path_pdf)
    texto = "\n".join(doc.load_page(i).get_text("text") for i in range(doc.page_count))

    periodo = extrair_periodo_mm_aaaa(texto)

    # corta TOTAL EMPRESA (não incluir essa seção)
    idx_total = texto.upper().find("TOTAL EMPRESA")
    texto_util = texto[:idx_total] if idx_total != -1 else texto

    # localizar todos os blocos "Func:" (antes do TOTAL EMPRESA)
    func_positions = [m.start() for m in re.finditer(r"^Func:", texto_util, flags=re.MULTILINE)]

    blocos = []
    for i, s in enumerate(func_positions):
        e = func_positions[i + 1] if i + 1 < len(func_positions) else len(texto_util)
        blocos.append(texto_util[s:e])

    registros = []
    headers_falharam = 0

    for bloco in blocos:
        header_area = bloco[:900]
        hm = RE_HEADER.search(header_area)
        if not hm:
            headers_falharam += 1
            continue

        mat = hm.group("mat").strip()
        nome = " ".join(hm.group("nome").split())
        adm = hm.group("adm").strip()
        dem = (hm.group("dem") or "").strip()

        linhas = [l for l in bloco.splitlines() if l.strip()]
        ev = {c: pegar_valor_evento_por_linhas(linhas, c) for c in EVENTOS}

        registros.append({
            "Período": periodo,
            "Matrícula": mat,
            "Nome do funcionário": nome,
            "Data de admissão": adm,
            "Data de demissão": dem,
            "Ev.900 FGTS": ev["900"],
            "Ev.902": ev["902"],
            "Ev.903": ev["903"],
            "Ev.908": ev["908"],
            "Ev.916": ev["916"],
            "Ev.917": ev["917"],
        })

    df_base = pd.DataFrame(registros)
    df_check = pd.DataFrame([{
        "arquivo": Path(path_pdf).name,
        "periodo": periodo,
        "func_encontrados": len(func_positions),
        "func_extraidos": len(df_base),
        "headers_falharam": headers_falharam,
    }])

    return df_base, df_check


def extrair_varios_pdfs_em_memoria(paths_pdfs: list[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    bases = []
    checks = []

    for pdf in paths_pdfs:
        df_base, df_check = extrair_pdf(pdf)
        bases.append(df_base)
        checks.append(df_check)

    df_all = pd.concat(bases, ignore_index=True) if bases else pd.DataFrame()
    df_check_all = pd.concat(checks, ignore_index=True) if checks else pd.DataFrame()

    divergencias = df_check_all[df_check_all["func_encontrados"] != df_check_all["func_extraidos"]]
    if not divergencias.empty:
        raise RuntimeError(
            "Divergência na extração (func_encontrados != func_extraidos) em:\n"
            f"{divergencias.to_string(index=False)}"
        )

    return df_all, df_check_all


def gerar_excel_bytes(df_base: pd.DataFrame, df_conf: pd.DataFrame) -> bytes:
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_base.to_excel(writer, index=False, sheet_name="Base")
        df_conf.to_excel(writer, index=False, sheet_name="Conferência")
    output.seek(0)
    return output.read()


# ================== UI STREAMLIT ==================
st.set_page_config(page_title="Extrator de Eventos (PDF → Excel)", layout="wide")
st.title("Extrator de Eventos (PDF → Excel)")
st.caption("Faça upload dos PDFs e baixe o Excel com Base + Conferência.")

with st.sidebar:
    st.subheader("Configurações")
    eventos_str = st.text_input("Eventos (separados por vírgula)", ",".join(EVENTOS))
    EVENTOS[:] = [e.strip() for e in eventos_str.split(",") if e.strip()]
    st.write("Eventos ativos:", EVENTOS)

uploaded_files = st.file_uploader(
    "Envie 1 ou vários PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

col1, col2 = st.columns([1, 1])

if uploaded_files:
    if st.button("Processar PDFs", type="primary"):
        with st.spinner("Extraindo dados..."):
            try:
                # Salva uploads temporariamente para o fitz abrir por caminho
                tmp_dir = tempfile.TemporaryDirectory()
                pdf_paths = []
                for uf in uploaded_files:
                    p = Path(tmp_dir.name) / uf.name
                    p.write_bytes(uf.getbuffer())
                    pdf_paths.append(str(p))

                df_base, df_conf = extrair_varios_pdfs_em_memoria(pdf_paths)

                st.success("Extração concluída!")

                with col1:
                    st.subheader("Prévia — Base")
                    st.dataframe(df_base, use_container_width=True)

                with col2:
                    st.subheader("Conferência")
                    st.dataframe(df_conf, use_container_width=True)

                excel_bytes = gerar_excel_bytes(df_base, df_conf)
                st.download_button(
                    label="Baixar base_eventos.xlsx",
                    data=excel_bytes,
                    file_name="base_eventos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Fecha o diretório temporário (libera arquivos)
                tmp_dir.cleanup()

            except Exception as e:
                st.error("Deu erro na extração.")
                st.exception(e)
else:
    st.info("Envie os PDFs acima para habilitar o processamento.")
