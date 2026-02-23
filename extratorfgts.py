import re
from datetime import datetime
from pathlib import Path
import tempfile

import fitz  # PyMuPDF
import pandas as pd
import streamlit as st


# ====== CONFIGURAÇÃO ======
EVENTOS = ["900", "901", "902", "903", "908", "916", "917"]

# número no formato brasileiro: 1.234,56 ou 123,45
RE_MONEY_LINE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")

# Período completo
RE_PERIODO_TIPO = re.compile(
    r"Período:\s*(?P<ini>\d{2}/\d{2}/\d{4})\s*a\s*(?P<fim>\d{2}/\d{2}/\d{4})\s*(?P<tipo>.*)$",
    re.MULTILINE
)

RE_HEADER = re.compile(
    r"Func:\s*(?P<mat>\d+)\s+(?P<nome>.+?)\s*Adm\s*(?P<adm>\d{2}/\d{2}/\d{4})\s*Dem:\s*(?P<dem>\d{2}/\d{2}/\d{4})?",
    re.DOTALL
)


def br_money_to_str(value: str) -> str:
    n = float(value.replace(".", "").replace(",", "."))
    return f"{n:.2f}".replace(".", ",")


def extrair_periodo_e_tipo(texto: str) -> tuple[str, str]:
    m = RE_PERIODO_TIPO.search(texto)
    if not m:
        return "", ""

    d_ini = datetime.strptime(m.group("ini"), "%d/%m/%Y")
    periodo = d_ini.strftime("%m/%Y")

    tipo = (m.group("tipo") or "").strip()

    if not tipo:
        tail = texto[m.end():]
        for ln in tail.splitlines():
            ln = ln.strip()
            if ln and not ln.upper().startswith(("FUNC:", "TOTAL", "EVENTO", "CÓD", "COD")):
                tipo = ln
                break

    tipo = " ".join(tipo.split())
    return periodo, tipo


def pegar_valor_evento_por_linhas(linhas: list[str], codigo_evento: str) -> str:
    for i, ln in enumerate(linhas):
        if ln.strip() == codigo_evento:
            for j in range(i - 1, max(-1, i - 13), -1):
                v = linhas[j].strip()
                if RE_MONEY_LINE.match(v):
                    return br_money_to_str(v)
            return ""
    return ""


def extrair_pdf(path_pdf: str):
    doc = fitz.open(path_pdf)
    texto = "\n".join(doc.load_page(i).get_text("text") for i in range(doc.page_count))

    periodo, tipo = extrair_periodo_e_tipo(texto)

    idx_total = texto.upper().find("TOTAL EMPRESA")
    texto_util = texto[:idx_total] if idx_total != -1 else texto

    func_positions = [m.start() for m in re.finditer(r"^Func:", texto_util, flags=re.MULTILINE)]

    blocos = []
    for i, s in enumerate(func_positions):
        e = func_positions[i + 1] if i + 1 < len(func_positions) else len(texto_util)
        blocos.append(texto_util[s:e])

    registros = []
    headers_falharam = 0

    for bloco in blocos:
        hm = RE_HEADER.search(bloco[:900])
        if not hm:
            headers_falharam += 1
            continue

        linhas = [l for l in bloco.splitlines() if l.strip()]
        ev = {c: pegar_valor_evento_por_linhas(linhas, c) for c in EVENTOS}

        registros.append({
            "Período": periodo,
            "Matrícula": hm.group("mat").strip(),
            "Nome do funcionário": " ".join(hm.group("nome").split()),
            "Data de admissão": hm.group("adm").strip(),
            "Data de demissão": (hm.group("dem") or "").strip(),
            "TIPO": tipo,
            "Ev.900 FGTS": ev.get("900", ""),
            "Ev.902": ev.get("902", ""),
            "Ev.903": ev.get("903", ""),
            "Ev.908": ev.get("908", ""),
            "Ev.916": ev.get("916", ""),
            "Ev.917": ev.get("917", ""),
            "EV. 901": ev.get("901", ""),
        })

    df_base = pd.DataFrame(registros)
    df_check = pd.DataFrame([{
        "arquivo": Path(path_pdf).name,
        "periodo": periodo,
        "tipo": tipo,
        "func_encontrados": len(func_positions),
        "func_extraidos": len(df_base),
        "headers_falharam": headers_falharam,
    }])

    return df_base, df_check


def extrair_varios_pdfs_em_memoria(paths_pdfs):
    bases, checks = [], []
    for pdf in paths_pdfs:
        b, c = extrair_pdf(pdf)
        bases.append(b)
        checks.append(c)

    df_all = pd.concat(bases, ignore_index=True) if bases else pd.DataFrame()
    df_check_all = pd.concat(checks, ignore_index=True) if checks else pd.DataFrame()
    return df_all, df_check_all


def gerar_excel_bytes(df_base, df_conf):
    import io
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_base.to_excel(writer, index=False, sheet_name="Base")
        df_conf.to_excel(writer, index=False, sheet_name="Conferência")
    output.seek(0)
    return output.read()


# ================== UI ==================
st.set_page_config(page_title="Extrator de Eventos (PDF → Excel)", layout="wide")

st.title("Extrator de Eventos (PDF → Excel)")
st.markdown(
    "<p style='margin-top:-10px; color:gray; font-size:12px;'><i>Criado por Caiã Ricardo Grade</i></p>",
    unsafe_allow_html=True
)
st.caption("Faça upload dos PDFs e baixe o Excel com Base + Conferência.")

with st.sidebar:
    st.subheader("Configurações")
    eventos_str = st.text_input("Eventos (separados por vírgula)", ",".join(EVENTOS))
    EVENTOS[:] = [e.strip() for e in eventos_str.split(",") if e.strip()]
    st.write("Eventos ativos:", EVENTOS)

uploaded_files = st.file_uploader("Envie 1 ou vários PDFs", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Processar PDFs", type="primary"):
        with st.spinner("Extraindo dados..."):
            tmp_dir = tempfile.TemporaryDirectory()
            pdf_paths = []
            for uf in uploaded_files:
                p = Path(tmp_dir.name) / uf.name
                p.write_bytes(uf.getbuffer())
                pdf_paths.append(str(p))

            df_base, df_conf = extrair_varios_pdfs_em_memoria(pdf_paths)

            st.success("Extração concluída!")
            st.dataframe(df_base, use_container_width=True)

            excel_bytes = gerar_excel_bytes(df_base, df_conf)
            st.download_button(
                label="Baixar base_eventos.xlsx",
                data=excel_bytes,
                file_name="base_eventos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            tmp_dir.cleanup()
