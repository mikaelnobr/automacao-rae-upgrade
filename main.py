import streamlit as st
import sys
import os
import re
import json
import time
import tempfile
from io import BytesIO

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# --- BANCO DE DADOS DE PROFISSIONAIS ---
PROFISSIONAIS = {
    "FRANCISCO DAVID MENESES DOS SANTOS": {
        "empresa": "FRANCISCO DAVID MENESES DOS SANTOS - F. D. MENESES DOS SANTOS",
        "cnpj": "54.801.096/0001-16",
        "cpf_emp": "058.756.003-73",
        "nome_resp": "FRANCISCO DAVID MENESES DOS SANTOS",
        "cpf_resp": "058.756.003-73",
        "registro": "336241CE"
    },
    "PALLOMA TEIXEIRA DA SILVA": {
        "empresa": "PALLOMA TEIXEIRA DA SILVA - PALLOMA TEIXEIRA ARQUITETURA LTDA",
        "cnpj": "54.862.474/0001-71",
        "cpf_emp": "064.943.593-10",
        "nome_resp": "PALLOMA TEIXEIRA DA SILVA",
        "cpf_resp": "064.943.593-10",
        "registro": "A184355-9"
    },
    "SANDY PEREIRA CORDEIRO": {
        "empresa": "SANDY PEREIRA CORDEIRO - CS ENGENHARIA",
        "cnpj": "54.794.898/0001-46",
        "cpf_emp": "071.222.553-60",
        "nome_resp": "SANDY PEREIRA CORDEIRO",
        "cpf_resp": "071.222.553-60",
        "registro": "356882CE"
    },
    "TIAGO VICTOR DE SOUSA": {
        "empresa": "TIAGO VICTOR DE SOUSA - T V S ENGENHARIA E ASSESSORIA",
        "cnpj": "54.806.521/0001-60",
        "cpf_emp": "068.594.803-00",
        "nome_resp": "TIAGO VICTOR DE SOUSA",
        "cpf_resp": "068.594.803-00",
        "registro": "346856CE"
    }
}

# --- PATCH DE METADADOS ---
try:
    import importlib.metadata as metadata
except ImportError:
    import importlib_metadata as metadata

_original_version = metadata.version
def patched_version(package_name):
    try:
        return _original_version(package_name)
    except Exception:
        versions = {
            'docling': '2.15.0', 'docling-core': '2.9.0', 'docling-parse': '2.4.0',
            'docling-ibm-models': '1.1.0', 'pypdfium2': '4.30.0', 'openpyxl': '3.1.5',
            'transformers': '4.40.0', 'torch': '2.2.0', 'torchvision': '0.17.0',
            'timm': '0.9.16', 'optree': '0.11.0'
        }
        return versions.get(package_name, "1.0.0")
metadata.version = patched_version

# --- IMPORTAÇÃO ---
try:
    import pandas as pd
    from openpyxl import load_workbook
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.datamodel.base_models import InputFormat
    import google.generativeai as genai
    DEPENDENCIAS_OK = True
except ImportError as e:
    DEPENDENCIAS_OK = False
    ERRO_IMPORT = str(e)

# --- ESTILIZAÇÃO ---
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    .stButton>button { width: 100%; border-radius: 8px; height: 3.5em; background-color: #4f46e5; color: white; font-weight: bold; border: none; }
    .stDownloadButton>button { width: 100%; border-radius: 8px; background-color: #059669; color: white; border: none; }
    </style>
    """, unsafe_allow_html=True)

@st.cache_resource
def get_converter():
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = True 
    return DocumentConverter(
        allowed_formats=[InputFormat.PDF],
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)}
    )

def call_gemini(api_key, prompt):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    for attempt in range(3):
        try:
            response = model.generate_content(prompt, generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0.1))
            return json.loads(response.text)
        except:
            time.sleep(2)
    raise Exception("Erro ao consultar a IA.")

def main():
    st.title("🏛️ Automação RAE CAIXA")
    st.markdown("##### Inteligência Artificial para Engenharia (Laudo + PLS + Alvará)")

    if not DEPENDENCIAS_OK:
        st.error(f"Erro de dependências: {ERRO_IMPORT}")
        return

    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")
        st.divider()
        st.subheader("👤 Responsável Técnico")
        resp_selecionado = st.selectbox("Selecione o Profissional:", options=list(PROFISSIONAIS.keys()))
        st.caption("v4.0 - Multi-Documentos")

    st.subheader("📂 Upload de Documentos")
    col1, col2 = st.columns(2)
    with col1:
        pdf_laudo = st.file_uploader("1. Laudo Técnico (PDF)", type=["pdf"])
        pdf_pls = st.file_uploader("3. PLS (PDF) - Opcional", type=["pdf"])
    with col2:
        excel_template = st.file_uploader("2. Planilha RAE (.xlsm)", type=["xlsm"])
        pdf_alvara = st.file_uploader("4. Alvará (PDF) - Opcional", type=["pdf"])

    if st.button("🚀 INICIAR PROCESSAMENTO COMPLETO"):
        if not api_key or not pdf_laudo or not excel_template:
            st.warning("A chave API, o Laudo e a Planilha são obrigatórios.")
            return

        try:
            with st.status("Processando documentos...", expanded=True) as status:
                converter = get_converter()
                texto_total = ""

                # Processar Documentos Enviados
                documentos = [("Laudo", pdf_laudo), ("PLS", pdf_pls), ("Alvará", pdf_alvara)]
                
                for nome, doc in documentos:
                    if doc:
                        st.write(f"📖 Lendo {nome}...")
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(doc.getbuffer())
                            res = converter.convert(tmp.name)
                            texto_total += f"\n--- INÍCIO DO DOCUMENTO: {nome} ---\n"
                            texto_total += res.document.export_to_markdown()
                            os.remove(tmp.name)

                st.write("🧠 IA: Analisando e cruzando informações...")
                prompt = f"""
                Atue como engenheiro revisor da CAIXA. Analise os documentos (Laudo, PLS, Alvará) e extraia para JSON:
                - CAMPOS: proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel, lat_s, long_w, etapas_original
                - OFICIO: Número após a matrícula em DOCUMENTOS (ex: 12345 / 3 / CE, ofício é 3).
                - COORDENADAS: GMS puro (ex: 06°24'08.8"). SEM letras.
                - CRONOGRAMA: etapas_original (Identifique no Cronograma da PLS ou do Laudo).
                - TABELAS: 'incidencias' (20 números coluna PESO % do orçamento), 'acumulado' (percentuais % ACUMULADO do cronograma).
                
                CONTEÚDO DOS DOCUMENTOS:
                {texto_total}
                """
                
                dados = call_gemini(api_key, prompt)

                st.write("📊 Gravando na planilha...")
                wb = load_workbook(BytesIO(excel_template.read()), keep_vba=True)
                wb.calculation.fullCalcOnLoad = True
                
                def to_f(v):
                    if isinstance(v, (int, float)): return v
                    try: return float(str(v).replace(',', '.').replace('%', '').strip())
                    except: return 0

                if "Início Vistoria" in wb.sheetnames:
                    ws = wb["Início Vistoria"]
                    mapping = {
                        "G43": "proponente", "AJ43": "cpf_cnpj", "AP43": "ddd", "AR43": "telefone",
                        "G49": "endereco", "AD49": "lat_s", "AH49": "long_w", "AL49": "complemento",
                        "G51": "bairro", "V51": "cep", "AA51": "municipio", "AS51": "uf_vistoria",
                        "AS53": "uf_registro", "G53": "valor_terreno", "Q53": "matricula",
                        "AA53": "oficio", "AJ53": "comarca"
                    }
                    for cell, key in mapping.items():
                        val = dados.get(key, "")
                        ws[cell] = to_f(val) if key == "valor_terreno" else str(val).upper()
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                if "RAE" in wb.sheetnames:
                    ws_rae = wb["RAE"]
                    ws_rae.sheet_state = 'visible'
                    ws_rae["AH66"] = to_f(dados.get("valor_imovel", 0))
                    ws_rae["AS66"] = to_f(dados.get("etapas_original", 0))
                    
                    prof = PROFISSIONAIS[resp_selecionado]
                    ws_rae["I315"], ws_rae["I316"], ws_rae["U316"] = prof["empresa"].upper(), prof["cnpj"], prof["cpf_emp"]
                    ws_rae["AE315"], ws_rae["AE316"], ws_rae["AO316"] = prof["nome_resp"].upper(), prof["cpf_resp"], prof["registro"].upper()
                    
                    incs, acus = dados.get("incidencias", []), dados.get("acumulado", [])
                    for i in range(20): ws_rae[f"S{69+i}"] = to_f(incs[i]) if i < len(incs) else 0
                    for i in range(len(acus)): 
                        if i < 37: ws_rae[f"AE{72+i}"] = to_f(acus[i])

                output = BytesIO()
                wb.save(output)
                
                proponente = dados.get("proponente", "").strip()
                primeiro_nome = proponente.split(' ')[0].upper() if proponente else "FINAL"
                
                status.update(label="✅ Processamento concluído!", state="complete", expanded=False)
                st.balloons()
                st.download_button(label=f"📥 BAIXAR RAE - {primeiro_nome}", data=output.getvalue(), file_name=f"RAE_{primeiro_nome}.xlsm", mime="application/vnd.ms-excel.sheet.macroEnabled.12")

        except Exception as e:
            st.error(f"Erro: {e}")

if __name__ == "__main__":
    main()
