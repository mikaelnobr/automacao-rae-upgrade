import streamlit as st
import sys
import os
import re
import json
import time
import tempfile
import gc
from io import BytesIO

# --- 1. CONFIGURAÇÃO DE PÁGINA (DEVE SER A PRIMEIRA LINHA) ---
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# --- 2. BANCO DE DADOS DE PROFISSIONAIS ---
PROFISSIONAIS = {
    "FRANCISCO DAVID MENESES DOS SANTOS": {
        "empresa": "FRANCISCO DAVID MENESES DOS SANTOS - F. D. MENESES DOS SANTOS",
        "cnpj": "54.801.096/0001-16", "cpf_emp": "058.756.003-73",
        "nome_resp": "FRANCISCO DAVID MENESES DOS SANTOS", "cpf_resp": "058.756.003-73", "registro": "336241CE"
    },
    "PALLOMA TEIXEIRA DA SILVA": {
        "empresa": "PALLOMA TEIXEIRA DA SILVA - PALLOMA TEIXEIRA ARQUITETURA LTDA",
        "cnpj": "54.862.474/0001-71", "cpf_emp": "064.943.593-10",
        "nome_resp": "PALLOMA TEIRA DA SILVA", "cpf_resp": "064.943.593-10", "registro": "A184355-9"
    },
    "SANDY PEREIRA CORDEIRO": {
        "empresa": "SANDY PEREIRA CORDEIRO - CS ENGENHARIA",
        "cnpj": "54.794.898/0001-46", "cpf_emp": "071.222.553-60",
        "nome_resp": "SANDY PEREIRA CORDEIRO", "cpf_resp": "071.222.553-60", "registro": "356882CE"
    },
    "TIAGO VICTOR DE SOUSA": {
        "empresa": "TIAGO VICTOR DE SOUSA - T V S ENGENHARIA E ASSESSORIA",
        "cnpj": "54.806.521/0001-60", "cpf_emp": "068.594.803-00",
        "nome_resp": "TIAGO VICTOR DE SOUSA", "cpf_resp": "068.594.803-00", "registro": "346856CE"
    }
}

# --- 3. FUNÇÕES DE SUPORTE ---
def to_f(v):
    try: 
        if v is None or v == "": return 0
        clean_v = str(v).replace('R$', '').replace('%', '').replace(' ', '')
        if ',' in clean_v and '.' in clean_v:
            clean_v = clean_v.replace('.', '').replace(',', '.')
        elif ',' in clean_v:
            clean_v = clean_v.replace(',', '.')
        clean_v = re.sub(r'[^\d.]', '', clean_v)
        return float(clean_v)
    except: return 0

def call_gemini(api_key, prompt):
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    for attempt in range(3):
        try:
            response = model.generate_content(prompt, generation_config=genai.types.GenerationConfig(response_mime_type="application/json", temperature=0.1))
            return json.loads(response.text)
        except:
            time.sleep(2)
    return None

def extrair_com_docling(doc, nome_doc):
    """
    IMPLEMENTAÇÃO DA SOLUÇÃO 'REAL': 
    Usa o PyPdfiumDocumentBackend para evitar carregar modelos de 1GB de RAM.
    """
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.datamodel.base_models import InputFormat
    # Importação do backend de baixo consumo
    from docling.backend.pypdfium2_backend import PyPdfiumDocumentBackend

    st.write(f"📂 Motor: Processando {nome_doc}...")
    
    pipeline_options = PdfPipelineOptions()
    
    # ESTRATÉGIA HÍBRIDA:
    # Laudo é digital e longo -> OCR desligado para não estourar RAM
    # PLS e Alvará precisam de precisão -> OCR ligado
    pipeline_options.do_ocr = (nome_doc != "LAUDO")
    pipeline_options.do_table_structure = (nome_doc == "PLS") # Tabelas só na PLS

    # A "BALA DE PRATA": Forçamos o conversor a usar o PyPdfium2
    # que é ordens de grandeza mais leve que o backend padrão.
    converter = DocumentConverter(
        allowed_formats=[InputFormat.PDF],
        format_options={
            InputFormat.PDF: PdfFormatOption(
                pipeline_options=pipeline_options,
                backend=PyPdfiumDocumentBackend # <--- A SOLUÇÃO REAL
            )
        }
    )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(doc.getbuffer())
        tmp_path = tmp.name

    try:
        res = converter.convert(tmp_path)
        markdown = res.document.export_to_markdown()
        
        # Limpeza agressiva
        del res
        del converter
        gc.collect()
        return markdown
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

def main():
    st.title("🏛️ Automação RAE CAIXA")
    st.markdown("##### v5.9 - Solução Definitiva: Backend PyPdfium2")

    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")
        st.divider()
        st.subheader("👤 Responsável Técnico")
        resp_selecionado = st.selectbox("Selecione o Profissional:", options=list(PROFISSIONAIS.keys()))
        st.divider()
        st.caption("Foco: Backend de Baixa Memória")

    st.subheader("📂 Documentação")
    col1, col2 = st.columns(2)
    with col1:
        pdf_laudo = st.file_uploader("1. Laudo Técnico (PDF)", type=["pdf"])
        pdf_pls = st.file_uploader("3. PLS (PDF)", type=["pdf"])
    with col2:
        excel_template = st.file_uploader("2. Modelo RAE (.xlsm)", type=["xlsm"])
        pdf_alvara = st.file_uploader("4. Alvará (PDF/Foto)", type=["pdf"])

    if st.button("🚀 INICIAR PROCESSAMENTO"):
        if not api_key or not pdf_laudo or not excel_template:
            st.warning("Preencha os campos obrigatórios.")
            return

        try:
            with st.status("Extraindo dados um por um...", expanded=True) as status:
                texto_total = ""
                
                # 1. LAUDO
                texto_total += f"\n--- DOCUMENTO: LAUDO ---\n{extrair_com_docling(pdf_laudo, 'LAUDO')}\n"
                gc.collect()
                
                # 2. PLS
                if pdf_pls:
                    texto_total += f"\n--- DOCUMENTO: PLS ---\n{extrair_com_docling(pdf_pls, 'PLS')}\n"
                    gc.collect()
                    
                # 3. ALVARÁ
                if pdf_alvara:
                    texto_total += f"\n--- DOCUMENTO: ALVARA ---\n{extrair_com_docling(pdf_alvara, 'ALVARA')}\n"
                    gc.collect()

                st.write("🧠 IA: Cruzando e Validando informações...")
                prompt = f"""
                Você é um engenheiro revisor da CAIXA. Analise os textos e gere um JSON.
                
                DADOS PRIORITÁRIOS:
                - valor_imovel: BUSQUE POR 'Avaliação Global', 'Valor de Mercado' ou 'Total do Imóvel'.
                - contratacao: Data da PLS (Célula AH63).
                - percentual_pls: 'Mensurado Acumulado Atual' da PLS (Célula W93).
                - acumulado_pls: Lista da coluna '% Acumulado' da PLS (AH72:AH108).
                - lat_s, long_w: Coordenadas GMS (ex: 06°24'08.8"). Remova letras.
                
                CONTEÚDO EXTRAÍDO:
                {texto_total}
                """
                
                dados = call_gemini(api_key, prompt)
                if not dados:
                    st.error("IA falhou. Memória insuficiente no servidor.")
                    return

                st.write("📊 Preenchendo Planilha Excel...")
                from openpyxl import load_workbook
                wb = load_workbook(BytesIO(excel_template.read()), keep_vba=True)
                
                if "Início Vistoria" in wb.sheetnames:
                    ws = wb["Início Vistoria"]
                    map_iv = {
                        "G43": "proponente", "AJ43": "cpf_cnpj", "AP43": "ddd", "AR43": "telefone",
                        "G49": "endereco", "AD49": "lat_s", "AH49": "long_w", "AL49": "complemento",
                        "G51": "bairro", "V51": "cep", "AA51": "municipio", "AS51": "uf_vistoria",
                        "AS53": "uf_registro", "G53": "valor_terreno", "Q53": "matricula",
                        "AA53": "oficio", "AJ53": "comarca"
                    }
                    for cell, key in map_iv.items():
                        val = dados.get(key, "")
                        ws[cell] = to_f(val) if key == "valor_terreno" else str(val).upper()
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                if "RAE" in wb.sheetnames:
                    ws_r = wb["RAE"]
                    ws_r["AH63"] = str(dados.get("contratacao", ""))
                    ws_r["AH66"] = to_f(dados.get("valor_imovel", 0))
                    ws_r["AS66"] = to_f(dados.get("etapas_original", 0))
                    ws_r["W93"] = to_f(dados.get("percentual_pls", 0))
                    ws_r["N95"] = "Sim" if pdf_alvara else "Não"
                    ws_r["M96"] = str(dados.get("alvara_emissao", ""))
                    ws_r["W96"] = str(dados.get("alvara_validade", ""))
                    ws_r["W102"] = str(dados.get("responsaveis_iguais", "Não")).capitalize()
                    
                    prof = PROFISSIONAIS[resp_selecionado]
                    ws_r["I315"], ws_r["I316"], ws_r["U316"] = prof["empresa"].upper(), prof["cnpj"], prof["cpf_emp"]
                    ws_r["AE315"], ws_r["AE316"], ws_r["AO316"] = prof["nome_resp"].upper(), prof["cpf_resp"], prof["registro"].upper()
                    
                    incs, acus_pls, acus_prop = dados.get("incidencias", []), dados.get("acumulado_pls", []), dados.get("acumulado", [])
                    for i in range(20): ws_r[f"S{69+i}"] = to_f(incs[i]) if i < len(incs) else 0
                    for i in range(len(acus_pls)): 
                        if i < 37: ws_r[f"AH{72+i}"] = to_f(acus_pls[i])
                    for i in range(len(acus_prop)):
                        if i < 37: ws_r[f"AE{72+i}"] = to_f(acus_prop[i])

                output = BytesIO()
                wb.save(output)
                status.update(label="✅ Concluído com sucesso!", state="complete", expanded=False)
                st.balloons()
                st.download_button(label="📥 BAIXAR RAE FINAL", data=output.getvalue(), file_name=f"RAE_FINAL.xlsm", mime="application/vnd.ms-excel.sheet.macroEnabled.12")

        except Exception as e:
            st.error(f"Erro: {e}")
            gc.collect()

if __name__ == "__main__":
    main()
