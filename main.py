import streamlit as st
import sys
import os
import re
import json
import time
import tempfile
import uuid
import gc
from io import BytesIO

# --- 1. CONFIGURAÇÃO INICIAL (OBRIGATORIAMENTE O PRIMEIRO COMANDO ST) ---
st.set_page_config(page_title="Automação RAE CAIXA", page_icon="🏛️", layout="centered")

# --- 2. PATCH DE METADADOS (EVITA CRASH NO BOOT) ---
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

# --- 3. BANCO DE DADOS DE PROFISSIONAIS ---
PROFISSIONAIS = {
    "FRANCISCO DAVID MENESES DOS SANTOS": {
        "empresa": "FRANCISCO DAVID MENESES DOS SANTOS - F. D. MENESES DOS SANTOS",
        "cnpj": "54.801.096/0001-16", "cpf_emp": "058.756.003-73",
        "nome_resp": "FRANCISCO DAVID MENESES DOS SANTOS", "cpf_resp": "058.756.003-73", "registro": "336241CE"
    },
    "PALLOMA TEIXEIRA DA SILVA": {
        "empresa": "PALLOMA TEIXEIRA DA SILVA - PALLOMA TEIXEIRA ARQUITETURA LTDA",
        "cnpj": "54.862.474/0001-71", "cpf_emp": "064.943.593-10",
        "nome_resp": "PALLOMA TEIXEIRA DA SILVA", "cpf_resp": "064.943.593-10", "registro": "A184355-9"
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

# --- 4. IMPORTAÇÕES ---
try:
    import pandas as pd
    from openpyxl import load_workbook
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.datamodel.base_models import InputFormat
    import google.generativeai as genai
    import onnxruntime
    import optree
    DEPENDENCIAS_OK = True
except Exception as e:
    DEPENDENCIAS_OK = False
    ERRO_IMPORT = str(e)

# --- 5. LÓGICA DE IA ---
def get_converter_instance():
    """Cria uma instância do conversor. Não usamos cache global para evitar acúmulo de RAM."""
    pipeline_options = PdfPipelineOptions()
    pipeline_options.do_table_structure = True 
    pipeline_options.table_structure_options.do_cell_matching = True
    return DocumentConverter(
        allowed_formats=[InputFormat.PDF],
        format_options={
            InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)
        }
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
    return None

def main():
    st.title("🏛️ Automação RAE CAIXA")
    st.markdown("##### Multidocumentos: Laudo + PLS + Alvará (OCR)")

    if not DEPENDENCIAS_OK:
        st.error(f"Erro Crítico: {ERRO_IMPORT}")
        return

    with st.sidebar:
        st.header("⚙️ Configurações")
        api_key = st.text_input("Gemini API Key:", type="password")
        st.divider()
        st.subheader("👤 Responsável Técnico")
        resp_selecionado = st.selectbox("Selecione o Profissional:", options=list(PROFISSIONAIS.keys()))
        st.divider()
        st.caption("v3.8 - Otimização de RAM (Sequential)")

    st.subheader("📂 Documentação")
    col1, col2 = st.columns(2)
    with col1:
        pdf_laudo = st.file_uploader("1. Laudo Técnico (PDF)", type=["pdf"])
        pdf_pls = st.file_uploader("3. PLS (PDF)", type=["pdf"])
    with col2:
        excel_template = st.file_uploader("2. Modelo RAE (.xlsm)", type=["xlsm"])
        pdf_alvara = st.file_uploader("4. Alvará (PDF/Foto)", type=["pdf"])

    if st.button("🚀 PROCESSAR TUDO"):
        if not api_key or not pdf_laudo or not excel_template:
            st.warning("Preencha a chave, o laudo e a planilha.")
            return

        try:
            with st.status("Extraindo dados com segurança de memória...", expanded=True) as status:
                texto_total = ""

                # Processamento Sequencial Estrito para economizar RAM
                documentos_para_processar = [
                    ("LAUDO", pdf_laudo),
                    ("PLS", pdf_pls),
                    ("ALVARA", pdf_alvara)
                ]

                for nome, doc in documentos_para_processar:
                    if doc:
                        st.write(f"📖 Lendo {nome} (Limpando memória antes)...")
                        # Força limpeza de memória antes de cada conversão
                        gc.collect() 
                        
                        # Criamos o conversor apenas para este arquivo e deletamos depois
                        converter = get_converter_instance()
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                            tmp.write(doc.getbuffer())
                            tmp_path = tmp.name
                        try:
                            res = converter.convert(tmp_path)
                            texto_total += f"\n--- INÍCIO: {nome} ---\n{res.document.export_to_markdown()}\n"
                            
                            # Limpeza imediata de objetos pesados
                            del res
                            del converter
                            gc.collect() 
                        finally:
                            if os.path.exists(tmp_path): os.remove(tmp_path)

                st.write("🧠 IA: Cruzando informações de todos os documentos...")
                prompt = f"""
                Você é um engenheiro revisor da CAIXA. Analise os documentos e extraia estritamente para este JSON.
                
                DADOS GERAIS: 
                - proponente, cpf_cnpj, ddd, telefone, endereco, bairro, cep, municipio, uf_vistoria, uf_registro, complemento, matricula, comarca, valor_terreno, valor_imovel, lat_s, long_w, etapas_original, oficio
                
                REGRAS ESPECÍFICAS:
                1. valor_imovel: É OBRIGATÓRIO. Procure por 'Valor de Mercado', 'Valor Global', 'Avaliação Final' ou 'Total do Imóvel'.
                2. contratacao: Data de contratação na PLS (AH63).
                3. percentual_pls: 'Mensurado Acumulado Atual' da PLS (W93).
                4. acumulado_pls: Lista da coluna '% Acumulado' da PLS (AH72:AH108).
                5. alvara: Data emissão (M96) e validade (W96). Marque responsaveis_iguais como 'Sim' se o RT da PLS for o mesmo do Alvará.
                6. Coordenadas: Apenas números e símbolos de graus (ex: 06°24'08.8"). Remova letras S, N, W, E.
                
                CONTEÚDO DOS DOCUMENTOS:
                {texto_total}
                """
                
                dados = call_gemini(api_key, prompt)
                if not dados:
                    st.error("A IA não conseguiu processar os dados. Verifique a chave API ou tente novamente.")
                    return

                st.write("📊 Gravando informações na planilha...")
                wb = load_workbook(BytesIO(excel_template.read()), keep_vba=True)
                wb.calculation.fullCalcOnLoad = True
                
                def to_f(v):
                    try: 
                        if v is None or v == "": return 0
                        return float(str(v).replace('.', '').replace(',', '.').replace('%', '').strip())
                    except: return 0

                # Aba Início Vistoria
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
                        if key == "valor_terreno":
                            ws[cell] = to_f(val)
                        else:
                            ws[cell] = str(val).upper() if val else ""
                    ws["Q54"], ws["Q55"], ws["Q56"] = "Casa", "Residencial", "Vistoria para aferição de obra"

                # Aba RAE
                if "RAE" in wb.sheetnames:
                    ws_rae = wb["RAE"]
                    ws_rae.sheet_state = 'visible'
                    
                    # Preenchimento das células solicitadas
                    ws_rae["AH63"] = str(dados.get("contratacao", ""))
                    ws_rae["AH66"] = to_f(dados.get("valor_imovel", 0))
                    ws_rae["AS66"] = to_f(dados.get("etapas_original", 0))
                    ws_rae["W93"] = to_f(dados.get("percentual_pls", 0))
                    
                    ws_rae["N95"] = "Sim" if pdf_alvara else "Não"
                    ws_rae["M96"] = str(dados.get("alvara_emissao", ""))
                    ws_rae["W96"] = str(dados.get("alvara_validade", ""))
                    ws_rae["W102"] = str(dados.get("responsaveis_iguais", "Não")).capitalize()
                    
                    # Dados do Profissional selecionado
                    prof = PROFISSIONAIS[resp_selecionado]
                    ws_rae["I315"] = prof["empresa"].upper()
                    ws_rae["I316"] = prof["cnpj"]
                    ws_rae["U316"] = prof["cpf_emp"]
                    ws_rae["AE315"] = prof["nome_resp"].upper()
                    ws_rae["AE316"] = prof["cpf_resp"]
                    ws_rae["AO316"] = prof["registro"].upper()
                    
                    # Tabelas e Cronogramas
                    incs, acus_pls, acus_prop = dados.get("incidencias", []), dados.get("acumulado_pls", []), dados.get("acumulado", [])
                    
                    for i in range(20): 
                        ws_rae[f"S{69+i}"] = to_f(incs[i]) if i < len(incs) else 0
                    
                    for i in range(len(acus_pls)):
                        if i < 37: ws_rae[f"AH{72+i}"] = to_f(acus_pls[i])
                    
                    for i in range(len(acus_prop)):
                        if i < 37: ws_rae[f"AE{72+i}"] = to_f(acus_prop[i])

                output = BytesIO()
                wb.save(output)
                status.update(label="✅ Tudo pronto!", state="complete", expanded=False)
                st.balloons()
                
                # Nome do arquivo de saída
                proponente_nome = str(dados.get("proponente", "FINAL")).split()[0].upper()
                st.download_button(label=f"📥 BAIXAR RAE - {proponente_nome}", data=output.getvalue(), file_name=f"RAE_{proponente_nome}.xlsm", mime="application/vnd.ms-excel.sheet.macroEnabled.12")

        except Exception as e:
            st.error(f"Erro técnico: {e}")
            st.info("💡 Se o erro for 'Out of Memory', tente processar apenas o Laudo e a PLS primeiro.")

if __name__ == "__main__":
    main()
