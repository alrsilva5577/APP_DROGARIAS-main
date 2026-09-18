import streamlit as st
from datetime import datetime
from docx import Document
import pandas as pd
import io

# 1. Configurações iniciais da página
st.set_page_config(page_title="Módulo de Inspeção", layout="wide")
st.title("💊 Sistema de Gestão de Inspeções - DROGARIAS")

# --- LEITURA DIRETA DO ARQUIVO EXCEL DE LEIS (LOCAL) ---
@st.cache_data 
def carregar_dados_excel():
    try:
        df = pd.read_excel("INFRACOES_DB.xlsx")
        return df
    except FileNotFoundError:
        st.error("❌ Erro: O arquivo 'INFRACOES_DB.xlsx' não foi encontrado na raiz do projeto do GitHub.")
        return None

df_excel = carregar_dados_excel()

# --- CONEXÃO COM O BANCO DE DADOS EM NUVEM (GOOGLE SHEETS) ---
# O Streamlit busca o link configurado de forma segura no arquivo secrets.toml
try:
    conn_nuvem = st.connection("gsheets", type="connections.gsheets")
except Exception:
    conn_nuvem = None

# --- CONTROLADORES DE TELA (SESSION STATE) ---
if "tela" not in st.session_state:
    st.session_state["tela"] = "fluxo_inicial"

if "dados_processo" not in st.session_state:
    st.session_state["dados_processo"] = {}

if "lista_pessoas" not in st.session_state:
    st.session_state["lista_pessoas"] = [{"nome": "", "cpf": "", "cargo": "", "email": ""}]

if "mostrar_observacao_rt" not in st.session_state:
    st.session_state["mostrar_observacao_rt"] = False


# ==============================================================================
# TELA 1: FLUXO INICIAL
# ==============================================================================
if st.session_state["tela"] == "fluxo_inicial":
    st.subheader("1. Identificação do Procedimento")
    
    tipo_licenca = st.radio(
        "Tipo de procedimento regulatório:",
        ["Renovação de Licença Sanitária", "Licença Sanitária Inicial"],
        horizontal=True
    )
    
    st.divider()
    
    if tipo_licenca == "Licença Sanitária Inicial":
        st.markdown("#### 🆕 Cadastro do Estabelecimento Novo")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Razão Social**")
            rz = st.text_input("Razão Social", key="ini_rz", label_visibility="collapsed")
            st.markdown("**CNPJ (apenas números)**")
            cnpj = st.text_input("CNPJ", key="ini_cnpj", max_chars=14, label_visibility="collapsed")
            st.markdown("**Número da WEB**")
            web = st.text_input("WEB", key="ini_web", label_visibility="collapsed")
        with col2:
            st.markdown("**Endereço (Rua e Número)**")
            end = st.text_input("Endereço", key="ini_end", label_visibility="collapsed")
            st.markdown("**Bairro**")
            bairro = st.text_input("Bairro", key="ini_bai", label_visibility="collapsed")
            st.markdown("**CEP**")
            cep = st.text_input("CEP", key="ini_cep", label_visibility="collapsed")
            st.markdown("**CNAE Fiscal**")
            cnae = st.text_input("CNAE", value="4771-7/01", key="ini_cnae", label_visibility="collapsed")
            
        if st.button("Avançar para Opções", type="primary", key="btn_ini_avancar"):
            if cnpj and web and rz:
                # No piloto em nuvem simplificado, avançamos guardando os dados na sessão
                st.session_state["dados_processo"] = {"cnpj": cnpj, "web": web}
                st.session_state["tela"] = "menu_opcoes"
                st.rerun()
            else:
                st.error("❌ Preencha os campos obrigatórios: Razão Social, CNPJ e Número da WEB.")

    else:
        st.markdown("#### 🔄 Renovação de Licença Existente")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**CNPJ do Estabelecimento (apenas números)**")
            cnpj = st.text_input("CNPJ Renovação", key="ren_cnpj", max_chars=14, label_visibility="collapsed")
        with col2:
            st.markdown("**Número da WEB**")
            web = st.text_input("WEB Renovação", key="ren_web", label_visibility="collapsed")
            
        if st.button("Avançar para Opções", type="primary", key="btn_ren_avancar"):
            if cnpj and web:
                st.session_state["dados_processo"] = {"cnpj": cnpj, "web": web}
                st.session_state["tela"] = "menu_opcoes"
                st.rerun()
            else:
                st.error("❌ Preencha os campos CNPJ e Número da WEB.")


# ==============================================================================
# TELA 2: MENU DE OPÇÕES (IS, AD, ADEQUAÇÕES)
# ==============================================================================
elif st.session_state["tela"] == "menu_opcoes":
    cnpj_ativo = st.session_state["dados_processo"]["cnpj"]
    web_ativa = st.session_state["dados_processo"]["web"]
    
    st.subheader(f"📂 Painel de Ações | CNPJ: {cnpj_ativo} | WEB: {web_ativa}")
    
    if st.button("⬅️ Alterar Empresa / Voltar"):
        st.session_state["tela"] = "fluxo_inicial"
        st.rerun()
        
    st.divider()
    
    col_menu1, col_menu2 = st.columns(2)
    
    with col_menu1:
        st.markdown("### 📋 Atos de Fiscalização")
        if st.button("🔍 Adicionar Inspeção Sanitária (IS)", use_container_width=True, type="primary"):
            st.session_state["tela"] = "inspecao_sanitaria"
            st.rerun()
            
        if st.button("📂 Adicionar Avaliação Documental (AD)", use_container_width=True, disabled=True):
            pass
            
    with col_menu2:
        st.markdown("### 📄 Relatórios e Documentos Oficiais")
        
        # --- BUSCA INTEGRADA EM NUVEM PARA LISTA DE ADEQUAÇÕES ---
        if st.button("📋 Gerar Lista de Adequações Pendentes", use_container_width=True):
            if conn_nuvem is not None:
                try:
                    # Lê a tabela do histórico direto da Planilha Google online
                    df_historico_completo = conn_nuvem.read(worksheet="INSPECOES_DB", ttl="5s")
                    
                    # Filtra apenas as linhas referentes a esta WEB específica
                    df_ult_inspecao = df_historico_completo[df_historico_completo["id_renovacao"].astype(str) == str(web_ativa)]
                    
                    if df_ult_inspecao.empty:
                        st.warning("⚠️ Nenhuma inspeção anterior registrada na nuvem para esta solicitação WEB.")
                    elif df_excel is None:
                        st.error("Planilha de Legislações ausente.")
                    else:
                        # Pega a última linha inserida daquela WEB
                        chaves_texto = str(df_ult_inspecao.iloc[-1]["lista_adequacoes_chaves"])
                        
                        if chaves_texto.strip() == "" or chaves_texto == "nan":
                            st.success("🟢 Excelente! A última inspeção não registrou nenhuma inadequação pendente.")
                        else:
                            lista_chaves = chaves_texto.split(",")
                            
                            doc_exigencias = Document()
                            doc_exigencias.add_heading("7- Orientações / Providências (Lista de Adequações Pendentes)", level=1)
                            doc_exigencias.add_paragraph("O responsável técnico ou quem este designar deverá adotar as seguintes providências:")
                            
                            item_numero = 1
                            for chave in lista_chaves:
                                linha_infra = df_excel[df_excel["id_lei_artigo"] == chave.strip()]
                                if not linha_infra.empty:
                                    frase_exigencia = str(linha_infra["frase_exigencia"].values[0])
                                    doc_exigencias.add_paragraph(f"7.{item_numero}. {frase_exigencia}")
                                    item_numero += 1
                            
                            buffer_exig = io.BytesIO()
                            doc_exigencias.save(buffer_exig)
                            buffer_exig.seek(0)
                            
                            st.download_button(
                                label="💾 Baixar Lista de Adequações (.docx)",
                                data=buffer_exig,
                                file_name=f"Lista_Adequacoes_Pendentes_{web_ativa}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            st.success("Lista numerada extraída com sucesso direto da nuvem!")
                except Exception as e:
                    st.error(f"Erro ao acessar o banco em nuvem: {e}")
            else:
                st.error("Conexão com a nuvem não configurada no secrets.toml.")


# ==============================================================================
# TELA 3: O FORMULÁRIO DE INSPEÇÃO SANITÁRIA (IS)
# ==============================================================================
elif st.session_state["tela"] == "inspecao_sanitaria":
    cnpj_estabelecimento = st.session_state["dados_processo"]["cnpj"]
    n_web = st.session_state["dados_processo"]["web"]
    data_atual = datetime.now().strftime("%d/%m/%Y")
    
    st.subheader(f"📝 Formulário de Inspeção Sanitária (IS) | WEB: {n_web}")
    
    if st.button("⬅️ Cancelar e Voltar ao Painel"):
        st.session_state["tela"] = "menu_opcoes"
        st.rerun()
        
    st.divider()
    
    tab_contato, tab_proxima = st.tabs(["👥 Contato", "⚙️ Próximas Etapas..."])
    
    with tab_contato:
        st.subheader("Informações de Contato e Equipe")
        st.markdown("#### Acompanhantes da Inspeção")
        
        for indice, p in enumerate(st.session_state["lista_pessoas"]):
            st.markdown(f"**Pessoa {indice + 1}**")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown("**Nome**")
            st.session_state["lista_pessoas"][indice]["nome"] = st.text_input("Nome", value=p["nome"], key=f"nome_{indice}", label_visibility="collapsed")
            
        with col2:
            st.markdown("CPF")
            st.session_state["lista_pessoas"][indice]["cpf"] = st.text_input("CPF", value=p["cpf"], key=f"cpf_{indice}", label_visibility="collapsed")
            
        with col3:
            st.markdown("Cargo")
            st.session_state["lista_pessoas"][indice]["cargo"] = st.text_input("Cargo", value=p["cargo"], key=f"cargo_{indice}", label_visibility="collapsed")
            
        with col4:
            st.markdown("E-mail")
            st.session_state["lista_pessoas"][indice]["email"] = st.text_input("E-mail", value=p["email"], key=f"email_{indice}", label_visibility="collapsed")
            
    if st.button("➕ Adicionar Pessoa"):
        st.session_state["lista_pessoas"].append({"nome": "", "cpf": "", "cargo": "", "email": ""})
        st.rerun()
        
    st.divider()
    st.markdown("#### Informações")
    st.markdown("Número de colaboradores:")
    num_colaboradores = st.number_input("Número de colaboradores:", min_value=0, step=1, value=0, label_visibility="collapsed")
    
    st.divider()
    col_rt, col_botao = st.columns([3, 1])
    
    with col_rt:
        st.markdown("Farmacêutico presente no ato da inspeção?")
        status_rt = st.radio("RT Status", ["C", "NC", "NA"], horizontal=True, label_visibility="collapsed")
        
    with col_botao:
        st.write("")
        if st.button("➕ Observação"):
            st.session_state["mostrar_observacao_rt"] = True
            st.rerun()
            
    obs_rt = ""
    if st.session_state["mostrar_observacao_rt"]:
        st.markdown("Digite as observações sobre a situação do farmacêutico:")
        obs_rt = st.text_area("Obs RT Texto", placeholder="Ex: Farmacêutico ausente...", label_visibility="collapsed")
        
    st.divider()
    st.markdown("### 💾 Finalizar Relatório")
    
    if st.button("📄 Gerar e Baixar Documento Word", type="primary"):
        if df_excel is None:
            st.error("Não é possível gerar o relatório sem carregar a planilha 'INFRACOES_DB.xlsx'.")
        else:
            doc = Document()
            cabecalho = f"Estabelecimento inspecionado em {data_atual} em atendimento a solicitação web nº {n_web} que trata da emissão de Licença Sanitária."
            doc.add_paragraph(cabecalho)
            
            texto_contatos = "A inspeção foi acompanhada por: "
            lista_contatos_formatados = []
            for p in st.session_state["lista_pessoas"]:
                if p["nome"].strip() != "":
                    info_pess = f"{p['nome']}, CPF: {p['cpf']}, cargo: {p['cargo']}, email: {p['email']}"
                    lista_contatos_formatados.append(info_pess)
                    
            if lista_contatos_formatados:
                texto_contatos += "; ".join(lista_contatos_formatados) + "."
                doc.add_paragraph(texto_contatos)
                
            frase_rt_final = ""
            linha_rt = df_excel[df_excel["id_lei_artigo"] == "RDC44_3"]
            lista_adequacoes_chaves = []
            
            if not linha_rt.empty:
                if status_rt == "C":
                    frase_rt_final = str(linha_rt["frase_conforme"].values[0])
                elif status_rt == "NC":
                    frase_rt_final = str(linha_rt["frase_nao_conforme"].values[0])
                    lista_adequacoes_chaves.append("RDC44_3")
                    
            if obs_rt.strip() != "":
                frase_rt_final += f" {obs_rt}" if frase_rt_final else obs_rt
                
            if frase_rt_final:
                doc.add_paragraph(frase_rt_final)
                
            doc.add_paragraph("")
            doc.add_heading("1- Informações gerais:", level=1)
            info_gerais = f"Estabelecimento que desenvolve atividade enquadrada no Agrupamento 28 – Comércio Varejista de Medicamentos – CNAE Fiscal 4771-7/01 – Comércio Varejista de Produtos Farmacêuticos sem Manipulação de Fórmulas.\nNúmero de colaboradores: {num_colaboradores}."
            doc.add_paragraph(info_gerais)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            # --- GRAVAÇÃO AUTOMÁTICA EM NUVEM (GOOGLE SHEETS) ---
            if conn_nuvem is not None:
                timestamp_chave = datetime.now().strftime("%Y%m%d_%H%M%S")
                chave_primaria_inspecao = f"{n_web}IS{timestamp_chave}"
                texto_pendencias = ",".join(lista_adequacoes_chaves)
                
                nova_linha = pd.DataFrame([{
                    "id_inspecao": chave_primaria_inspecao,
                    "id_renovacao": n_web,
                    "cnpj_estabelecimento": cnpj_estabelecimento,
                    "tipo_acao": "IS",
                    "data_procedimento": data_atual,
                    "lista_adequacoes_chaves": texto_pendencias
                }])
                
                try:
                    # Adiciona a linha direto no Google Sheets online de forma transparente
                    conn_nuvem.create(worksheet="INSPECOES_DB", data=nova_linha)
                    st.toast("💾 Dados sincronizados na nuvem com sucesso!", icon="☁️")
                except Exception as e:
                    st.warning(f"⚠️ Relatório gerado localmente, mas houve uma falha de sincronismo com a nuvem: {e}")
                    
            st.download_button(
                label="💾 Clique aqui para salvar o arquivo .docx",
                data=buffer,
                file_name=f"Relatorio_Inspecao_{n_web}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("✨ Relatório gerado com sucesso!")
