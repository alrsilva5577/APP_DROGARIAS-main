import streamlit as st
from datetime import datetime
from docx import Document
import pandas as pd
import requests
import io
from sqlalchemy import text
import json

# 1. Configurações iniciais da página
st.set_page_config(page_title="Módulo de Inspeção", layout="wide")
st.title("💊 Sistema de Gestão de Inspeções - DROGARIAS")

# --- LEITURA DOS ARQUIVOS EXCEL LOCAIS (LEIS E FISCAIS) ---
@st.cache_data 
def carregar_dados_locais():
    try:
        df_inf = pd.read_excel("INFRACOES_DB.xlsx")
    # esse banco contém info sobre as chaves de infrações e os textos a serem adicionados no relatório
    except FileNotFoundError:
        st.error("❌ Erro: O arquivo 'INFRACOES_DB.xlsx' não foi encontrado.")
        df_inf = None
        
    try:
        df_fisc = pd.read_excel("FISCAIS_DB.xlsx")
    # esse banco contém info sobre os fiscais e seus apelidos para seleção no app
    except FileNotFoundError:
        st.error("❌ Erro: O arquivo 'FISCAIS_DB.xlsx' não foi encontrado.")
        df_fisc = None
        
    return df_inf, df_fisc

df_excel, df_fiscais = carregar_dados_locais()

# --- CONEXÃO COM O BANCO DE DADOS POSTGRESQL (SUPABASE) ---
conn_nuvem = None
try:
    # O Streamlit se conecta nativamente usando a chave [connections.sql] do secrets
    conn_nuvem = st.connection("postgresql", type="sql")
except Exception as e_conexao:
    st.sidebar.error(f"⚠️ Falha na conexão com o banco PostgreSQL: {e_conexao}")


# --- CONTROLADORES DE TELA E HISTÓRICO DE RESPOSTAS ---
if "tela" not in st.session_state:
    st.session_state["tela"] = "fluxo_inicial"

if "dados_processo" not in st.session_state:
    st.session_state["dados_processo"] = {}

if "lista_pessoas" not in st.session_state:
    st.session_state["lista_pessoas"] = [{"nome": "", "cpf": "", "cargo": "", "email": ""}]

if "mostrar_observacao_rt" not in st.session_state:
    st.session_state["mostrar_observacao_rt"] = False

# GAVETA MÁGICA: Guarda as respostas anteriores trazidas do banco de dados
if "historico_pre_preenchido" not in st.session_state:
    st.session_state["historico_pre_preenchido"] = {}


# --- FUNÇÃO AUXILIAR: BUSCA HISTÓRICO NO POSTGRESQL ---
def carregar_historico_cnpj(cnpj_limpo):
    """
    Busca o histórico mestre da última inspeção no Supabase trazendo TODAS 
    as colunas de uma vez e prepara a memória da sessão.
    """
    # Reseta o estado para o padrão de segurança (caso seja uma empresa nova)
    st.session_state["historico_pre_preenchido"] = {}
    
    if conn_nuvem is None:
        return False
        
    try:
        with conn_nuvem.session as session:
            # O 'SELECT *' traz todas as informações
            resultado = session.execute(
                text("""
                SELECT * 
                FROM inspecoes_db 
                WHERE cnpj = :cnpj 
                ORDER BY data_procedimento DESC, id_inspecao DESC 
                LIMIT 1;
                """),
                {"cnpj": cnpj_limpo}
            ).mappings().fetchone()
            
            if resultado:
                # Salva todas as colunas do banco diretamente no dicionário
                dados_linha = dict(resultado)
                
                # Converte a string JSON do banco de volta para dicionário Python
                campo_json = dados_linha.get("locais_controle_geladeiras_json")
                if isinstance(campo_json, str) and campo_json.strip():
                    try:
                        dados_linha["locais_controle_geladeiras_json"] = json.loads(campo_json)
                    except Exception:
                        dados_linha["locais_controle_geladeiras_json"] = {}
                # ─────────────────────────────────────────────────────────────────────────────
                # Realiza a leitura reversa de segurança para os campos JSONB da Tab 4
                for col_j in ["medicamentos_controlados_verificados_json", "antimicrobianos_verificados_json"]:
                    c_json = dados_linha.get(col_j)
                    if isinstance(c_json, str) and c_json.strip():
                        try:
                            dados_linha[col_j] = json.loads(c_json)
                        except Exception:
                            dados_linha[col_j] = []


                st.session_state["historico_pre_preenchido"] = dados_linha
                return True
                
    except Exception as e:
        print(f"Erro ao carregar histórico completo do Supabase: {e}")
        
    return False


# ==============================================================================
# TELA 1: FLUXO INICIAL
# ==============================================================================
if st.session_state["tela"] == "fluxo_inicial":
    st.subheader("1. Identificação da Inspeção")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Número da Solicitação WEB**")
        web = st.text_input("WEB Entrada", key="init_web", label_visibility="collapsed")
        
        st.markdown("**Fiscal Responsável**")
        # Puxa dinamicamente a lista de apelidos do arquivo Excel FISCAIS_DB
        lista_apelidos = ["Selecione o fiscal..."]
        if df_fiscais is not None and "apelido" in df_fiscais.columns:
            lista_apelidos.extend(df_fiscais["apelido"].dropna().tolist())
            
        fiscal_selecionado = st.selectbox("Fiscal Entrada", options=lista_apelidos, key="init_fiscal", label_visibility="collapsed")

    with col2:
        st.markdown("**CNPJ do Estabelecimento (apenas números)**")
        cnpj_digitado = st.text_input("CNPJ Entrada", key="init_cnpj_text", max_chars=14, label_visibility="collapsed", placeholder="Digite os 14 números")
        
        # --- MOTOR DE AUTOCOMPLETAR CNPJ EM TEMPO REAL ---
        cnpj_limpo = "".join(filter(str.isdigit, cnpj_digitado))
        if len(cnpj_limpo) >= 3 and conn_nuvem is not None:
            try:
                # Passamos a string pura de texto. O Streamlit gerencia o cache e o hash perfeitamente.
                query_sugestao = "SELECT cnpj, razao_social FROM empresas_db WHERE cnpj LIKE :termo LIMIT 5;"
                df_sug = conn_nuvem.query(query_sugestao, params={"termo": f"{cnpj_limpo}%"}, ttl=0)
                
                if not df_sug.empty:
                    st.markdown("*Drogarias encontradas na base:*")
                    for idx, row in df_sug.iterrows():
                        st.caption(f"🔹 **{row['cnpj']}** - {row['razao_social']}")
            except Exception as e_sug:
                # Imprime discretamente no terminal caso haja falha de sintaxe
                print(f"Erro na sugestão: {e_sug}")

    st.divider()

        # --- BOTÃO PRINCIPAL DE AVANÇO COM INTELIGÊNCIA ARTIFICIAL DE CADASTRO ---
    if st.button("Avançar para o Painel de Ações", type="primary", use_container_width=True):
        if not web or fiscal_selecionado == "Selecione o fiscal..." or len(cnpj_limpo) != 14:
            st.error("❌ Preencha todos os campos obrigatórios corretamente. O CNPJ precisa ter exatamente 14 dígitos.")
        else:
            # Resgata os dados completos do Fiscal selecionado com base no apelido
            info_fiscal = {"apelido": fiscal_selecionado, "nome": fiscal_selecionado, "cargo": "Fiscal", "credencial": "000"}
            if df_fiscais is not None:
                linha_f = df_fiscais[df_fiscais["apelido"] == fiscal_selecionado]
                if not linha_f.empty:
                    info_fiscal = {
                        "apelido": str(linha_f.iloc[0]["apelido"]),
                        "nome": str(linha_f.iloc[0]["nome"]),
                        "cargo": str(linha_f.iloc[0]["cargo"]),
                        "credencial": str(linha_f.iloc[0]["credencial"])
                    }
            st.session_state["fiscal_ativo"] = info_fiscal
            
            # --- VERIFICAÇÃO SE O CNPJ JÁ EXISTE NO SUPABASE ---
            empresa_localizada = False
            if conn_nuvem is not None:
                try:
                    query_checagem = "SELECT * FROM empresas_db WHERE cnpj = :cnpj LIMIT 1;"
                    df_emp = conn_nuvem.query(query_checagem, params={"cnpj": cnpj_limpo}, ttl=0)
                    
                    if not df_emp.empty:
                        st.session_state["empresa_encontrada"] = {
                            "razao_social": df_emp.iloc[0]["razao_social"],
                            "endereco": df_emp.iloc[0]["endereco"],
                            "bairro": df_emp.iloc[0]["bairro"],
                            "cep": df_emp.iloc[0]["cep"],
                            "cnae": df_emp.iloc[0]["cnae"],
                            "atividade": df_emp.iloc[0]["atividade"]
                        }
                        empresa_localizada = True
                except Exception as e:
                    st.error(f"Erro ao consultar base de dados: {e}")
            
            # --- AUTO-CADASTRO: SE NÃO EXISTIR, CONSULTA A BRASILAPI NA HORA ---
            if not empresa_localizada:
                with st.spinner("🕵️ CNPJ inédito! Consultando base nacional da Receita Federal..."):
                    try:
                        url_api = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}"
                        resposta = requests.get(url_api, timeout=10)
                        
                        if resposta.status_code == 200:
                            dados = resposta.json()
                            rz_social = dados.get("razao_social") or dados.get("nome_fantasia") or "Estabelecimento Novo"
                            logr = dados.get("logradouro") or ""
                            num = dados.get("numero") or ""
                            compl = dados.get("complemento") or ""
                            end_completo = f"{logr}, Nº {num}" + (f" - {compl}" if compl and compl != "NONE" else "")
                            bairro_api = dados.get("bairro") or "Não Informado"
                            cep_api = dados.get("cep") or "Não Informado"
                            cnae_api = str(dados.get("cnae_fiscal") or "4771-7/01")
                            ativ_api = dados.get("cnae_fiscal_descricao") or "Comércio Varejista Farmacêutico"
                            
                            # Salva na memória do aplicativo
                            st.session_state["empresa_encontrada"] = {
                                "razao_social": rz_social, "endereco": end_completo, "bairro": bairro_api,
                                "cep": cep_api, "cnae": cnae_api, "atividade": ativ_api
                            }
                            
                            # Grava automaticamente na tabela empresas_db do Supabase
                            if conn_nuvem is not None:
                                with conn_nuvem.session as session:
                                    session.execute(text("""
                                        INSERT INTO empresas_db (cnpj, razao_social, endereco, bairro, cep, cnae, atividade)
                                        VALUES (:cnpj, :razao_social, :endereco, :bairro, :cep, :cnae, :atividade);
                                    """), {"cnpj": cnpj_limpo, "razao_social": rz_social, "endereco": end_completo, "bairro": bairro_api, "cep": cep_api, "cnae": cnae_api, "atividade": ativ_api})
                                    session.commit()
                                st.toast(f"🆕 {rz_social} cadastrada na base com sucesso!", icon="🏢")
                                empresa_localizada = True
                            else:
                                st.error(f"❌ CNPJ não encontrado na Receita Federal (Erro API: {resposta.status_code}).")

                    except Exception as e_api:
                        st.error(f"⚠️ Falha de comunicação com a base cadastral: {e_api}")
            
            # Se a empresa foi achada ou cadastrada com sucesso, carrega o histórico e avança!
            if empresa_localizada:
                carregar_historico_cnpj(cnpj_limpo)
                st.session_state["dados_processo"] = {"cnpj": cnpj_limpo, "web": web}
                st.session_state["tela"] = "menu_opcoes"
                st.rerun()

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
        if st.button("🔍 Adicionar Inspeção Sanitária", use_container_width=True, type="primary"):
            st.session_state["tela"] = "inspecao_sanitaria"
            st.rerun()
        if st.button("📂 Adicionar Avaliação Documental", use_container_width=True, disabled=True):
            pass
            
    with col_menu2:
        st.markdown("### 📄 Relatórios e Documentos Oficiais")
        
        # 1. BOTÃO: Lista de Adequações Pendentes (Adaptado para ler do banco Supabase)
        if st.button("📋 Gerar Lista de Adequações Pendentes", use_container_width=True):
            if conn_nuvem is not None:
                try:
                    with conn_nuvem.session as session:
                        # Busca o histórico da inspeção atual baseado no número WEB ativo
                        query_adequacoes = """
                            SELECT status_inspecao_itens 
                            FROM inspecoes_db 
                            WHERE web = :web 
                            ORDER BY data_procedimento DESC, id_inspecao DESC 
                            LIMIT 1;
                        """
                        resultado = session.execute(text(query_adequacoes), {"web": web_ativa}).fetchone()
                        
                        if not resultado:
                            st.warning("⚠️ Nenhuma inspeção anterior registrada na nuvem para esta solicitação WEB.")
                        elif df_excel is None:
                            st.error("Planilha de Legislações 'INFRACOES_DB.xlsx' ausente.")
                        else:
                            # Filtra quais itens foram salvos como Não Conforme ("NC")
                            string_status = resultado[0] if resultado[0] else ""
                            lista_inadequados = []
                            
                            if ":" in string_status:
                                pares = string_status.split(",")
                                for par in pares:
                                    if ":" in par:
                                        chave, valor = par.split(":")
                                        if valor.strip() == "NC":
                                            lista_inadequados.append(chave.strip())
                            
                            if not lista_inadequados:
                                st.success("🟢 Excelente! A última inspeção não registrou nenhuma inadequação pendente.")
                            else:
                                doc_exigencias = Document()
                                doc_exigencias.add_heading("7- Orientações / Providências", level=1)
                                item_numero = 1
                                
                                for chave in lista_inadequados:
                                    linha_infra = df_excel[df_excel["id_lei_artigo"] == chave]
                                    if not linha_infra.empty:
                                        # Substitua 'frase_exigencia' pelo nome exato da coluna de providência no seu Excel
                                        frase_exigencia = str(linha_infra["frase_exigencia"].values[0])
                                        doc_exigencias.add_paragraph(f"7.{item_numero}. {frase_exigencia}")
                                        item_numero += 1
                                        
                                buffer_exig = io.BytesIO()
                                doc_exigencias.save(buffer_exig)
                                buffer_exig.seek(0)
                                st.download_button(
                                    label="💾 Baixar Lista de Adequações (.docx)", 
                                    data=buffer_exig, 
                                    file_name=f"Lista_Adequacoes_{web_ativa}.docx", 
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True
                                )
                except Exception as e:
                    st.error(f"Erro ao acessar base de dados: {e}")
                    
        # 2. BOTÃO: Auto de Infração
        if st.button("🚨 Gerar Auto de Infração", use_container_width=True):
            # A lógica interna de montagem do Word será descrita e codificada mais tarde
            st.info("ℹ️ Recurso selecionado. A rotina de geração do Auto de Infração será configurada em breve.")
            
        # 3. BOTÃO: Auto de Imposição de Penalidade
        if st.button("⚖️ Gerar Auto de Imposição de Penalidade", use_container_width=True):
            # A lógica interna de montagem do Word será descrita e codificada mais tarde
            st.info("ℹ️ Recurso selecionado. A rotina de geração do Auto de Imposição de Penalidade será configurada em breve.")

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
    
    # --- MAPEAMENTOS GERAIS DE STATUS PARA PREENCHIMENTO AUTOMÁTICO ---
    mapeamento_status = {"Não avaliado": 0, "C": 1, "NC": 2, "NA": 3}
    opcoes_status = ["Não avaliado", "C", "NC", "NA"]

    mapeamento_sim_nao = {"Não avaliado": 0, "Sim": 1, "Não": 2}
    opcoes_sim_nao = ["Não avaliado", "Sim", "Não"]
    
    # Atalho para o dicionário do histórico mestre carregado na Tela 1
    historico = st.session_state.get("historico_pre_preenchido", {})
    
    # Cria uma chave única baseada no número WEB e no horário atual do sistema
    timestamp_chave = datetime.now().strftime("%Y%m%d_%H%M%S")
    id_inspecao_unico = f"{n_web or 'INSP'}_{timestamp_chave}"

    # Dicionário mestre de coleta (agora contendo o id_inspecao obrigatório!)
    dados_inspecao_atual = {
        "id_inspecao": id_inspecao_unico, # ⬅️ A correção cirúrgica está aqui!
        "cnpj": cnpj_estabelecimento,
        "web": n_web,
        "data_procedimento": datetime.now().date()
    }
    
    # --- CRIAÇÃO DAS TABS DO APLICATIVO ---
    (
        tab_contato, 
        tab_documentacao, 
        tab_controles, 
        tab_comercializacao, 
        tab_areas, 
        tab_sala, 
        tab_funcionarios
    ) = st.tabs([
        "👥 Contato", 
        "📄 Documentação", 
        "❄️ Controles T e U", 
        "📦 Comercialização", 
        "🏢 Áreas Físicas", 
        "🩺 Sala de Serviços", 
        "👷 Funcionários"
    ])

    # ==============================================================================
    # TAB 1: CONTATO (Preenchimento das Pessoas Contatadas)
    # ==============================================================================
    with tab_contato:
        st.markdown("### 👥 Acompanhantes da Inspeção")
        
        # --- MOTOR DE RECONSTRUÇÃO DO HISTÓRICO ---
        # Se o fiscal mudou de empresa ou iniciou o formulário agora, preparamos a lista de pessoas
        if "lista_pessoas_formulario" not in st.session_state or st.session_state.get("ultimo_cnpj_visto") != cnpj_estabelecimento:
            st.session_state["ultimo_cnpj_visto"] = cnpj_estabelecimento
            texto_banco_contatos = str(historico.get("acompanhantes_inspecao", ""))
            
            # Se existirem dados anteriores gravados na última inspeção, reconstrói a lista
            if texto_banco_contatos.strip() != "" and texto_banco_contatos != "None" and "/" in texto_banco_contatos:
                lista_reconstruida = []
                # Se houver mais de uma pessoa, quebra no " | ", senão trata como lista de um item só
                if " | " in texto_banco_contatos:
                    pessoas_registradas = texto_banco_contatos.split(" | ")
                else:
                    pessoas_registradas = [texto_banco_contatos]
                    
                for registro in pessoas_registradas:
                    if "/" in registro:
                        partes = registro.split("/")
                        # Garante que existem os 4 campos na quebra para não dar erro de índice
                        if len(partes) == 4:
                            lista_reconstruida.append({
                                "nome": partes[0].strip(),
                                "cpf": partes[1].strip(),
                                "cargo": partes[2].strip(),
                                "email": partes[3].strip()
                            })
                st.session_state["lista_pessoas_formulario"] = lista_reconstruida if lista_reconstruida else [{"nome": "", "cpf": "", "cargo": "", "email": ""}]
            else:
                # Se for uma empresa nova sem nenhuma inspeção anterior, inicia o formulário em branco
                st.session_state["lista_pessoas_formulario"] = [{"nome": "", "cpf": "", "cargo": "", "email": ""}]

        # --- RENDERIZAÇÃO DOS CAMPOS DE CADA PESSOA ---
        lista_atualizada_salvamento = []
        
        for idx, pessoa in enumerate(st.session_state["lista_pessoas_formulario"]):
            st.markdown(f"**Dados da Pessoa nº {idx + 1}**")
            col_nome, col_cpf, col_cargo, col_email = st.columns([2, 1, 1, 1.5])
            
            with col_nome:
                nome_p = st.text_input(f"Nome completo:", value=pessoa["nome"], key=f"p_nome_{idx}")
            with col_cpf:
                cpf_p = st.text_input(f"CPF:", value=pessoa["cpf"], max_chars=14, key=f"p_cpf_{idx}")
            with col_cargo:
                cargo_p = st.text_input(f"Cargo / Função:", value=pessoa["cargo"], key=f"p_cargo_{idx}")
            with col_email:
                email_p = st.text_input(f"E-mail de contato:", value=pessoa["email"], key=f"p_email_{idx}")
                
            # Adiciona os dados coletados na lista de envio para o banco
            if nome_p.strip() != "":
                lista_atualizada_salvamento.append(f"{nome_p.strip()}/{cpf_p.strip()}/{cargo_p.strip()}/{email_p.strip()}")
                
        # --- BOTÕES DE GERENCIAMENTO DINÂMICO ---
        col_btn_add, col_btn_rem, _ = st.columns([1, 1, 3])
        
        with col_btn_add:
            if st.button("➕ Adicionar Pessoa", use_container_width=True, key="btn_add_pessoa"):
                # Captura o estado digitado atualmente antes de adicionar a nova linha em branco
                nova_lista_temp = []
                for i in range(len(st.session_state["lista_pessoas_formulario"])):
                    nova_lista_temp.append({
                        "nome": st.session_state[f"p_nome_{i}"],
                        "cpf": st.session_state[f"p_cpf_{i}"],
                        "cargo": st.session_state[f"p_cargo_{i}"],
                        "email": st.session_state[f"p_email_{i}"]
                    })
                nova_lista_temp.append({"nome": "", "cpf": "", "cargo": "", "email": ""})
                st.session_state["lista_pessoas_formulario"] = nova_lista_temp
                st.rerun()
                
        with col_btn_rem:
            if len(st.session_state["lista_pessoas_formulario"]) > 1:
                if st.button("❌ Remover Última", use_container_width=True, key="btn_rem_pessoa"):
                    # Salva as informações mantendo as linhas anteriores e eliminando apenas o último item
                    nova_lista_temp = []
                    for i in range(len(st.session_state["lista_pessoas_formulario"]) - 1):
                        nova_lista_temp.append({
                            "nome": st.session_state[f"p_nome_{i}"],
                            "cpf": st.session_state[f"p_cpf_{i}"],
                            "cargo": st.session_state[f"p_cargo_{i}"],
                            "email": st.session_state[f"p_email_{i}"]
                        })
                    st.session_state["lista_pessoas_formulario"] = nova_lista_temp
                    st.rerun()

        # Transforma o array em um texto compacto separado por barras prontas para persistir no Supabase
        dados_inspecao_atual["acompanhantes_inspecao"] = " | ".join(lista_atualizada_salvamento)

        st.divider()

        # --- PERGUNTA SANITÁRIA COMPLEMENTAR (RT PRESENTE) ---
        st.markdown("### 💊 Responsabilidade Técnica no Ato")
        st.markdown("**Há farmacêutico responsável no momento da inspeção?**")
        
        # Mapeamento especial: Se não houver dados históricos, inicia marcado em "Não avaliado" (índice 0)
        mapeamento_rt = {"Não avaliado": 0, "C": 1, "NC": 2, "NA": 3}
        
        # Busca o histórico da sua nova coluna dedicada
        val_anterior_rt = historico.get("farmaceutico_presente", "NA") if historico else "NA"
        idx_inicial_rt = mapeamento_rt.get(val_anterior_rt, 0)
        
        # Salva o resultado diretamente na nova coluna dedicada da tabela
        dados_inspecao_atual["farmaceutico_presente"] = st.radio(
            "Avaliação RT Presente:",
            options=["Não avaliado", "C", "NC", "NA"],
            index=idx_inicial_rt,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_farmaceutico_presente_ato"
        )
     
    # ==============================================================================
    # TAB 2: DOCUMENTAÇÃO
    # ==============================================================================

    with tab_documentacao:
        st.markdown("## 📄 Seção: Documentação Regulamentar")
        
        # --------------------------------------------------------------------------
        # 2.1 AUTORIZAÇÃO DE FUNCIONAMENTO DE EMPRESA (AFE)
        # --------------------------------------------------------------------------
        st.markdown("### 🏢 Autorização de Funcionamento (AFE / ANVISA)")
        st.markdown("**O estabelecimento possui Autorização de Funcionamento de Empresa (AFE) publicada pela ANVISA?**")
        
        # 1. Recupera a avaliação sanitária anterior da AFE
        val_possui_afe = historico.get("possui_afe", "Não avaliado")
        idx_possui_afe = mapeamento_status.get(val_possui_afe, 0)
        
        dados_inspecao_atual["possui_afe"] = st.radio(
            "Avaliação AFE:",
            options=opcoes_status,
            index=idx_possui_afe,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_afe"
        )
        
        # Lógica Dinâmica: Só exibe os campos de detalhamento se o item for Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_afe"] in ["C", "NC"]:
            st.markdown("#### 📝 Dados Cadastrais da AFE")
            col_afe_num, col_afe_dat, col_afe_sts = st.columns(3)
            
            with col_afe_num:
                val_num_afe = historico.get("numero_afe", "")
                dados_inspecao_atual["numero_afe"] = st.text_input(
                    "Número da AFE (Processo/Resolução):", 
                    value=str(val_num_afe) if val_num_afe is not None and str(val_num_afe) != "None" else "", 
                    key="ti_numero_afe"
                )
                
            with col_afe_dat:
                val_dat_afe = historico.get("data_afe", None)
                # Converte o retorno do banco para o tipo date do Python de forma segura
                dt_padrao_afe = pd.to_datetime(val_dat_afe).date() if val_dat_afe and not pd.isna(val_dat_afe) else None
                
                dados_inspecao_atual["data_afe"] = st.date_input(
                    "Data de Publicação no DOU:", 
                    value=dt_padrao_afe, 
                    key="di_data_afe"
                )
                
            with col_afe_sts:
                # Mapeamento do status documental (Ativa, Em análise, Cancelada)
                opcoes_status_afe = ["Não informado", "Ativa", "Em análise", "Cancelada"]
                mapeamento_status_afe = {v: i for i, v in enumerate(opcoes_status_afe)}
                
                val_sts_afe = historico.get("status_afe", "Não informado")
                idx_sts_afe = mapeamento_status_afe.get(val_sts_afe, 0)
                
                dados_inspecao_atual["status_afe"] = st.selectbox(
                    "Situação Cadastral da AFE:",
                    options=opcoes_status_afe,
                    index=idx_sts_afe,
                    key="sb_status_afe"
                )
        else:
            # Caso esteja marcado como "Não se aplica" ou "Não avaliado", zera as variáveis no dicionário de envio
            dados_inspecao_atual["numero_afe"] = None
            dados_inspecao_atual["data_afe"] = None
            dados_inspecao_atual["status_afe"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.2 CERTIDÃO DE REGULARIDADE TÉCNICA (CRT)
        # --------------------------------------------------------------------------
        st.markdown("### 📜 Certidão de Regularidade Técnica (CRT)")
        st.markdown("**O estabelecimento possui Certidão de Regularidade Técnica emitida pelo Conselho Regional de Farmácia (CRF) atualizada?**")
        
        # 1. Recupera a avaliação sanitária anterior da CRT (Padrão: "Não avaliado")
        val_possui_crt = historico.get("possui_crt", "Não avaliado")
        idx_possui_crt = mapeamento_status.get(val_possui_crt, 0)
        
        dados_inspecao_atual["possui_crt"] = st.radio(
            "Avaliação CRT:",
            options=opcoes_status,
            index=idx_possui_crt,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_crt"
        )
        
        # Lógica Dinâmica: Só exibe os campos de detalhamento se o item for Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_crt"] in ["C", "NC"]:
            st.markdown("#### 📝 Dados Cadastrais da CRT")
            col_crt_num, col_crt_dat, col_crt_sts = st.columns(3)
            
            with col_crt_num:
                val_num_crt = historico.get("numero_crt", "")
                dados_inspecao_atual["numero_crt"] = st.text_input(
                    "Número da Certidão / CRT:", 
                    value=str(val_num_crt) if val_num_crt is not None and str(val_num_crt) != "None" else "", 
                    key="ti_numero_crt"
                )
                
            with col_crt_dat:
                val_dat_crt = historico.get("data_crt", None)
                # Converte o retorno do banco para o tipo date do Python de forma segura
                dt_padrao_crt = pd.to_datetime(val_dat_crt).date() if val_dat_crt and not pd.isna(val_dat_crt) else None
                
                dados_inspecao_atual["data_crt"] = st.date_input(
                    "Data de Emissão da Certidão:", 
                    value=dt_padrao_crt, 
                    key="di_data_crt"
                )
                
            with col_crt_sts:
                # Mapeamento do status documental específico (Válida, Inválida, Pendente)
                opcoes_status_crt = ["Não informado", "Válida", "Inválida", "Pendente"]
                mapeamento_status_crt = {v: i for i, v in enumerate(opcoes_status_crt)}
                
                val_sts_crt = historico.get("status_crt", "Não informado")
                idx_sts_crt = mapeamento_status_crt.get(val_sts_crt, 0)
                
                dados_inspecao_atual["status_crt"] = st.selectbox(
                    "Situação Cadastral da CRT:",
                    options=opcoes_status_crt,
                    index=idx_sts_crt,
                    key="sb_status_crt"
                )
                
            # Campo para o Responsável Técnico descrito no documento
            val_rt_crt = historico.get("rt_crt", "")
            dados_inspecao_atual["rt_crt"] = st.text_input(
                "Nome do Farmacêutico Responsável Técnico (descrito na CRT):",
                value=str(val_rt_crt) if val_rt_crt is not None and str(val_rt_crt) != "None" else "",
                key="ti_rt_crt"
            )
        else:
            # Caso esteja oculto, limpa os campos para gravação no banco
            dados_inspecao_atual["numero_crt"] = None
            dados_inspecao_atual["data_crt"] = None
            dados_inspecao_atual["status_crt"] = None
            dados_inspecao_atual["rt_crt"] = None

        st.divider()
       
        # --------------------------------------------------------------------------
        # 2.3 MANUAIS E PROCEDIMENTOS OPERACIONAIS (MBPF E POPS)
        # --------------------------------------------------------------------------
        st.markdown("### 📚 Manuais e Procedimentos Operacionais")
        
        col_mbpf, col_pops = st.columns(2)
        
        with col_mbpf:
            st.markdown("**Possui Manual de Boas Práticas Farmacêuticas (MBPF) atualizado e adequado à realidade do estabelecimento?**")
            
            # Recupera o histórico do MBPF ou assume "Não avaliado" de fábrica
            val_mbpf = historico.get("possui_mbpf", "Não avaliado")
            idx_mbpf = mapeamento_status.get(val_mbpf, 0)
            
            dados_inspecao_atual["possui_mbpf"] = st.radio(
                "Avaliação MBPF:",
                options=opcoes_status,
                index=idx_mbpf,
                horizontal=True,
                label_visibility="collapsed",
                key="rad_possui_mbpf"
            )
            
        with col_pops:
            st.markdown("**Possui Procedimentos Operacionais Padrão (POPs) implantados, atualizados e assinados pelo RT?**")
            
            # Recupera o histórico dos POPs ou assume "Não avaliado" de fábrica
            val_pops = historico.get("possui_pops", "Não avaliado")
            idx_pops = mapeamento_status.get(val_pops, 0)
            
            dados_inspecao_atual["possui_pops"] = st.radio(
                "Avaliação POPs:",
                options=opcoes_status,
                index=idx_pops,
                horizontal=True,
                label_visibility="collapsed",
                key="rad_possui_pops"
            )

        st.divider()

        # --------------------------------------------------------------------------
        # 2.4 CERTIFICADO DE CONTROLE DE PRAGAS (DEDETIZAÇÃO)
        # --------------------------------------------------------------------------
        st.markdown("### 🪰 Certificado de Controle de Vetores e Pragas")
        st.markdown("**O estabelecimento possui comprovante de execução de serviço de controle de vetores e pragas urbanas por empresa licenciada?**")
        
        # 1. Recupera a avaliação sanitária anterior do controle de pragas (Padrão: "Não avaliado")
        val_possui_pragas = historico.get("possui_pragas", "Não avaliado")
        idx_possui_pragas = mapeamento_status.get(val_possui_pragas, 0)
        
        dados_inspecao_atual["possui_pragas"] = st.radio(
            "Avaliação Controle de Pragas:",
            options=opcoes_status,
            index=idx_possui_pragas,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_pragas"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_pragas"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados do Comprovante de Dedetização")
            
            # Linha 1: Informações da Empresa Executora
            col_praga_emp, col_praga_cnpj = st.columns(2)
            
            with col_praga_emp:
                val_emp_pragas = historico.get("empresa_pragas", "")
                dados_inspecao_atual["empresa_pragas"] = st.text_input(
                    "Razão Social da Empresa Controladora:", 
                    value=str(val_emp_pragas) if val_emp_pragas is not None and str(val_emp_pragas) != "None" else "", 
                    key="ti_empresa_pragas"
                )
                
            with col_praga_cnpj:
                val_cnpj_pragas = historico.get("cnpj_pragas", "")
                dados_inspecao_atual["cnpj_pragas"] = st.text_input(
                    "CNPJ da Empresa Controladora:", 
                    value=str(val_cnpj_pragas) if val_cnpj_pragas is not None and str(val_cnpj_pragas) != "None" else "", 
                    key="ti_cnpj_pragas"
                )
                
            # Linha 2: Datas Importantes e Prazos
            col_praga_dt, col_praga_val, col_praga_alv = st.columns(3)
            
            with col_praga_dt:
                val_dt_pragas = historico.get("data_pragas", None)
                dt_padrao_pragas = pd.to_datetime(val_dt_pragas).date() if val_dt_pragas and not pd.isna(val_dt_pragas) else None
                
                dados_inspecao_atual["data_pragas"] = st.date_input(
                    "Data da Aplicação:", 
                    value=dt_padrao_pragas, 
                    key="di_data_pragas"
                )
                
            with col_praga_val:
                val_vld_pragas = historico.get("validade_pragas", None)
                vld_padrao_pragas = pd.to_datetime(val_vld_pragas).date() if val_vld_pragas and not pd.isna(val_vld_pragas) else None
                
                dados_inspecao_atual["validade_pragas"] = st.date_input(
                    "Validade do Laudo / Certificado:", 
                    value=vld_padrao_pragas, 
                    key="di_validade_pragas"
                )
                
            with col_praga_alv:
                val_alv_pragas = historico.get("val_alvara_pragas", None)
                alv_padrao_pragas = pd.to_datetime(val_alv_pragas).date() if val_alv_pragas and not pd.isna(val_alv_pragas) else None
                
                dados_inspecao_atual["val_alvara_pragas"] = st.date_input(
                    "Validade do Alvará da Empresa Controladora:", 
                    value=alv_padrao_pragas, 
                    key="di_val_alvara_pragas"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento
            dados_inspecao_atual["empresa_pragas"] = None
            dados_inspecao_atual["cnpj_pragas"] = None
            dados_inspecao_atual["data_pragas"] = None
            dados_inspecao_atual["validade_pragas"] = None
            dados_inspecao_atual["val_alvara_pragas"] = None

        st.divider()
        # --------------------------------------------------------------------------
        # 2.5 CERTIFICADO DE HIGIENIZAÇÃO DO RESERVATÓRIO DE ÁGUA
        # --------------------------------------------------------------------------
        st.markdown("### 🚰 Certificado de Higienização do Reservatório de Água")
        st.markdown("**O estabelecimento possui comprovante de execução nos últimos 6 meses?**")
        
        # 1. Recupera a avaliação sanitária anterior da limpeza de água (Padrão: "Não avaliado")
        val_possui_agua = historico.get("possui_limpeza_reserv_agua", "Não avaliado")
        idx_possui_agua = mapeamento_status.get(val_possui_agua, 0)
        
        dados_inspecao_atual["possui_limpeza_reserv_agua"] = st.radio(
            "Avaliação Limpeza Reservatório:",
            options=opcoes_status,
            index=idx_possui_agua,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_limpeza_reserv_agua"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_limpeza_reserv_agua"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados do Comprovante de Limpeza de Caixa d'Água")
            
            # Linha 1: Informações da Empresa Contratada e Data do Serviço
            col_agua_emp, col_agua_cnpj, col_agua_dt = st.columns([2, 1.5, 1])
            
            with col_agua_emp:
                val_emp_agua = historico.get("empresa_reserv_agua", "")
                dados_inspecao_atual["empresa_reserv_agua"] = st.text_input(
                    "Razão Social da Empresa Executora:", 
                    value=str(val_emp_agua) if val_emp_agua is not None and str(val_emp_agua) != "None" else "", 
                    key="ti_empresa_reserv_agua"
                )
                
            with col_agua_cnpj:
                val_cnpj_agua = historico.get("cnpj_reserv_agua", "")
                dados_inspecao_atual["cnpj_reserv_agua"] = st.text_input(
                    "CNPJ da Empresa Executora:", 
                    value=str(val_cnpj_agua) if val_cnpj_agua is not None and str(val_cnpj_agua) != "None" else "", 
                    key="ti_cnpj_reserv_agua"
                )
                
            with col_agua_dt:
                val_dt_agua = historico.get("data_reserv_agua", None)
                # Converte o retorno do banco para o tipo date de forma segura
                dt_padrao_agua = pd.to_datetime(val_dt_agua).date() if val_dt_agua and not pd.isna(val_dt_agua) else None
                
                dados_inspecao_atual["data_reserv_agua"] = st.date_input(
                    "Data da Execução:", 
                    value=dt_padrao_agua, 
                    key="di_data_reserv_agua"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento documental
            dados_inspecao_atual["empresa_reserv_agua"] = None
            dados_inspecao_atual["cnpj_reserv_agua"] = None
            dados_inspecao_atual["data_reserv_agua"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.6 AUTO DE VISTORIA DO CORPO DE BOMBEIROS (AVCB / CLCB / Protocolo)
        # --------------------------------------------------------------------------
        st.markdown("### 🚒 Auto de Vistoria / Licença do Corpo de Bombeiros")
        st.markdown("**O estabelecimento possui Auto de Vistoria do Corpo de Bombeiros (AVCB), Certificado de Licença (CLCB) ou Protocolo válido?**")
        
        # 1. Recupera a avaliação sanitária anterior do Corpo de Bombeiros (Padrão: "Não avaliado")
        val_possui_bombeiro = historico.get("possui_vistoria_bombeiro", "Não avaliado")
        idx_possui_bombeiro = mapeamento_status.get(val_possui_bombeiro, 0)
        
        dados_inspecao_atual["possui_vistoria_bombeiro"] = st.radio(
            "Avaliação Bombeiros:",
            options=opcoes_status,
            index=idx_possui_bombeiro,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_vistoria_bombeiro"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_vistoria_bombeiro"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados do Documento do Corpo de Bombeiros")
            
            # Linha 1: Tipo, Número, Emissão e Validade organizados em 4 colunas
            col_bomb_tipo, col_bomb_num, col_bomb_dt, col_bomb_val = st.columns([1.2, 1.5, 1, 1])
            
            with col_bomb_tipo:
                # Mapeamento do tipo de documento do bombeiro
                opcoes_tipo_bomb = ["Não informado", "AVCB", "CLCB", "Protocolo"]
                mapeamento_tipo_bomb = {v: i for i, v in enumerate(opcoes_tipo_bomb)}
                
                val_tipo_bomb = historico.get("tipo_vistoria_bombeiro", "Não informado")
                idx_tipo_bomb = mapeamento_tipo_bomb.get(val_tipo_bomb, 0)
                
                dados_inspecao_atual["tipo_vistoria_bombeiro"] = st.selectbox(
                    "Tipo de Documento:",
                    options=opcoes_tipo_bomb,
                    index=idx_tipo_bomb,
                    key="sb_tipo_vistoria_bombeiro"
                )
            
            with col_bomb_num:
                val_num_bomb = historico.get("num_vistoria_bombeiro", "")
                dados_inspecao_atual["num_vistoria_bombeiro"] = st.text_input(
                    "Número do Documento:", 
                    value=str(val_num_bomb) if val_num_bomb is not None and str(val_num_bomb) != "None" else "", 
                    placeholder="Ex: 54321/2026",
                    key="ti_num_vistoria_bombeiro"
                )
                
            with col_bomb_dt:
                val_dt_bomb = historico.get("data_vistoria_bombeiro", None)
                # Converte o retorno do banco para o tipo date de forma segura
                dt_padrao_bomb = pd.to_datetime(val_dt_bomb).date() if val_dt_bomb and not pd.isna(val_dt_bomb) else None
                
                dados_inspecao_atual["data_vistoria_bombeiro"] = st.date_input(
                    "Data de Emissão:", 
                    value=dt_padrao_bomb, 
                    key="di_data_vistoria_bombeiro"
                )
                
            with col_bomb_val:
                val_vld_bomb = historico.get("validade_vistoria_bombeiro", None)
                vld_padrao_bomb = pd.to_datetime(val_vld_bomb).date() if val_vld_bomb and not pd.isna(val_vld_bomb) else None
                
                dados_inspecao_atual["validade_vistoria_bombeiro"] = st.date_input(
                    "Data de Validade:", 
                    value=vld_padrao_bomb, 
                    key="di_validade_vistoria_bombeiro"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento documental
            dados_inspecao_atual["tipo_vistoria_bombeiro"] = None
            dados_inspecao_atual["num_vistoria_bombeiro"] = None
            dados_inspecao_atual["data_vistoria_bombeiro"] = None
            dados_inspecao_atual["validade_vistoria_bombeiro"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.7 PLANO DE GERENCIAMENTO DE RESÍDUOS DE SERVIÇOS DE SAÚDE (PGRSS)
        # --------------------------------------------------------------------------
        st.markdown("### 🪣 Plano de Gerenciamento de Resíduos (PGRSS)")
        st.markdown("**O estabelecimento possui PGRSS implantado e atualizado?**")
        
        # 1. Recupera a avaliação sanitária anterior do PGRSS (Padrão: "Não avaliado")
        val_possui_pgrss = historico.get("possui_pgrss", "Não avaliado")
        idx_possui_pgrss = mapeamento_status.get(val_possui_pgrss, 0)
        
        dados_inspecao_atual["possui_pgrss"] = st.radio(
            "Avaliação PGRSS:",
            options=opcoes_status,
            index=idx_possui_pgrss,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_pgrss"
        )
        
        # Lógica Dinâmica: Só exibe os detalhes se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_pgrss"] in ["C", "NC"]:
            st.markdown("#### 📄 Detalhes e Comprovantes do PGRSS")
            
            col_pgrss_dt, col_pgrss_coleta = st.columns([1, 2])
            
            with col_pgrss_dt:
                val_dt_pgrss = historico.get("data_pgrss", None)
                # Converte o retorno do banco para o tipo date de forma segura
                dt_padrao_pgrss = pd.to_datetime(val_dt_pgrss).date() if val_dt_pgrss and not pd.isna(val_dt_pgrss) else None
                
                dados_inspecao_atual["data_pgrss"] = st.date_input(
                    "Data de Emissão do PGRSS:", 
                    value=dt_padrao_pgrss, 
                    key="di_data_pgrss"
                )
                
            with col_pgrss_coleta:
                # Configuração das opções personalizadas que você especificou para a coleta
                opcoes_comp_coleta = [
                    "Não informado", 
                    "Sim", 
                    "Não", 
                    "Utiliza planilha para assinatura do responsável pela coleta de resíduo infectante"
                ]
                mapeamento_comp_coleta = {v: i for i, v in enumerate(opcoes_comp_coleta)}
                
                val_comp_coleta = historico.get("comp_coleta_pgrss", "Não informado")
                idx_comp_coleta = mapeamento_comp_coleta.get(val_comp_coleta, 0)
                
                dados_inspecao_atual["comp_coleta_pgrss"] = st.selectbox(
                    "Apresentou comprovante de coleta de resíduos grupos A, B e E?",
                    options=opcoes_comp_coleta,
                    index=idx_comp_coleta,
                    key="sb_comp_coleta_pgrss"
                )
            
            # Campo adicional para a data da última retirada de resíduos
            val_dt_coleta = historico.get("data_coleta_residuo", None)
            dt_padrao_coleta = pd.to_datetime(val_dt_coleta).date() if val_dt_coleta and not pd.isna(val_dt_coleta) else None
            
            dados_inspecao_atual["data_coleta_residuo"] = st.date_input(
                "Data da última retirada de resíduo infectante:",
                value=dt_padrao_coleta,
                key="di_data_coleta_residuo"
            )
        else:
            # Limpa as variáveis no dicionário caso o item principal não seja exigido/avaliado
            dados_inspecao_atual["data_pgrss"] = None
            dados_inspecao_atual["comp_coleta_pgrss"] = None
            dados_inspecao_atual["data_coleta_residuo"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.8 PROGRAMA DE CONTROLE MÉDICO DE SAÚDE OCUPACIONAL (PCMSO)
        # --------------------------------------------------------------------------
        st.markdown("### 🩺 Programa de Controle Médico de Saúde Ocupacional (PCMSO)")
        st.markdown("**O estabelecimento possui o Programa de Controle Médico de Saúde Ocupacional (PCMSO) implantado e atualizado?**")
        
        # 1. Recupera a avaliação sanitária anterior do PCMSO (Padrão: "Não avaliado")
        val_possui_pcmso = historico.get("possui_pcmso", "Não avaliado")
        idx_possui_pcmso = mapeamento_status.get(val_possui_pcmso, 0)
        
        dados_inspecao_atual["possui_pcmso"] = st.radio(
            "Avaliação PCMSO:",
            options=opcoes_status,
            index=idx_possui_pcmso,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_pcmso"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_pcmso"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados do Documento do PCMSO")
            
            # Linha 1: Informações da Empresa Executora e Médico Responsável
            col_pcmso_emp, col_pcmso_med = st.columns(2)
            
            with col_pcmso_emp:
                val_emp_pcmso = historico.get("empresa_pcmso", "")
                dados_inspecao_atual["empresa_pcmso"] = st.text_input(
                    "Empresa de Medicina do Trabalho Executora:", 
                    value=str(val_emp_pcmso) if val_emp_pcmso is not None and str(val_emp_pcmso) != "None" else "", 
                    key="ti_empresa_pcmso"
                )
                
            with col_pcmso_med:
                val_med_pcmso = historico.get("medico_pcmso", "")
                dados_inspecao_atual["medico_pcmso"] = st.text_input(
                    "Nome do Médico Coordenador / Responsável:", 
                    value=str(val_med_pcmso) if val_med_pcmso is not None and str(val_med_pcmso) != "None" else "", 
                    placeholder="Ex: Dr. Carlos Eduardo - CRM/SP XXXXX",
                    key="ti_medico_pcmso"
                )
                
            # Linha 2: Datas Importantes (Emissão e Validade)
            col_pcmso_emi, col_pcmso_val = st.columns(2)
            
            with col_pcmso_emi:
                val_emi_pcmso = historico.get("emissao_pcmso", None)
                # Converte o retorno do banco para o tipo date de forma segura
                emi_padrao_pcmso = pd.to_datetime(val_emi_pcmso).date() if val_emi_pcmso and not pd.isna(val_emi_pcmso) else None
                
                dados_inspecao_atual["emissao_pcmso"] = st.date_input(
                    "Data de Emissão:", 
                    value=emi_padrao_pcmso, 
                    key="di_emissao_pcmso"
                )
                
            with col_pcmso_val:
                val_vld_pcmso = historico.get("validade_pcmso", None)
                vld_padrao_pcmso = pd.to_datetime(val_vld_pcmso).date() if val_vld_pcmso and not pd.isna(val_vld_pcmso) else None
                
                dados_inspecao_atual["validade_pcmso"] = st.date_input(
                    "Data de Validade / Renovação:", 
                    value=vld_padrao_pcmso, 
                    key="di_validade_pcmso"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento documental
            dados_inspecao_atual["emissao_pcmso"] = None
            dados_inspecao_atual["validade_pcmso"] = None
            dados_inspecao_atual["empresa_pcmso"] = None
            dados_inspecao_atual["medico_pcmso"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.9 PROGRAMA DE GERENCIAMENTO DE RISCOS (PGR)
        # --------------------------------------------------------------------------
        st.markdown("### 🛠️ Programa de Gerenciamento de Riscos (PGR)")
        st.markdown("**O estabelecimento possui o PGR implantado e em vigência?**")
        
        # 1. Recupera a avaliação sanitária anterior do PGR (Padrão: "Não avaliado")
        val_possui_pgr = historico.get("possui_pgr", "Não avaliado")
        idx_possui_pgr = mapeamento_status.get(val_possui_pgr, 0)
        
        dados_inspecao_atual["possui_pgr"] = st.radio(
            "Avaliação PGR:",
            options=opcoes_status,
            index=idx_possui_pgr,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_pgr"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_pgr"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados do Documento do PGR")
            
            # Linha 1: Responsável Elaborador
            val_resp_pgr = historico.get("responsavel_pgr", "")
            dados_inspecao_atual["responsavel_pgr"] = st.text_input(
                "Nome do Responsável Técnico pela Elaboração do PGR:", 
                value=str(val_resp_pgr) if val_resp_pgr is not None and str(val_resp_pgr) != "None" else "", 
                placeholder="Ex: Engenheiro/ Técnico de Segurança do Trabalho / nº MTE",
                key="ti_responsavel_pgr"
            )
            
            # Linha 2: Datas Importantes (Emissão e Validade) organizadas em 2 colunas
            col_pgr_emi, col_pgr_val = st.columns(2)
            
            with col_pgr_emi:
                val_emi_pgr = historico.get("emissao_pgr", None)
                # Converte o retorno do banco para o tipo date de forma segura
                emi_padrao_pgr = pd.to_datetime(val_emi_pgr).date() if val_emi_pgr and not pd.isna(val_emi_pgr) else None
                
                dados_inspecao_atual["emissao_pgr"] = st.date_input(
                    "Data de Emissão / Início da Vigência:", 
                    value=emi_padrao_pgr, 
                    key="di_emissao_pgr"
                )
                
            with col_pgr_val:
                val_vld_pgr = historico.get("validade_pgr", None)
                vld_padrao_pgr = pd.to_datetime(val_vld_pgr).date() if val_vld_pgr and not pd.isna(val_vld_pgr) else None
                
                dados_inspecao_atual["validade_pgr"] = st.date_input(
                    "Data de Validade:", 
                    value=vld_padrao_pgr, 
                    key="di_validade_pgr"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento documental
            dados_inspecao_atual["responsavel_pgr"] = None
            dados_inspecao_atual["emissao_pgr"] = None
            dados_inspecao_atual["validade_pgr"] = None

        st.divider()

        # --------------------------------------------------------------------------
        # 2.10 ATESTADO DE SAÚDE OCUPACIONAL (ASO)
        # --------------------------------------------------------------------------
        st.markdown("### 🪪 Atestado de Saúde Ocupacional (ASO)")
        st.markdown("**O estabelecimento apresenta os ASOs atualizados?**")
        
        # 1. Recupera a avaliação sanitária anterior do ASO (Padrão: "Não avaliado")
        val_possui_aso = historico.get("possui_aso", "Não avaliado")
        idx_possui_aso = mapeamento_status.get(val_possui_aso, 0)
        
        dados_inspecao_atual["possui_aso"] = st.radio(
            "Avaliação ASO:",
            options=opcoes_status,
            index=idx_possui_aso,
            horizontal=True,
            label_visibility="collapsed",
            key="rad_possui_aso"
        )
        
        # Lógica Dinâmica: Só exibe os campos documentais se for avaliado como Conforme (C) ou Não Conforme (NC)
        if dados_inspecao_atual["possui_aso"] in ["C", "NC"]:
            st.markdown("#### 📄 Dados de Controle dos ASOs")
            col_aso_emp, col_aso_num = st.columns(2)
            
            with col_aso_emp:
                val_emp_aso = historico.get("empresa_aso", "")
                dados_inspecao_atual["empresa_aso"] = st.text_input(
                    "Clínica / Empresa responsável pela emissão dos ASOs:", 
                    value=str(val_emp_aso) if val_emp_aso is not None and str(val_emp_aso) != "None" else "", 
                    key="ti_empresa_aso"
                )
                
            with col_aso_num:
                val_num_aso = historico.get("numero_aso", "")
                dados_inspecao_atual["numero_aso"] = st.text_input(
                    "Quantidade de ASOs verificados / Observações de controle:", 
                    value=str(val_num_aso) if val_num_aso is not None and str(val_num_aso) != "None" else "", 
                    key="ti_numero_aso"
                )
        else:
            # Limpa as variáveis no dicionário caso o item não exija detalhamento documental
            dados_inspecao_atual["empresa_aso"] = None
            dados_inspecao_atual["numero_aso"] = None

        st.divider()

    # ==============================================================================
    # TAB 3: CONTROLES T E U
    # ==============================================================================
        # ==============================================================================
    # TAB 3: CONTROLES T E U (ÁREAS TÉRMICAS E GELADEIRAS)
    # ==============================================================================
    with tab_controles:
        st.markdown("## ❄️ Seção: Áreas de Controle Térmico/ Umidade e Geladeiras")
        
        # --- RECONSTRUÇÃO DO HISTÓRICO DO SUPABASE ---
        # Resgata o objeto JSONB salvo na última inspeção
        historico_json = historico.get("locais_controle_geladeiras_json", {})
        if not isinstance(historico_json, dict):
            historico_json = {}
            
        historico_lista_tu = historico_json.get("lista_tu", [])
        historico_lista_t = historico_json.get("lista_t", [])

        # --- INICIALIZAÇÃO SEGURA DOS ESTADOS DE SESSÃO DOS LOOPS ---
        if "lista_tu_formulario" not in st.session_state or st.session_state.get("ultimo_cnpj_visto_tu") != cnpj_estabelecimento:
            st.session_state["ultimo_cnpj_visto_tu"] = cnpj_estabelecimento
            st.session_state["lista_tu_formulario"] = historico_lista_tu if historico_lista_tu else []
            st.session_state["lista_t_formulario"] = historico_lista_t if historico_lista_t else []

        # Opções estritas de marcação para os seletores de Calibração
        opcoes_tipo_comprovacao = ["Calibração", "Aquisição recente"]

        # ==============================================================================
        # PARTE A: LOOP - CONTROLE DE TEMPERATURA E UMIDADE (T E U)
        # ==============================================================================
        st.markdown("### 📍 Áreas de Temperatura e Umidade Controladas")
        
        lista_tu_atualizada = []
        
        for idx, item in enumerate(st.session_state["lista_tu_formulario"]):
            with st.expander(f"📋 Registro T e U nº {idx + 1} - {item.get('local_controle_t_u', 'Novo Local')}", expanded=True):
                # Linha 1: Identificação e Medições
                col_tu_nome, col_tu_t, col_tu_u = st.columns([2, 1, 1])
                with col_tu_nome:
                    local_tu = st.text_input("Nome do local ou equipamento (Estoque/Geladeira):", value=item.get("local_controle_t_u", ""), key=f"tu_nome_{idx}")
                with col_tu_t:
                    t_agora = st.text_input("Temperatura Agora (°C):", value=item.get("t_agora", ""), key=f"tu_t_{idx}")
                with col_tu_u:
                    u_agora = st.text_input("Umidade Agora (%):", value=item.get("u_agora", ""), key=f"tu_u_{idx}")

                # Linha 2: Planilha e Adequação de Valores
                col_tu_plan, col_tu_conf, col_tu_marca = st.columns([1.2, 1.2, 1.5])
                with col_tu_plan:
                    v_plan_tu = item.get("planilha_controle_t_u", "Não avaliado")
                    plan_tu = st.radio("Possui planilha de controle?", opcoes_status, index=mapeamento_status.get(v_plan_tu, 0), key=f"tu_plan_{idx}", horizontal=True)
                with col_tu_conf:
                    v_conf_tu = item.get("conformidade_t_u", "Não avaliado")
                    conf_tu = st.radio("Valores são adequados?", opcoes_status, index=mapeamento_status.get(v_conf_tu, 0), key=f"tu_conf_{idx}", horizontal=True)
                with col_tu_marca:
                    marca_tu = st.text_input("Marca/Modelo do Termo-higrômetro:", value=item.get("marca_modelo_termohigrometro", ""), key=f"tu_marca_{idx}")

                # Linha 3: Calibração
                st.markdown("**Validação Metrológica do Equipamento**")
                v_cal_tu = item.get("possui_calibracao", "Não avaliado")
                cal_tu = st.radio("Avaliação da Calibração:", opcoes_status, index=mapeamento_status.get(v_cal_tu, 0), key=f"tu_cal_{idx}", horizontal=True)

                # Dicionários internos de controle de sub-dados
                dados_cal_tu = {"certificado": "", "empresa": "", "data": None, "validade": None}
                dados_aq_tu = {"nf": "", "data": None, "fornecedor": ""}

                if cal_tu in ["C", "NC"]:
                    # Pergunta se é Calibração padrão ou Aquisição recente
                    v_tipo_tu = item.get("tipo_comprovacao_tu", "Calibração")
                    idx_tipo_tu = 0 if v_tipo_tu == "Calibração" else 1
                    tipo_comp_tu = st.radio("Tipo de comprovação:", opcoes_tipo_comprovacao, index=idx_tipo_tu, key=f"tu_tipo_comp_{idx}", horizontal=True)

                    if tipo_comp_tu == "Calibração":
                        col_c1, col_c2, col_c3, col_c4 = st.columns(4)
                        with col_c1:
                            dados_cal_tu["certificado"] = st.text_input("Nº Certificado:", value=item.get("certificado_calibracao", ""), key=f"tu_cert_{idx}")
                        with col_c2:
                            dados_cal_tu["empresa"] = st.text_input("Empresa Prestadora:", value=item.get("empresa_calibracao", ""), key=f"tu_emp_cal_{idx}")
                        with col_c3:
                            d_c_tu = item.get("data_calibracao", None)
                            dt_c_tu = pd.to_datetime(d_c_tu).date() if d_c_tu and not pd.isna(d_c_tu) else None
                            dados_cal_tu["data"] = st.date_input("Data Calibração:", value=dt_c_tu, key=f"tu_dt_cal_{idx}")
                        with col_c4:
                            v_c_tu = item.get("validade_calibracao", None)
                            dt_v_tu = pd.to_datetime(v_c_tu).date() if v_c_tu and not pd.isna(v_c_tu) else None
                            dados_cal_tu["validade"] = st.date_input("Validade:", value=dt_v_tu, key=f"tu_val_cal_{idx}")
                    else:
                        col_a1, col_a2, col_a3 = st.columns([1, 1, 2])
                        with col_a1:
                            dados_aq_tu["nf"] = st.text_input("Nº Nota Fiscal:", value=item.get("nota_aquisicao", ""), key=f"tu_nf_{idx}")
                        with col_a2:
                            d_a_tu = item.get("data_aquisicao", None)
                            dt_a_tu = pd.to_datetime(d_a_tu).date() if d_a_tu and not pd.isna(d_a_tu) else None
                            dados_aq_tu["data"] = st.date_input("Data Aquisição:", value=dt_a_tu, key=f"tu_dt_aq_{idx}")
                        with col_a3:
                            dados_aq_tu["fornecedor"] = st.text_input("Fornecedor:", value=item.get("empresa_aquisicao", ""), key=f"tu_forn_{idx}")

                # Compacta as informações capturadas deste índice
                lista_tu_atualizada.append({
                    "local_controle_t_u": local_tu, "t_agora": t_agora, "u_agora": u_agora,
                    "planilha_controle_t_u": plan_tu, "conformidade_t_u": conf_tu, "marca_modelo_termohigrometro": marca_tu,
                    "possui_calibracao": cal_tu, "tipo_comprovacao_tu": tipo_comp_tu if cal_tu in ["C", "NC"] else "",
                    "certificado_calibracao": dados_cal_tu["certificado"], "empresa_calibracao": dados_cal_tu["empresa"],
                    "data_calibracao": str(dados_cal_tu["data"]) if dados_cal_tu["data"] else "", "validade_calibracao": str(dados_cal_tu["validade"]) if dados_cal_tu["validade"] else "",
                    "nota_aquisicao": dados_aq_tu["nf"], "data_aquisicao": str(dados_aq_tu["data"]) if dados_aq_tu["data"] else "", "empresa_aquisicao": dados_aq_tu["fornecedor"]
                })

        # Botoes de Controle da Lista T e U
        col_btn_tu_add, col_btn_tu_rem, _ = st.columns([1.5, 1.5, 5])
        with col_btn_tu_add:
            if st.button("➕ Adicionar Área/ Equipamento", use_container_width=True, key="btn_add_tu"):
                # Salva o estado atual das caixas antes de injetar uma nova linha limpa
                lista_temp = []
                for i in range(len(st.session_state["lista_tu_formulario"])):
                    lista_temp.append({
                        "local_controle_t_u": st.session_state[f"tu_nome_{i}"], "t_agora": st.session_state[f"tu_t_{i}"], "u_agora": st.session_state[f"tu_u_{i}"],
                        "planilha_controle_t_u": st.session_state[f"tu_plan_{i}"], "conformidade_t_u": st.session_state[f"tu_conf_{i}"], "marca_modelo_termohigrometro": st.session_state[f"tu_marca_{i}"],
                        "possui_calibracao": st.session_state[f"tu_cal_{i}"], "tipo_comprovacao_tu": st.session_state.get(f"tu_tipo_comp_{i}", "Calibração")
                    })
                lista_temp.append({"local_controle_t_u": "", "t_agora": "", "u_agora": "", "planilha_controle_t_u": "Não avaliado", "conformidade_t_u": "Não avaliado", "marca_modelo_termohigrometro": "", "possui_calibracao": "Não avaliado"})
                st.session_state["lista_tu_formulario"] = lista_temp
                st.rerun()
        with col_btn_tu_rem:
            if len(st.session_state["lista_tu_formulario"]) > 0:
                if st.button("❌ Remover Último T e U", use_container_width=True, key="btn_rem_tu"):
                    st.session_state["lista_tu_formulario"].pop()
                    st.rerun()

        st.divider()

        # ==============================================================================
        # PARTE B: LOOP - CONTROLE APENAS DE TEMPERATURA (T)
        # ==============================================================================
        st.markdown("### 🌡️ Áreas de Temperatura Controlada")
        
        lista_t_atualizada = []
        
        for idx, item in enumerate(st.session_state["lista_t_formulario"]):
            with st.expander(f"📋 Registro T nº {idx + 1} - {item.get('local_controle_t', 'Novo Local')}", expanded=True):
                # Linha 1: Identificação e Medições
                col_t_nome, col_t_val = st.columns([3, 1])
                with col_t_nome:
                    local_t = st.text_input("Nome do local ou equipamento (Estoque/Geladeira):", value=item.get("local_controle_t", ""), key=f"t_nome_{idx}")
                with col_t_val:
                    t_inspecao = st.text_input("Temperatura na Inspeção (°C):", value=item.get("t_inspecao", ""), key=f"t_val_{idx}")
                
                # Linha 2: Planilha e Adequação de Valores
                col_t_plan, col_t_conf, col_t_marca = st.columns([1.2, 1.2, 1.5])
                with col_t_plan:
                    v_plan_t = item.get("planilha_controle_t", "Não avaliado")
                    plan_t = st.radio("Possui planilha de controle?", opcoes_status, index=mapeamento_status.get(v_plan_t, 0), key=f"t_plan_{idx}", horizontal=True)
                with col_t_conf:
                    v_conf_t = item.get("conformidade_t", "Não avaliado")
                    conf_t = st.radio("Valores são adequados?", opcoes_status, index=mapeamento_status.get(v_conf_t, 0), key=f"t_conf_{idx}", horizontal=True)
                with col_t_marca:
                    marca_t = st.text_input("Marca/Modelo do Termômetro:", value=item.get("marca_modelo_termometro", ""), key=f"t_marca_{idx}")
                
                # Linha 3: Calibração
                st.markdown("Validação Metrológica do Equipamento")
                v_cal_t = item.get("possui_calibracao_t", "Não avaliado")
                cal_t = st.radio("Avaliação da Calibração:", opcoes_status, index=mapeamento_status.get(v_cal_t, 0), key=f"t_cal_{idx}", horizontal=True)
                
                dados_cal_t = {"certificado": "", "empresa": "", "data": None, "validade": None}
                dados_aq_t = {"nf": "", "data": None, "fornecedor": ""}
                
                if cal_t in ["C", "NC"]:
                    v_tipo_t = item.get("tipo_comprovacao_t", "Calibração")
                    idx_tipo_t = 0 if v_tipo_t == "Calibração" else 1
                    tipo_comp_t = st.radio("Tipo de comprovação:", opcoes_tipo_comprovacao, index=idx_tipo_t, key=f"t_tipo_comp_{idx}", horizontal=True)
                    
                    if tipo_comp_t == "Calibração":
                        col_tc1, col_tc2, col_tc3, col_tc4 = st.columns(4)
                        with col_tc1:
                            dados_cal_t["certificado"] = st.text_input("Nº Certificado:", value=item.get("certificado_calibracao_t", ""), key=f"t_cert_{idx}")
                        with col_tc2:
                            dados_cal_t["empresa"] = st.text_input("Empresa Prestadora:", value=item.get("empresa_calibracao_t", ""), key=f"t_emp_cal_{idx}")
                        with col_tc3:
                            d_c_t = item.get("data_calibracao_t", None)
                            dt_c_t = pd.to_datetime(d_c_t).date() if d_c_t and not pd.isna(d_c_t) else None
                            dados_cal_t["data"] = st.date_input("Data Calibração:", value=dt_c_t, key=f"t_dt_cal_{idx}")
                        with col_tc4:
                            v_c_t = item.get("validade_calibracao_t", None)
                            dt_v_t = pd.to_datetime(v_c_t).date() if v_c_t and not pd.isna(v_c_t) else None
                            dados_cal_t["validade"] = st.date_input("Validade:", value=dt_v_t, key=f"t_val_cal_{idx}")
                    else:
                        col_ta1, col_ta2, col_ta3 = st.columns([1, 1, 2])
                        with col_ta1:
                            dados_aq_t["nf"] = st.text_input("Nº Nota Fiscal:", value=item.get("nota_aquisicao_t", ""), key=f"t_nf_{idx}")
                        with col_ta2:
                            d_a_t = item.get("data_aquisicao_t", None)
                            dt_a_t = pd.to_datetime(d_a_t).date() if d_a_t and not pd.isna(d_a_t) else None
                            dados_aq_t["data"] = st.date_input("Data Aquisição:", value=dt_a_t, key=f"t_dt_aq_{idx}")
                        with col_ta3:
                            dados_aq_t["fornecedor"] = st.text_input("Fornecedor:", value=item.get("empresa_aquisicao_t", ""), key=f"t_forn_{idx}")
                
                # Compacta as informações capturadas deste índice
                lista_t_atualizada.append({
                    "local_controle_t": local_t, 
                    "t_inspecao": t_inspecao,
                    "planilha_controle_t": plan_t, 
                    "conformidade_t": conf_t, 
                    "marca_modelo_termometro": marca_t,
                    "possui_calibracao_t": cal_t, 
                    "tipo_comprovacao_t": tipo_comp_t if cal_t in ["C", "NC"] else "",
                    "certificado_calibracao_t": dados_cal_t["certificado"], 
                    "empresa_calibracao_t": dados_cal_t["empresa"],
                    "data_calibracao_t": str(dados_cal_t["data"]) if dados_cal_t["data"] else "", 
                    "validade_calibracao_t": str(dados_cal_t["validade"]) if dados_cal_t["validade"] else "",
                    "nota_aquisicao_t": dados_aq_t["nf"], 
                    "data_aquisicao_t": str(dados_aq_t["data"]) if dados_aq_t["data"] else "", 
                    "empresa_aquisicao_t": dados_aq_t["fornecedor"]
                })

        # Botoes de Controle da Lista T 
        col_btn_t_add, col_btn_t_rem, _ = st.columns([1.5, 1.5, 5])
        with col_btn_t_add:
            if st.button("➕ Adicionar Área/ Equipamento", use_container_width=True, key="btn_add_t"):
                lista_temp_t = []
                for i in range(len(st.session_state["lista_t_formulario"])):
                    lista_temp_t.append({
                        "local_controle_t": st.session_state[f"t_nome_{i}"], 
                        "t_inspecao": st.session_state[f"t_val_{i}"],
                        "planilha_controle_t": st.session_state[f"t_plan_{i}"], 
                        "conformidade_t": st.session_state[f"t_conf_{i}"], 
                        "marca_modelo_termometro": st.session_state[f"t_marca_{i}"],
                        "possui_calibracao_t": st.session_state[f"t_cal_{i}"], 
                        "tipo_comprovacao_t": st.session_state.get(f"t_tipo_comp_{i}", "Calibração")
                    })
                lista_temp_t.append({"local_controle_t": "", "t_inspecao": "", "planilha_controle_t": "Não avaliado", "conformidade_t": "Não avaliado", "marca_modelo_termometro": "", "possui_calibracao_t": "Não avaliado"})
                st.session_state["lista_t_formulario"] = lista_temp_t
                st.rerun()
                
        with col_btn_t_rem:
            if len(st.session_state["lista_t_formulario"]) > 0:
                if st.button("❌ Remover Último Termômetro", use_container_width=True, key="btn_rem_t"):
                    st.session_state["lista_t_formulario"].pop()
                    st.rerun()

        # Vincula o empacotamento das duas sublistas em um único objeto JSON estruturado para a coluna JSONB do banco
        dados_inspecao_atual["locais_controle_geladeiras_json"] = {
            "lista_tu": lista_tu_atualizada,
            "lista_t": lista_t_atualizada
        }


    # ==============================================================================
    # TAB 4: COMERCIALIZAÇÃO (PRODUTOS E MEDICAMENTOS CONTROLADOS)
    # ==============================================================================
    with tab_comercializacao:
        st.markdown("## 📦 Seção: Comercialização de Produtos e Controle Especial")

        # --- 4.1 PRODUTOS GERAIS ---
        st.markdown("### 🔍 Comercialização de Produtos Gerais")
        
        col_emb, col_perm = st.columns(2)
        
        with col_emb:
            st.markdown("**A embalagem e rotulagem dos produtos comercializados verificados estão conforme a legislação vigente?**")
            val_emb = historico.get("embalagem_produtos", "Não avaliado")
            dados_inspecao_atual["embalagem_produtos"] = st.radio(
                "Avaliação Embalagens:", opcoes_status, index=mapeamento_status.get(val_emb, 0), key="rad_embalagem_produtos", horizontal=True
            )
            if dados_inspecao_atual["embalagem_produtos"] == "NC":
                val_lp_nc = historico.get("lista_produtos_nc", "")
                dados_inspecao_atual["lista_produtos_nc"] = st.text_area(
                    "Lista de produtos não conformes encontrados:", 
                    value=str(val_lp_nc) if val_lp_nc and str(val_lp_nc) != "None" else "",
                    key="ta_lista_produtos_nc"
                )
            else:
                dados_inspecao_atual["lista_produtos_nc"] = None

        with col_perm:
            st.markdown("**Os produtos disponíveis para a venda são permitidos pela legislação para a atividade do estabelecimento?**")
            val_perm = historico.get("permitidos_venda_produtos", "Não avaliado")
            dados_inspecao_atual["permitidos_venda_produtos"] = st.radio(
                "Avaliação Produtos Permitidos:", opcoes_status, index=mapeamento_status.get(val_perm, 0), key="rad_permitidos_venda_produtos", horizontal=True
            )
            if dados_inspecao_atual["permitidos_venda_produtos"] == "NC":
                val_lp_np = historico.get("lista_produtos_np", "")
                dados_inspecao_atual["lista_produtos_np"] = st.text_area(
                    "Lista de produtos não permitidos encontrados:", 
                    value=str(val_lp_np) if val_lp_np and str(val_lp_np) != "None" else "",
                    key="ta_lista_produtos_np"
                )
            else:
                dados_inspecao_atual["lista_produtos_np"] = None

        st.divider()

        # --- 4.2 MEDICAMENTOS SUJEITOS A CONTROLE ESPECIAL (PORTARIA 344/98) ---
        st.markdown("### 🔒 Controle Especial (Portaria SVS nº 344/98)")
        
        # Pergunta norteadora: Inicia em "Não avaliado" se for empresa nova
        val_disp_ce = historico.get("dispensa_controle_especial", "Não avaliado")
        idx_disp_ce = mapeamento_sim_nao.get(val_disp_ce, 0)
        
        st.markdown("**O estabelecimento dispensa Medicamentos sujeitos à Controle Especial (Portaria SVS n° 344/98)?**")
        dados_inspecao_atual["dispensa_controle_especial"] = st.radio(
            "Dispensa Controle Especial?", opcoes_sim_nao, index=idx_disp_ce, key="rad_dispensa_controle_especial", horizontal=True
        )

        # LÓGICA DINÂMICA: Abre todas as perguntas internas apenas se clicar em "Sim"
        if dados_inspecao_atual["dispensa_controle_especial"] == "Sim":
            st.markdown("#### 📂 Requisitos Operacionais e Escrituração")
            
            col_ce1, col_ce2 = st.columns(2)
            with col_ce1:
                st.markdown("**Possui local de acesso restrito (trancado à chave) para armazenamento de medicamentos sujeitos a controle especial?**")
                val_loc_ce = historico.get("local_armazenamento_controle_especial", "Não avaliado")
                dados_inspecao_atual["local_armazenamento_controle_especial"] = st.radio(
                    "Local Acesso Restrito:", opcoes_status, index=mapeamento_status.get(val_loc_ce, 0), key="rad_local_armazenamento_controle_especial", horizontal=True
                )
            with col_ce2:
                st.markdown("**Possui área segregada para guarda temporária de medicamentos controlados pela Portaria 344/98 a serem inutilizados?**")
                val_seg_ce = historico.get("area_segregada_inutilizados_mc", "Não avaliado")
                dados_inspecao_atual["area_segregada_inutilizados_mc"] = st.radio(
                    "Área Segregada Inutilizados:", opcoes_status, index=mapeamento_status.get(val_seg_ce, 0), key="rad_area_segregada_inutilizados_mc", horizontal=True
                )

            st.divider()

            # --- SUB-BLOCO: RETINÓICOS (LISTA C2) ---
            st.markdown("##### 💊 Medicamentos Retinóicos (Lista C2)")
            st.markdown("**Dispensa medicamentos retinóicos (lista C2) de forma regular?**")
            val_ret = historico.get("dispensa_retinoicos", "Não avaliado")
            dados_inspecao_atual["dispensa_retinoicos"] = st.radio(
                "Avaliação Retinóicos:", opcoes_status, index=mapeamento_status.get(val_ret, 0), key="rad_dispensa_retinoicos", horizontal=True
            )
            
            if dados_inspecao_atual["dispensa_retinoicos"] in ["C", "NC"]:
                col_r1, col_r2, col_r3 = st.columns(3)
                with col_r1:
                    val_dt_aut = historico.get("data_autorizacao", None)
                    dt_p_aut = pd.to_datetime(val_dt_aut).date() if val_dt_aut and not pd.isna(val_dt_aut) else None
                    dados_inspecao_atual["data_autorizacao"] = st.date_input("Data da Autorização:", value=dt_p_aut, key="di_data_autorizacao")
                with col_r2:
                    val_ed_visa = historico.get("n_edital_visa_autorizacao", "")
                    dados_inspecao_atual["n_edital_visa_autorizacao"] = st.text_input(
                        "Número Edital VISA:", value=str(val_ed_visa) if val_ed_visa and str(val_ed_visa) != "None" else "", key="ti_n_edital_visa_autorizacao"
                    )
                with col_r3:
                    val_sis_aut = historico.get("sistema_informatizado_autorizacao", "")
                    dados_inspecao_atual["sistema_informatizado_autorizacao"] = st.text_input(
                        "Qual Sistema Informatizado Utilizado?", value=str(val_sis_aut) if val_sis_aut and str(val_sis_aut) != "None" else "", key="ti_sistema_informatizado_autorizacao"
                    )
                
                opcoes_escrituracao = [
                    "Não informado", 
                    "Livro Informatizado com termo de abertura e encerramento pela VISA", 
                    "Informações no sistema sem lavratura de livro digital", 
                    "Livros físicos", 
                    "Observação"
                ]
                map_escrituracao = {v: i for i, v in enumerate(opcoes_escrituracao)}
                val_esc_tipo = historico.get("tipo_escrituracao", "Não informado")
                dados_inspecao_atual["tipo_escrituracao"] = st.selectbox(
                    "Tipo de Escrituração:", options=opcoes_escrituracao, index=map_escrituracao.get(val_esc_tipo, 0), key="sb_tipo_escrituracao"
                )
            else:
                dados_inspecao_atual["data_autorizacao"] = None
                dados_inspecao_atual["n_edital_visa_autorizacao"] = None
                dados_inspecao_atual["sistema_informatizado_autorizacao"] = None
                dados_inspecao_atual["tipo_escrituracao"] = None

            st.divider()

            # --- SUB-BLOCO: LISTA A ---
            st.markdown("##### 🟡 Medicamentos da Lista A (Receita Amarela)")
            st.markdown("**Dispensa medicamentos da lista A (receita amarela) de forma regular?**")
            val_list_a = historico.get("dispensa_lista_a", "Não avaliado")
            dados_inspecao_atual["dispensa_lista_a"] = st.radio(
                "Avaliação Lista A:", opcoes_status, index=mapeamento_status.get(val_list_a, 0), key="rad_dispensa_lista_a", horizontal=True
            )
            if dados_inspecao_atual["dispensa_lista_a"] in ["C", "NC"]:
                val_dt_rmnra = historico.get("data_protocolo_rmnra", None)
                dt_p_rmnra = pd.to_datetime(val_dt_rmnra).date() if val_dt_rmnra and not pd.isna(val_dt_rmnra) else None
                dados_inspecao_atual["data_protocolo_rmnra"] = st.date_input("Data do último protocolo da RMNRA entregue à VISA:", value=dt_p_rmnra, key="di_data_protocolo_rmnra")
                
                val_obs_a = historico.get("obs_lista_a", "") # criará campo livre complementar seguro
                dados_inspecao_atual["obs_lista_a"] = st.text_input("Observação (Lista A):", value=str(val_obs_a) if val_obs_a and str(val_obs_a) != "None" else "", key="ti_obs_lista_a")
            else:
                dados_inspecao_atual["data_protocolo_rmnra"] = None
                dados_inspecao_atual["obs_lista_a"] = None

            st.divider()

            # --- SUB-BLOCO: LISTA B2 ---
            st.markdown("##### 🧪 Medicamentos da Lista B2 (Anorexígenos)")
            st.markdown("**Dispensa medicamentos da lista B2 (anorexígenos) de forma regular?**")
            val_list_b2 = historico.get("dispensa_lista_b2", "Não avaliado")
            dados_inspecao_atual["dispensa_lista_b2"] = st.radio(
                "Avaliação Lista B2:", opcoes_status, index=mapeamento_status.get(val_list_b2, 0), key="rad_dispensa_lista_b2", horizontal=True
            )
            if dados_inspecao_atual["dispensa_lista_b2"] in ["C", "NC"]:
                val_dt_rmnrb2 = historico.get("data_protocolo_rmnrb2", None)
                dt_p_rmnrb2 = pd.to_datetime(val_dt_rmnrb2).date() if val_dt_rmnrb2 and not pd.isna(val_dt_rmnrb2) else None
                dados_inspecao_atual["data_protocolo_rmnrb2"] = st.date_input("Data do último protocolo da RMNRB2 entregue à VISA:", value=dt_p_rmnrb2, key="di_data_protocolo_rmnrb2")
                val_obs_b2 = historico.get("obs_lista_b2", "")
                dados_inspecao_atual["obs_lista_b2"] = st.text_input("Observação (Lista B2):", value=str(val_obs_b2) if val_obs_b2 and str(val_obs_b2) != "None" else "", key="ti_obs_lista_b2")
            else:
                dados_inspecao_atual["data_protocolo_rmnrb2"] = None
                dados_inspecao_atual["obs_lista_b2"] = None

            st.divider()

            # --- SUB-BLOCO: BALANÇOS BMPO ---
            st.markdown("##### 📊 Checagem de Envio de Balanços (BMPO) para VISA")
            col_bmpo1, col_bmpo2 = st.columns(2)
            with col_bmpo1:
                val_bmpo_t = historico.get("data_envio_bmpo_trimestral", None)
                dt_p_bmpo_t = pd.to_datetime(val_bmpo_t).date() if val_bmpo_t and not pd.isna(val_bmpo_t) else None
                dados_inspecao_atual["data_envio_bmpo_trimestral"] = st.date_input("Trimestral / Data último protocolo:", value=dt_p_bmpo_t, key="di_data_envio_bmpo_trimestral")
            with col_bmpo2:
                val_bmpo_a = historico.get("data_envio_bmpo_anual", None)
                dt_p_bmpo_a = pd.to_datetime(val_bmpo_a).date() if val_bmpo_a and not pd.isna(val_bmpo_a) else None
                dados_inspecao_atual["data_envio_bmpo_anual"] = st.date_input("Anual / Data último protocolo:", value=dt_p_bmpo_a, key="di_data_envio_bmpo_anual")
            
            val_obs_bmpo = historico.get("obs_bmpo", "")
            dados_inspecao_atual["obs_bmpo"] = st.text_input("Observação (Balanços BMPO):", value=str(val_obs_bmpo) if val_obs_bmpo and str(val_obs_bmpo) != "None" else "", key="ti_obs_bmpo")

            st.divider()

            
            # --- MOTOR DE CONFERÊNCIA DE ESTOQUE (MEDICAMENTOS CONTROLADOS E ANTIMICROBIANOS)
            
            st.markdown("### 📋 Conferência de Estoque")
            
            # Recupera as listas estruturadas salvas nas colunas JSONB correspondentes do Supabase
            hist_mc = historico.get("medicamentos_controlados_verificados_json", [])
            if not isinstance(hist_mc, list): 
                hist_mc = []
            hist_am = historico.get("antimicrobianos_verificados_json", [])
            if not isinstance(hist_am, list): 
                hist_am = []

            if "lista_mc_formulario" not in st.session_state or st.session_state.get("ultimo_cnpj_visto_comercial") != cnpj_estabelecimento:
                st.session_state["ultimo_cnpj_visto_comercial"] = cnpj_estabelecimento
                st.session_state["lista_mc_formulario"] = hist_mc
                st.session_state["lista_am_formulario"] = hist_am

            # --- LOOP A: MEDICAMENTOS CONTROLADOS ---
            st.markdown("##### 🚨 Medicamentos Controlados Verificados")
            lista_mc_atualizada = []
            for idx, item_mc in enumerate(st.session_state["lista_mc_formulario"]):
                with st.expander(f"📦 Controlado nº {idx + 1} - {item_mc.get('principio_ativo_mc', 'Novo Ativo')}", expanded=True):
                    c1, c2, c3 = st.columns([2, 2, 1.5])
                    with c1: 
                        p_ativo_mc = st.text_input("Princípio Ativo:", value=item_mc.get("principio_ativo_mc", ""), key=f"mc_pa_{idx}")
                    with c2: 
                        apres_mc = st.text_input("Apresentação / Concentração:", value=item_mc.get("apresentacao_mc", ""), key=f"mc_ap_{idx}")
                    with c3: 
                        fab_mc = st.text_input("Fabricante:", value=item_mc.get("fabricante_mc", ""), key=f"mc_fa_{idx}")
                    
                    c4, c5, c6, c7 = st.columns([1.5, 1.5, 1, 1])
                    with c4: 
                        lote_mc = st.text_input("Lote:", value=item_mc.get("lote_mc", ""), key=f"mc_lo_{idx}")
                    with c5:
                        v_v_mc = item_mc.get("validade_mc", "")
                        val_mc = st.text_input("Validade:", value=v_v_mc, key=f"mc_va_{idx}")
                    with c6: 
                        q_fis_mc = st.text_input("Qtd (Estoque Físico):", value=item_mc.get("quantidade_fisico_mc", ""), key=f"mc_qf_{idx}")
                    with c7: 
                        q_sis_mc = st.text_input("Qtd (Sistema):", value=item_mc.get("quantidade_sistema_mc", ""), key=f"mc_qs_{idx}")
                    
                    obs_mc = st.text_input("Observação / Desvio encontrado:", value=item_mc.get("obs_mc", ""), key=f"mc_ob_{idx}")
                    
                    lista_mc_atualizada.append({
                        "principio_ativo_mc": p_ativo_mc, "apresentacao_mc": apres_mc, "fabricante_mc": fab_mc,
                        "lote_mc": lote_mc, "validade_mc": val_mc, "quantidade_fisico_mc": q_fis_mc, "quantidade_sistema_mc": q_sis_mc, "obs_mc": obs_mc
                    })

            col_btn_mc_add, col_btn_mc_rem, _ = st.columns([1.5, 1.5, 5])
            with col_btn_mc_add:
                if st.button("➕ Adicionar Medicamento Controlado", use_container_width=True, key="btn_add_mc"):
                    l_temp = []
                    for i in range(len(st.session_state["lista_mc_formulario"])):
                        l_temp.append({
                            "principio_ativo_mc": st.session_state[f"mc_pa_{i}"], "apresentacao_mc": st.session_state[f"mc_ap_{i}"], "fabricante_mc": st.session_state[f"mc_fa_{i}"],
                            "lote_mc": st.session_state[f"mc_lo_{i}"], "validade_mc": st.session_state[f"mc_va_{i}"], "quantidade_fisico_mc": st.session_state[f"mc_qf_{i}"], "quantidade_sistema_mc": st.session_state[f"mc_qs_{i}"], "obs_mc": st.session_state[f"mc_ob_{i}"]
                        })
                    l_temp.append({"principio_ativo_mc": "", "apresentacao_mc": "", "fabricante_mc": "", "lote_mc": "", "validade_mc": "", "quantidade_fisico_mc": "", "quantidade_sistema_mc": "", "obs_mc": ""})
                    st.session_state["lista_mc_formulario"] = l_temp
                    st.rerun()
            with col_btn_mc_rem:
                if len(st.session_state["lista_mc_formulario"]) > 0:
                    if st.button("❌ Remover Último Controlado", use_container_width=True, key="btn_rem_mc"):
                        st.session_state["lista_mc_formulario"].pop()
                        st.rerun()

            st.divider()
           
            # --- LOOP B: ANTIMICROBIANOS ---
            st.markdown("##### 🧫 Antimicrobianos Verificados")
            lista_am_atualizada = []
            
            for idx, item_am in enumerate(st.session_state["lista_am_formulario"]):
                with st.expander(f"💊 Antimicrobiano nº {idx + 1} - {item_am.get('principio_ativo_am', 'Novo Ativo')}", expanded=True):
                    a1, a2, a3 = st.columns([2, 2, 1.5])
                    with a1: 
                        p_ativo_am = st.text_input("Princípio Ativo:", value=item_am.get("principio_ativo_am", ""), key=f"am_pa_{idx}")
                    with a2: 
                        apres_am = st.text_input("Apresentação / Concentração:", value=item_am.get("apresentacao_am", ""), key=f"am_ap_{idx}")
                    with a3: 
                        fab_am = st.text_input("Fabricante:", value=item_am.get("fabricante_am", ""), key=f"am_fa_{idx}")
                    
                    a4, a5, a6, a7 = st.columns([1.5, 1.5, 1, 1])
                    with a4: 
                        lote_am = st.text_input("Lote:", value=item_am.get("lote_am", ""), key=f"am_lo_{idx}")
                    with a5:
                        v_v_am = item_am.get("validade_am", "")
                        val_am = st.text_input("Validade:", value=v_v_am, key=f"am_va_{idx}")
                    with a6: 
                        q_fis_am = st.text_input("Qtd (Estoque Físico):", value=item_am.get("quantidade_fisico_am", ""), key=f"am_qf_{idx}")
                    with a7: 
                        q_sis_am = st.text_input("Qtd (Sistema):", value=item_am.get("quantidade_sistema_am", ""), key=f"am_qs_{idx}")
                    
                    obs_am = st.text_input("Observação / Desvio encontrado:", value=item_am.get("obs_am", ""), key=f"am_ob_{idx}")
                    
                    lista_am_atualizada.append({
                        "principio_ativo_am": p_ativo_am, 
                        "apresentacao_am": apres_am, 
                        "fabricante_am": fab_am,
                        "lote_am": lote_am, 
                        "validade_am": val_am, 
                        "quantidade_fisico_am": q_fis_am, 
                        "quantidade_sistema_am": q_sis_am, 
                        "obs_am": obs_am
                    })

            # Botoes de Controle da Lista AM (Ficam fora do loop for)
            col_btn_am_add, col_btn_am_rem, _ = st.columns([1.5, 1.5, 5])
            with col_btn_am_add:
                if st.button("➕ Adicionar Antimicrobiano", use_container_width=True, key="btn_add_am"):
                    l_temp_am = []
                    for i in range(len(st.session_state["lista_am_formulario"])):
                        l_temp_am.append({
                            "principio_ativo_am": st.session_state[f"am_pa_{i}"], 
                            "apresentacao_am": st.session_state[f"am_ap_{i}"], 
                            "fabricante_am": st.session_state[f"am_fa_{i}"],
                            "lote_am": st.session_state[f"am_lo_{i}"], 
                            "validade_am": st.session_state[f"am_va_{i}"], 
                            "quantidade_fisico_am": st.session_state[f"am_qf_{i}"], 
                            "quantidade_sistema_am": st.session_state[f"am_qs_{i}"], 
                            "obs_am": st.session_state[f"am_ob_{i}"]
                        })
                    l_temp_am.append({"principio_ativo_am": "", "apresentacao_am": "", "fabricante_am": "", "lote_am": "", "validade_am": "", "quantidade_fisico_am": "", "quantidade_sistema_am": "", "obs_am": ""})
                    st.session_state["lista_am_formulario"] = l_temp_am
                    st.rerun()
                    
            with col_btn_am_rem:
                if len(st.session_state["lista_am_formulario"]) > 0:
                    if st.button("❌ Remover Último Antimicrobiano", use_container_width=True, key="btn_rem_am"):
                        st.session_state["lista_am_formulario"].pop()
                        st.rerun()

            # Vincula o empacotamento das listas dinâmicas de conferência de estoque para o dicionário mestre de envio
            dados_inspecao_atual["medicamentos_controlados_verificados_json"] = lista_mc_atualizada
            dados_inspecao_atual["antimicrobianos_verificados_json"] = lista_am_atualizada
            
            st.divider()

            # --- SUB-BLOCO: INUTILIZAÇÃO E SNGPC ---
            st.markdown("##### 📑 Inutilização de Produtos e SNGPC")
            col_inut, col_sngpc = st.columns(2)
            
            with col_inut:
                st.markdown("Possui termo de Inutilização de Medicamentos conforme?")
                val_t_inut = historico.get("possui_termo_inutilizacao_visa", "Não avaliado")
                dados_inspecao_atual["possui_termo_inutilizacao_visa"] = st.radio(
                    "Termo Inutilização:", opcoes_status, index=mapeamento_status.get(val_t_inut, 0), key="rad_possui_termo_inutilizacao_visa", horizontal=True
                )
                if dados_inspecao_atual["possui_termo_inutilizacao_visa"] in ["C", "NC"]:
                    val_dt_inut = historico.get("data_termo_inutilizacao_visa", None)
                    dt_p_inut = pd.to_datetime(val_dt_inut).date() if val_dt_inut and not pd.isna(val_dt_inut) else None
                    dados_inspecao_atual["data_termo_inutilizacao_visa"] = st.date_input("Data último protocolo entregue na VISA:", value=dt_p_inut, key="di_data_termo_inutilizacao_visa")
                else:
                    dados_inspecao_atual["data_termo_inutilizacao_visa"] = None
                    
            with col_sngpc:
                st.markdown("Realiza transmissão ao SNGPC de maneira regular?")
                val_trans_s = historico.get("faz_transmissao_sngpc", "Não avaliado")
                dados_inspecao_atual["faz_transmissao_sngpc"] = st.radio(
                    "Transmissão SNGPC:", opcoes_status, index=mapeamento_status.get(val_trans_s, 0), key="rad_faz_transmissao_sngpc", horizontal=True
                )
                if dados_inspecao_atual["faz_transmissao_sngpc"] in ["C", "NC"]:
                    val_dt_sngpc = historico.get("data__transmissao_sngpc", None)
                    dt_p_sngpc = pd.to_datetime(val_dt_sngpc).date() if val_dt_sngpc and not pd.isna(val_dt_sngpc) else None
                    dados_inspecao_atual["data__transmissao_sngpc"] = st.date_input("Data do último envio:", value=dt_p_sngpc, key="di_data__transmissao_sngpc")
                    
                    val_per_sngpc = historico.get("periodo_transmissao_sngpc", "")
                    dados_inspecao_atual["periodo_transmissao_sngpc"] = st.text_input("Período da última transmissão enviada:", value=str(val_per_sngpc) if val_per_sngpc and str(val_per_sngpc) != "None" else "", key="ti_periodo_transmissao_sngpc")
                    
                    val_vld_sngpc = historico.get("validacao_transmissao_sngpc", "")
                    dados_inspecao_atual["validacao_transmissao_sngpc"] = st.text_input("Data da Validação (Retorno Sistema):", value=str(val_vld_sngpc) if val_vld_sngpc and str(val_vld_sngpc) != "None" else "", key="ti_validacao_transmissao_sngpc")
                else:
                    dados_inspecao_atual["data__transmissao_sngpc"] = None
                    dados_inspecao_atual["periodo_transmissao_sngpc"] = None
                    dados_inspecao_atual["validacao_transmissao_sngpc"] = None
        else:
            # Caso o fiscal marque "Não" na pergunta principal, zera de forma forçada todas as colunas do sub-bloco da Portaria 344 para o banco
            chaves_zerar = [
                "local_armazenamento_controle_especial", "area_segregada_inutilizados_mc", "dispensa_retinoicos",
                "data_autorizacao", "n_edital_visa_autorizacao", "sistema_informatizado_autorizacao", "tipo_escrituracao",
                "dispensa_lista_a", "data_protocolo_rmnra", "obs_lista_a", "dispensa_lista_b2", "data_protocolo_rmnrb2",
                "obs_lista_b2", "data_envio_bmpo_trimestral", "data_envio_bmpo_anual", "obs_bmpo",
                "possui_termo_inutilizacao_visa", "data_termo_inutilizacao_visa", "faz_transmissao_sngpc",
                "data__transmissao_sngpc", "periodo_transmissao_sngpc", "validacao_transmissao_sngpc"
            ]
            for chave in chaves_zerar: 
                dados_inspecao_atual[chave] = None
            dados_inspecao_atual["medicamentos_controlados_verificados_json"] = []
            dados_inspecao_atual["antimicrobianos_verificados_json"] = []
            
        st.divider()

    # ==============================================================================
    # TAB 5: ÁREAS FÍSICAS (SUB-ABAS DINÂMICAS)
    # ==============================================================================
    with tab_areas:
        st.markdown("## 🏢 Seção: Avaliação da Área Física")
        
        # --- DICIONÁRIO DE MAPEAMENTO DAS ÁREAS ---
        # Conecta o nome amigável da tela com a coluna de flag "possui_xxx" do Supabase
        mapeamento_areas_config = {
            "Recebimento": "possui_recebimento",
            "Estoque": "possui_estoque",
            "Sala de Serviços": "possui_sala_servicos",
            "DML": "possui_dml",
            "Sanitário Funcionários": "possui_sanitarios_funcionarios",
            "Sanitário Clientes": "possui_sanitarios_clientes",
            "Copa": "possui_copa",
            "ADM": "possui_adm"
        }
        
        # --- RECUPERAÇÃO DO HISTÓRICO DE ÁREAS ATIVAS ---
        # Analisa a última inspeção e descobre quais áreas estavam marcadas como "Sim"
        areas_ativas_historico = []
        for nome_amigavel, coluna_flag in mapeamento_areas_config.items():
            if historico.get(coluna_flag, "Não") == "Sim":
                areas_ativas_historico.append(nome_amigavel)
                
        st.markdown("**Selecione abaixo quais setores físicos estão presentes e serão avaliados nesta inspeção:**")
        
        # O multiselect atua como as bolinhas de marcação fixas no topo da página
        areas_selecionadas = st.multiselect(
            "Adicionar área:",
            options=list(mapeamento_areas_config.keys()),
            default=areas_ativas_historico,
            key="ms_areas_fisicas_selecionadas"
        )
        
        # --- ATRIBUIÇÃO AUTOMÁTICA DAS FLAGS DE PRESENÇA (SIM / NÃO) ---
        for nome_amigavel, coluna_flag in mapeamento_areas_config.items():
            if nome_amigavel in areas_selecionadas:
                dados_inspecao_atual[coluna_flag] = "Sim"
            else:
                dados_inspecao_atual[coluna_flag] = "Não"

        st.divider()

        # --- MOTOR DE GERAÇÃO DAS SUB-ABAS DINÂMICAS ---
        if not areas_selecionadas:
            st.info("💡 Nenhuma área física selecionada. Marque os setores acima para abrir os questionários estruturais.")
        else:
            # Cria os objetos das sub-abas no Streamlit baseado apenas no que o fiscal marcou
            sub_tabs = st.tabs(areas_selecionadas)
            
            # Mapeia qual objeto de sub-aba pertence a cada nome para renderizarmos na ordem correta
            map_abas_geradas = {nome: sub_tabs[i] for i, nome in enumerate(areas_selecionadas)}
            
            # --------------------------------------------------------------------------
            # SUB-ABA: RECEBIMENTO
            # --------------------------------------------------------------------------
            if "Recebimento" in map_abas_geradas:
                with map_abas_geradas["Recebimento"]:
                    st.markdown("### 📥 Setor de Recebimento de Mercadorias")
                    
                    st.markdown("**O recebimento possui condições adequadas de estrutura?**")
                    val_est_rec = historico.get("condicoes_estruturais_recebimento", "Não avaliado")
                    dados_inspecao_atual["condicoes_estruturais_recebimento"] = st.radio(
                        "Estrutura Recebimento:", opcoes_status, index=mapeamento_status.get(val_est_rec, 0), key="rad_condicoes_estruturais_recebimento", horizontal=True, label_visibility="collapsed" 
                    )
                    st.write("")
                    
                    st.markdown("**O recebimento possui condições higiênico-sanitárias adequadas?**")
                    val_hig_rec = historico.get("condicoes_higienicas_recebimento", "Não avaliado")
                    dados_inspecao_atual["condicoes_higienicas_recebimento"] = st.radio(
                        "Higiene Recebimento:", opcoes_status, index=mapeamento_status.get(val_hig_rec, 0), key="rad_condicoes_higienicas_recebimento", horizontal=True, label_visibility="collapsed"
                    )
                    st.write("")
                    
                    st.markdown("**O estabelecimento realiza o devido controle das mercadorias recebidas, com registros?**")
                    val_ctrl_rec = historico.get("controle_recebimento", "Não avaliado")
                    dados_inspecao_atual["controle_recebimento"] = st.radio(
                        "Controle Recebimento:", opcoes_status, index=mapeamento_status.get(val_ctrl_rec, 0), key="rad_controle_recebimento", horizontal=True, label_visibility="collapsed"
                    )
            
            # --------------------------------------------------------------------------
            # SUB-ABA: ESTOQUE
            # --------------------------------------------------------------------------
            if "Estoque" in map_abas_geradas:
                with map_abas_geradas["Estoque"]:
                    st.markdown("### 📦 Setor de Armazenamento / Estoque de Produtos")
                    
                    # Organização em duas colunas para aproveitar o espaço vertical
                    col_est1, col_est2 = st.columns(2)
                    
                    with col_est1:
                        st.markdown("**O estoque possui condições adequadas de estrutura?**")
                        val_est_est = historico.get("condicoes_estruturais_estoque", "Não avaliado")
                        dados_inspecao_atual["condicoes_estruturais_estoque"] = st.radio(
                            "Estrutura Estoque:", opcoes_status, index=mapeamento_status.get(val_est_est, 0), key="rad_condicoes_estruturais_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**O piso é liso, lavável e impermeável?**")
                        val_piso = historico.get("piso_estoque", "Não avaliado")
                        dados_inspecao_atual["piso_estoque"] = st.radio(
                            "Piso Estoque:", opcoes_status, index=mapeamento_status.get(val_piso, 0), key="rad_piso_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**As paredes são lisas, laváveis e impermeáveis?**")
                        val_parede = historico.get("parede_estoque", "Não avaliado")
                        dados_inspecao_atual["parede_estoque"] = st.radio(
                            "Parede Estoque:", opcoes_status, index=mapeamento_status.get(val_parede, 0), key="rad_parede_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**O teto é liso, lavável e impermeável?**")
                        val_teto = historico.get("teto_estoque", "Não avaliado")
                        dados_inspecao_atual["teto_estoque"] = st.radio(
                            "Teto Estoque:", opcoes_status, index=mapeamento_status.get(val_teto, 0), key="rad_teto_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**O estoque possui condições higiênico-sanitárias adequadas?**")
                        val_hig_est = historico.get("condicoes_higienicas_estoque", "Não avaliado")
                        dados_inspecao_atual["condicoes_higienicas_estoque"] = st.radio(
                            "Higiene Estoque:", opcoes_status, index=mapeamento_status.get(val_hig_est, 0), key="rad_condicoes_higienicas_estoque", horizontal=True, label_visibility="collapsed"
                        )

                    with col_est2:
                        st.markdown("**Há devida proteção contra entrada de pragas e insetos?**")
                        val_prag_est = historico.get("protecao_pragas_estoque", "Não avaliado")
                        dados_inspecao_atual["protecao_pragas_estoque"] = st.radio(
                            "Pragas Estoque:", opcoes_status, index=mapeamento_status.get(val_prag_est, 0), key="rad_protecao_pragas_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A ventilação é adequada?**")
                        val_vent_est = historico.get("ventilacao_estoque", "Não avaliado")
                        dados_inspecao_atual["ventilacao_estoque"] = st.radio(
                            "Ventilação Estoque:", opcoes_status, index=mapeamento_status.get(val_vent_est, 0), key="rad_ventilacao_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A iluminação é adequada?**")
                        val_ilum_est = historico.get("iluminacao_estoque", "Não avaliado")
                        dados_inspecao_atual["iluminacao_estoque"] = st.radio(
                            "Iluminação Estoque:", opcoes_status, index=mapeamento_status.get(val_ilum_est, 0), key="rad_iluminacao_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**As estruturas de armazenamento são afastadas de piso, parede e teto?**")
                        val_afast = historico.get("afastamento_estrutura_estoque", "Não avaliado")
                        dados_inspecao_atual["afastamento_estrutura_estoque"] = st.radio(
                            "Afastamento Estrutura:", opcoes_status, index=mapeamento_status.get(val_afast, 0), key="rad_afastamento_estrutura_estoque", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Há área destinada para produtos não conformes, devidamente identificada?**")
                        val_prod_nc = historico.get("produtos_nao_conformes_estoque", "Não avaliado")
                        dados_inspecao_atual["produtos_nao_conformes_estoque"] = st.radio(
                            "Produtos NC Estoque:", opcoes_status, index=mapeamento_status.get(val_prod_nc, 0), key="rad_produtos_nao_conformes_estoque", horizontal=True, label_visibility="collapsed"
                        )
                    
                    st.write("")
                    st.markdown("**O estabelecimento realiza a coleta de medicamentos a serem descartados pela comunidade?**")
                    val_coleta_com = historico.get("coleta_medicamentos_clientes_estoque", "Não avaliado")
                    dados_inspecao_atual["coleta_medicamentos_clientes_estoque"] = st.radio(
                        "Coleta Descarte Comunidade:", opcoes_status, index=mapeamento_status.get(val_coleta_com, 0), key="rad_coleta_medicamentos_clientes_estoque", horizontal=True, label_visibility="collapsed"
                    )
            # --------------------------------------------------------------------------
            # SUB-ABA: SALA DE SERVIÇOS
            # --------------------------------------------------------------------------
            if "Sala de Serviços" in map_abas_geradas:
                with map_abas_geradas["Sala de Serviços"]:
                    st.markdown("### 🩺 Sala de Serviços Farmacêuticos")
                    
                    # --- BLOCO 1: ESTRUTURA GERAL ---
                    st.markdown("#### 🧱 Infraestrutura")
                    col_ss_est1, col_ss_est2 = st.columns(2)
                    
                    with col_ss_est1:
                        st.markdown("**Existe local específico para a prestação de serviços farmacêuticos, diverso à dispensação e devidamente isolado?**")
                        val_sep_ss = historico.get("separacao_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["separacao_sala_servicos"] = st.radio(
                            "Isolamento da Sala:", opcoes_status, index=mapeamento_status.get(val_sep_ss, 0), key="rad_separacao_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A sala de serviços possui condições adequadas de estrutura?**")
                        val_est_ss = historico.get("condicoes_estruturais_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["condicoes_estruturais_sala_servicos"] = st.radio(
                            "Estrutura da Sala:", opcoes_status, index=mapeamento_status.get(val_est_ss, 0), key="rad_condicoes_estruturais_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**O piso é liso, lavável e impermeável?**")
                        val_piso_ss = historico.get("piso_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["piso_sala_servicos"] = st.radio(
                            "Piso Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_piso_ss, 0), key="rad_piso_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**As paredes são lisas, laváveis e impermeáveis?**")
                        val_parede_ss = historico.get("parede_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["parede_sala_servicos"] = st.radio(
                            "Parede Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_parede_ss, 0), key="rad_parede_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**O teto é liso, lavável e impermeável?**")
                        val_teto_ss = historico.get("teto_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["teto_sala_servicos"] = st.radio(
                            "Teto Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_teto_ss, 0), key="rad_teto_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    with col_ss_est2:
                        st.markdown("**A sala possui condições higiênico-sanitárias adequadas?**")
                        val_hig_ss = historico.get("condicoes_higienicas_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["condicoes_higienicas_sala_servicos"] = st.radio(
                            "Higiene Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_hig_ss, 0), key="rad_condicoes_higienicas_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Há devida proteção contra entrada de pragas e insetos?**")
                        val_prag_ss = historico.get("protecao_pragas_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["protecao_pragas_sala_servicos"] = st.radio(
                            "Pragas Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_prag_ss, 0), key="rad_protecao_pragas_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A ventilação é adequada?**")
                        val_vent_ss = historico.get("ventilacao_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["ventilacao_sala_servicos"] = st.radio(
                            "Ventilação Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_vent_ss, 0), key="rad_ventilacao_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A iluminação é adequada?**")
                        val_ilum_ss = historico.get("iluminacao_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["iluminacao_sala_servicos"] = st.radio(
                            "Iluminação Sala Serviços:", opcoes_status, index=mapeamento_status.get(val_ilum_ss, 0), key="rad_iluminacao_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    st.divider()

                    # --- BLOCO 2: HIGIENE DAS MÃOS ---
                    st.markdown("#### 🧼 Higiene das Mãos")
                    col_ss_hig1, col_ss_hig2 = st.columns(2)
                    
                    with col_ss_hig1:
                        st.markdown("**É provido de lavatório com água corrente?**")
                        val_pia_ss = historico.get("pia_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["pia_sala_servicos"] = st.radio(
                            "Lavatório com Água:", opcoes_status, index=mapeamento_status.get(val_pia_ss, 0), key="rad_pia_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui sabonete líquido?**")
                        val_sab_ss = historico.get("sabonete_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["sabonete_sala_servicos"] = st.radio(
                            "Sabonete Líquido:", opcoes_status, index=mapeamento_status.get(val_sab_ss, 0), key="rad_sabonete_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui toalha de uso individual e descartável?**")
                        val_papel_ss = historico.get("papel_toalha_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["papel_toalha_sala_servicos"] = st.radio(
                            "Toalha Descartável:", opcoes_status, index=mapeamento_status.get(val_papel_ss, 0), key="rad_papel_toalha_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    with col_ss_hig2:
                        st.markdown("**Possui lixeira comum com tampa e pedal?**")
                        val_lix_com = historico.get("lixeira_comum_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["lixeira_comum_sala_servicos"] = st.radio(
                            "Lixeira Comum:", opcoes_status, index=mapeamento_status.get(val_lix_com, 0), key="rad_lixeira_comum_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui álcool para a desinfecção?**")
                        val_alc_ss = historico.get("alcool_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["alcool_sala_servicos"] = st.radio(
                            "Álcool Desinfecção:", opcoes_status, index=mapeamento_status.get(val_alc_ss, 0), key="rad_alcool_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    st.divider()

                    # --- BLOCO 3: RESÍDUOS GERADOS ---
                    st.markdown("#### ☣️ Resíduos de Saúde")
                    col_ss_res1, col_ss_res2 = st.columns(2)
                    
                    with col_ss_res1:
                        st.markdown("**Possui recipiente adequado para o descarte de perfurocortantes?**")
                        val_perf_ss = historico.get("caixa_perfurocortante_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["caixa_perfurocortante_sala_servicos"] = st.radio(
                            "Descarte Perfurocortantes:", opcoes_status, index=mapeamento_status.get(val_perf_ss, 0), key="rad_caixa_perfurocortante_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    with col_ss_res2:
                        st.markdown("**Possui coletor para resíduos infectantes, devidamente identificado?**")
                        val_inf_ss = historico.get("lixeira_infectantes_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["lixeira_infectantes_sala_servicos"] = st.radio(
                            "Coletor Infectantes:", opcoes_status, index=mapeamento_status.get(val_inf_ss, 0), key="rad_lixeira_infectantes_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    st.divider()

                    # --- BLOCO 4: MATERIAIS E DOCUMENTOS ---
                    st.markdown("#### 📄 Suporte e Segurança")
                    col_ss_doc1, col_ss_doc2 = st.columns(2)
                    
                    with col_ss_doc1:
                        st.markdown("**Possui conjunto de materiais de primeiros socorros identificado e de fácil acesso?**")
                        val_soc_ss = historico.get("primeiros_socorros_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["primeiros_socorros_sala_servicos"] = st.radio(
                            "Primeiros Socorros:", opcoes_status, index=mapeamento_status.get(val_soc_ss, 0), key="rad_primeiros_socorros_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui Lista de Estabelecimentos de Saúde próximos?**")
                        val_lista_h = historico.get("lista_estab_saude_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["lista_estab_saude_sala_servicos"] = st.radio(
                            "Lista Hospitais Próximos:", opcoes_status, index=mapeamento_status.get(val_lista_h, 0), key="rad_lista_estab_saude_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )

                    with col_ss_doc2:
                        st.markdown("**Possui Declaração de Serviços prestados?**")
                        val_decl_ss = historico.get("declaracao_serviços_sala_servicos", "Não avaliado")
                        dados_inspecao_atual["declaracao_serviços_sala_servicos"] = st.radio(
                            "Declaração de Serviços:", opcoes_status, index=mapeamento_status.get(val_decl_ss, 0), key="rad_declaracao_serviços_sala_servicos", horizontal=True, label_visibility="collapsed"
                        )
            # --------------------------------------------------------------------------
            # SUB-ABA: DML
            # --------------------------------------------------------------------------
            if "DML" in map_abas_geradas:
                with map_abas_geradas["DML"]:
                    st.markdown("### 🧹 Depósito de Material de Limpeza (DML)")
                    
                    st.markdown("**O DML possui condições adequadas de estrutura?**")
                    val_est_dml = historico.get("condicoes_estruturais_dml", "Não avaliado")
                    dados_inspecao_atual["condicoes_estruturais_dml"] = st.radio(
                        "Estrutura DML:", opcoes_status, index=mapeamento_status.get(val_est_dml, 0), key="rad_condicoes_estruturais_dml", horizontal=True, label_visibility="collapsed"
                    )
                    st.write("")
                    
                    st.markdown("**O DML possui condições higiênico-sanitárias adequadas?**")
                    val_hig_dml = historico.get("condicoes_higienicas_dml", "Não avaliado")
                    dados_inspecao_atual["condicoes_higienicas_dml"] = st.radio(
                        "Higiene DML:", opcoes_status, index=mapeamento_status.get(val_hig_dml, 0), key="rad_condicoes_higienicas_dml", horizontal=True, label_visibility="collapsed"
                    )
                    st.write("")
                    
                    st.markdown("**Os materiais e produtos de limpeza utilizados são regularizados na ANVISA?**")
                    val_reg_dml = historico.get("produto_regularizados_dml", "Não avaliado")
                    dados_inspecao_atual["produto_regularizados_dml"] = st.radio(
                        "Produtos Regularizados DML:", opcoes_status, index=mapeamento_status.get(val_reg_dml, 0), key="rad_produto_regularizados_dml", horizontal=True, label_visibility="collapsed"
                    )

            # --------------------------------------------------------------------------
            # SUB-ABA: SANITÁRIO FUNCIONÁRIOS
            # --------------------------------------------------------------------------
            if "Sanitário Funcionários" in map_abas_geradas:
                with map_abas_geradas["Sanitário Funcionários"]:
                    st.markdown("### 🚻 Sanitários dos Funcionários")
                    
                    col_sf1, col_sf2 = st.columns(2)
                    
                    with col_sf1:
                        st.markdown("**Possui condições adequadas de estrutura?**")
                        val_est_sf = historico.get("condicoes_estruturais_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["condicoes_estruturais_sanitarios_funcionarios"] = st.radio(
                            "Estrutura Sanitário Func.:", opcoes_status, index=mapeamento_status.get(val_est_sf, 0), key="rad_condicoes_estruturais_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui condições higiênico-sanitárias adequadas?**")
                        val_hig_sf = historico.get("condicoes_higienicas_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["condicoes_higienicas_sanitarios_funcionarios"] = st.radio(
                            "Higiene Sanitário Func.:", opcoes_status, index=mapeamento_status.get(val_hig_sf, 0), key="rad_condicoes_higienicas_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**É provido de lavatório com água corrente?**")
                        val_lav_sf = historico.get("lavatorio_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["lavatorio_sanitarios_funcionarios"] = st.radio(
                            "Lavatório Func.:", opcoes_status, index=mapeamento_status.get(val_lav_sf, 0), key="rad_lavatorio_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui local destinado à guarda de pertences dos funcionários?**")
                        val_pert_sf = historico.get("local_pertences_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["local_pertences_sanitarios_funcionarios"] = st.radio(
                            "Guarda de Pertences:", opcoes_status, index=mapeamento_status.get(val_pert_sf, 0), key="rad_local_pertences_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )

                    with col_sf2:
                        st.markdown("**Possui sabonete líquido para a higienização?**")
                        val_sab_sf = historico.get("sabonete_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["sabonete_sanitarios_funcionarios"] = st.radio(
                            "Sabonete Líquido Func.:", opcoes_status, index=mapeamento_status.get(val_sab_sf, 0), key="rad_sabonete_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui toalha de papel de uso individual e descartável?**")
                        val_pap_sf = historico.get("papel_toalha_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["papel_toalha_sanitarios_funcionarios"] = st.radio(
                            "Papel Toalha Func.:", opcoes_status, index=mapeamento_status.get(val_pap_sf, 0), key="rad_papel_toalha_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui lixeira com tampa e acionamento por pedal?**")
                        val_lix_sf = historico.get("lixeira_sanitarios_funcionarios", "Não avaliado")
                        dados_inspecao_atual["lixeira_sanitarios_funcionarios"] = st.radio(
                            "Lixeira Sanitário Func.:", opcoes_status, index=mapeamento_status.get(val_lix_sf, 0), key="rad_lixeira_sanitarios_funcionarios", horizontal=True, label_visibility="collapsed"
                        )

            # --------------------------------------------------------------------------
            # SUB-ABA: SANITÁRIO CLIENTES
            # --------------------------------------------------------------------------
            if "Sanitário Clientes" in map_abas_geradas:
                with map_abas_geradas["Sanitário Clientes"]:
                    st.markdown("### 🚾 Sanitários dos Clientes")
                    
                    col_sc1, col_sc2 = st.columns(2)
                    
                    with col_sc1:
                        st.markdown("**Possui condições adequadas de estrutura?**")
                        val_est_sc = historico.get("condicoes_estruturais_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["condicoes_estruturais_sanitarios_clientes"] = st.radio(
                            "Estrutura Sanitário Cli.:", opcoes_status, index=mapeamento_status.get(val_est_sc, 0), key="rad_condicoes_estruturais_sanitarios_clientes", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui condições higiênico-sanitárias adequadas?**")
                        val_hig_sc = historico.get("condicoes_higienicas_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["condicoes_higienicas_sanitarios_clientes"] = st.radio(
                            "Higiene Sanitário Cli.:", opcoes_status, index=mapeamento_status.get(val_hig_sc, 0), key="rad_condicoes_higienicas_sanitarios_clientes", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**É provido de lavatório com água corrente?**")
                        val_lav_sc = historico.get("lavatorio_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["lavatorio_sanitarios_clientes"] = st.radio(
                            "Lavatório Cli.:", opcoes_status, index=mapeamento_status.get(val_lav_sc, 0), key="rad_lavatorio_sanitarios_clientes", horizontal=True, label_visibility="collapsed"
                        )

                    with col_sc2:
                        st.markdown("**Possui sabonete líquido para a higienização?**")
                        val_sab_sc = historico.get("sabonete_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["sabonete_sanitarios_clientes"] = st.radio(
                            "Sabonete Líquido Cli.:", opcoes_status, index=mapeamento_status.get(val_sab_sc, 0), key="rad_sabonete_sanitarios_clientes", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui toalha de papel de uso individual e descartável?**")
                        val_pap_sc = historico.get("papel_toalha_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["papel_toalha_sanitarios_clientes"] = st.radio(
                            "Papel Toalha Cli.:", opcoes_status, index=mapeamento_status.get(val_pap_sc, 0), key="rad_papel_toalha_sanitarios_clientes", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**Possui lixeira com tampa e acionamento por pedal?**")
                        val_lix_sc = historico.get("lixeira_sanitarios_clientes", "Não avaliado")
                        dados_inspecao_atual["lixeira_sanitarios_clientes"] = st.radio(
                            "Lixeira Sanitário Cli.:", 
                            options=opcoes_status, 
                            index=mapeamento_status.get(val_lix_sc, 0), 
                            key="rad_lixeira_sanitarios_clientes", 
                            horizontal=True
                        )
            # --------------------------------------------------------------------------
            # SUB-ABA: COPA
            # --------------------------------------------------------------------------
            if "Copa" in map_abas_geradas:
                with map_abas_geradas["Copa"]:
                    st.markdown("### ☕ Copa / Refeitório")
                    
                    st.markdown("**Possui local separado, isolado e exclusivo para essa finalidade?**")
                    val_sep_copa = historico.get("separacao_copa", "Não avaliado")
                    dados_inspecao_atual["separacao_copa"] = st.radio(
                        "Isolamento da Copa:", opcoes_status, index=mapeamento_status.get(val_sep_copa, 0), key="rad_separacao_copa", horizontal=True, label_visibility="collapsed"
                    )
                    st.write("")
                    
                    st.markdown("**O ambiente possui condições adequadas de estrutura?**")
                    val_est_copa = historico.get("condicoes_estruturais_copa", "Não avaliado")
                    dados_inspecao_atual["condicoes_estruturais_copa"] = st.radio(
                        "Estrutura Copa:", opcoes_status, index=mapeamento_status.get(val_est_copa, 0), key="rad_condicoes_estruturais_copa", horizontal=True, label_visibility="collapsed"
                    )
                    st.write("")
                    
                    st.markdown("**O ambiente possui condições higiênico-sanitárias adequadas?**")
                    val_hig_copa = historico.get("condicoes_higienicas_copa", "Não avaliado")
                    dados_inspecao_atual["condicoes_higienicas_copa"] = st.radio(
                        "Higiene Copa:", opcoes_status, index=mapeamento_status.get(val_hig_copa, 0), key="rad_condicoes_higienicas_copa", horizontal=True, label_visibility="collapsed"
                    )

            # --------------------------------------------------------------------------
            # SUB-ABA: ADM
            # --------------------------------------------------------------------------
            if "ADM" in map_abas_geradas:
                with map_abas_geradas["ADM"]:
                    st.markdown("### 💼 Administração / Escritório")
                    
                    col_adm1, col_adm2 = st.columns(2)
                    
                    with col_adm1:
                        st.markdown("**A área administrativa possui condições adequadas de estrutura?**")
                        val_est_adm = historico.get("condicoes_estruturais_adm", "Não avaliado")
                        dados_inspecao_atual["condicoes_estruturais_adm"] = st.radio(
                            "Estrutura ADM:", opcoes_status, index=mapeamento_status.get(val_est_adm, 0), key="rad_condicoes_estruturais_adm", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A área administrativa possui condições higiênico-sanitárias adequadas?**")
                        val_hig_adm = historico.get("condicoes_higienicas_adm", "Não avaliado")
                        dados_inspecao_atual["condicoes_higienicas_adm"] = st.radio(
                            "Higiene ADM:", opcoes_status, index=mapeamento_status.get(val_hig_adm, 0), key="rad_condicoes_higienicas_adm", horizontal=True, label_visibility="collapsed"
                        )

                    with col_adm2:
                        st.markdown("**Há devida proteção contra entrada de pragas e insetos?**")
                        val_prag_adm = historico.get("protecao_pragas_adm", "Não avaliado")
                        dados_inspecao_atual["protecao_pragas_adm"] = st.radio(
                            "Pragas ADM:", opcoes_status, index=mapeamento_status.get(val_prag_adm, 0), key="rad_protecao_pragas_adm", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A ventilação no ambiente é adequada?**")
                        val_vent_adm = historico.get("ventilacao_adm", "Não avaliado")
                        dados_inspecao_atual["ventilacao_adm"] = st.radio(
                            "Ventilação ADM:", opcoes_status, index=mapeamento_status.get(val_vent_adm, 0), key="rad_ventilacao_adm", horizontal=True, label_visibility="collapsed"
                        )
                        st.write("")
                        
                        st.markdown("**A iluminação no ambiente é adequada?**")
                        val_ilum_adm = historico.get("iluminacao_adm", "Não avaliado")
                        dados_inspecao_atual["iluminacao_adm"] = st.radio(
                            "Iluminação ADM:", opcoes_status, index=mapeamento_status.get(val_ilum_adm, 0), key="rad_iluminacao_adm", horizontal=True, label_visibility="collapsed"
                        )


        # ==============================================================================
    # TAB 6: SALA DE SERVIÇOS (AÇÕES FARMACÊUTICAS AVANÇADAS)
    # ==============================================================================
    with tab_sala:
        st.markdown("## 🩺 Seção: Atividades na Sala de Serviços Farmacêuticos")
        
        # LÓGICA DE CONDICIONAL MESTRE: Só exibe se a flag foi marcada como 'Sim' na Tab 5
        if dados_inspecao_atual.get("possui_sala_servicos") != "Sim":
            st.info("💡 Esta aba só é ativada se o setor 'Sala de Serviços' for marcado como presente na aba 'Áreas Físicas'.")
        else:
            # --- 6.1 CADASTRO E CONTROLE GERAL DA SALA ---
            st.markdown("### 📋 Informações Gerais das Atividades")
            
            # Recupera a lista de serviços salvos anteriormente (armazenado como texto separado por vírgula no banco)
            val_serv_antigos = str(historico.get("servicos_realizados_sala_servicos", ""))
            lista_serv_padrao = [s.strip() for s in val_serv_antigos.split(",") if s.strip()]
            
            opcoes_servicos_sala = [
                "Aplicação de injetáveis", 
                "Aferição de pressão arterial", 
                "Verificação de temperatura corporal", 
                "Perfuração de lóbulo auricular", 
                "Dispensação via remota", 
                "Atenção ou Orientação Farmacêutica", 
                "Atendimento farmacêutico domiciliar", 
                "Vacinação", 
                "Exames de Análises Clínicas (EAC) I"
            ]
            
            st.markdown("**Quais os serviços realizados no estabelecimento?**")
            servicos_selecionados = st.multiselect(
                "Serviços ativos:",
                options=opcoes_servicos_sala,
                default=[s for s in lista_serv_padrao if s in opcoes_servicos_sala],
                key="ms_servicos_realizados_sala_servicos"
            )
            # Une a lista em uma string corrida separada por vírgula para persistência segura no banco
            dados_inspecao_atual["servicos_realizados_sala_servicos"] = ", ".join(servicos_selecionados)

            st.write("")
            col_ss_g1, col_ss_g2, col_ss_g3 = st.columns(3)
            
            with col_ss_g1:
                st.markdown("**Há procedimentos escritos para todas as atividades realizadas (POP)?**")
                v_pop_ss = historico.get("procedimentos_escritos_sala_servicos", "Não avaliado")
                dados_inspecao_atual["procedimentos_escritos_sala_servicos"] = st.radio(
                    label="POP Serviços", options=opcoes_status, index=mapeamento_status.get(v_pop_ss, 0), key="rad_procedimentos_escritos_sala_servicos", horizontal=True, label_visibility="collapsed"
                )
                
            with col_ss_g2:
                st.markdown("**Há entrega de Declaração de Serviços prestados ao usuário após cada prestação?**")
                v_dec_ss = historico.get("declaracao_sala_servicos", "Não avaliado") # mapeado para o campo correspondente
                dados_inspecao_atual["declaracao_sala_servicos"] = st.radio(
                    label="Declaração Serviços", options=opcoes_status, index=mapeamento_status.get(v_dec_ss, 0), key="rad_declaracao_sala_servicos_bloco_mestre", horizontal=True, label_visibility="collapsed"
                )
                
            with col_ss_g3:
                st.markdown("**Há registro de higienização da sala de serviços?**")
                v_limp_ss = historico.get("registro_limpeza_sala_servicos", "Não avaliado")
                dados_inspecao_atual["registro_limpeza_sala_servicos"] = st.radio(
                    label="Limpeza Sala", options=opcoes_status, index=mapeamento_status.get(v_limp_ss, 0), key="rad_registro_limpeza_sala_servicos", horizontal=True, label_visibility="collapsed"
                )

            st.divider()

            # ==============================================================================
            # DETALHAMENTO DINÂMICO DOS SERVIÇOS SELECIONADOS
            # ==============================================================================
            
            # --- 1) AFERIÇÃO DE PRESSÃO ARTERIAL ---
            if "Aferição de pressão arterial" in servicos_selecionados:
                dados_inspecao_atual["pressao_arterial"] = "sim"
                st.markdown("#### 🩺 Especificações: Aferição de Pressão Arterial")
                
                col_pa1, col_pa2 = st.columns([2, 1.5])
                with col_pa1:
                    v_marca_pa = historico.get("marca_modelo_pressao_arterial", "")
                    dados_inspecao_atual["marca_modelo_pressao_arterial"] = st.text_input("Marca e modelo do aparelho para aferição:", value=str(v_marca_pa) if v_marca_pa and str(v_marca_pa) != "None" else "", key="ti_marca_modelo_pressao_arterial")
                with col_pa2:
                    st.markdown("**Avaliação da Calibração:**")
                    v_cal_pa = historico.get("possui_calibracao_pressao_arterial", "Não avaliado")
                    dados_inspecao_atual["possui_calibracao_pressao_arterial"] = st.radio(
                        label="Calibração PA", options=opcoes_status, index=mapeamento_status.get(v_cal_pa, 0), key="rad_possui_calibracao_pressao_arterial", horizontal=True, label_visibility="collapsed"
                    )
                
                # Inicializa variáveis para o banco evitar erros de campos nulos
                dados_cal_pa = {"cert": "", "emp": "", "dt": None, "val": None}
                dados_aq_pa = {"nf": "", "dt": None, "forn": ""}

                if dados_inspecao_atual["possui_calibracao_pressao_arterial"] in ["C", "NC"]:
                    v_tipo_pa = historico.get("tipo_comprovacao_pressao_arterial", "Calibração")
                    idx_tipo_pa = 0 if v_tipo_pa == "Calibração" else 1
                    tipo_comp_pa = st.radio("Tipo de comprovação:", ["Calibração", "Aquisição recente"], index=idx_tipo_pa, key="rad_tipo_comp_pa", horizontal=True)
                    dados_inspecao_atual["tipo_comprovacao_pressao_arterial"] = tipo_comp_pa

                    if tipo_comp_pa == "Calibração":
                        col_pac1, col_pac2, col_pac3, col_pac4 = st.columns(4)
                        with col_pac1:
                            v_c1 = historico.get("certificado_calibracao_pressao_arterial", "")
                            dados_cal_pa["cert"] = st.text_input("Nº certificado:", value=str(v_c1) if v_c1 and str(v_c1) != "None" else "", key="ti_cert_pa")
                        with col_pac2:
                            v_c2 = historico.get("empresa_calibracao_pressao_arterial", "")
                            dados_cal_pa["emp"] = st.text_input("Empresa prestadora:", value=str(v_c2) if v_c2 and str(v_c2) != "None" else "", key="ti_emp_pa")
                        with col_pac3:
                            v_c3 = historico.get("data_calibracao_pressao_arterial", None)
                            dt_c3 = pd.to_datetime(v_c3).date() if v_c3 and not pd.isna(v_c3) else None
                            dados_cal_pa["dt"] = st.date_input("Data Calibração:", value=dt_c3, key="di_dt_pa")
                        with col_pac4:
                            v_c4 = historico.get("validade_calibracao_pressao_arterial", None)
                            dt_c4 = pd.to_datetime(v_c4).date() if v_c4 and not pd.isna(v_c4) else None
                            dados_cal_pa["val"] = st.date_input("Validade:", value=dt_c4, key="di_val_pa")
                    else:
                        col_paa1, col_paa2, col_paa3 = st.columns(3)
                        with col_paa1:
                            v_a1 = historico.get("nota_aquisicao_pressao_arterial", "")
                            dados_aq_pa["nf"] = st.text_input("Nº Nota Fiscal:", value=str(v_a1) if v_a1 and str(v_a1) != "None" else "", key="ti_nf_pa")
                        with col_paa2:
                            v_a2 = historico.get("data_aquisicao_pressao_arterial", None)
                            dt_a2 = pd.to_datetime(v_a2).date() if v_a2 and not pd.isna(v_a2) else None
                            dados_aq_pa["dt"] = st.date_input("Data Aquisição:", value=dt_a2, key="di_dt_aq_pa")
                        with col_paa3:
                            v_a3 = historico.get("empresa_aquisicao_pressao_arterial", "")
                            dados_aq_pa["forn"] = st.text_input("Fornecedor:", value=str(v_a3) if v_a3 and str(v_a3) != "None" else "", key="ti_forn_pa")
                
                # Atribui o empacotamento das variáveis para envio estável
                dados_inspecao_atual["certificado_calibracao_pressao_arterial"] = dados_cal_pa["cert"]
                dados_inspecao_atual["empresa_calibracao_pressao_arterial"] = dados_cal_pa["emp"]
                dados_inspecao_atual["data_calibracao_pressao_arterial"] = dados_cal_pa["dt"]
                dados_inspecao_atual["validade_calibracao_pressao_arterial"] = dados_cal_pa["val"]
                dados_inspecao_atual["nota_aquisicao_pressao_arterial"] = dados_aq_pa["nf"]
                dados_inspecao_atual["data_aquisicao_pressao_arterial"] = dados_aq_pa["dt"]
                dados_inspecao_atual["empresa_aquisicao_pressao_arterial"] = dados_aq_pa["forn"]
                st.divider()
            else:
                dados_inspecao_atual["pressao_arterial"] = "não"

            # --- 2) VERIFICAÇÃO DE TEMPERATURA CORPORAL ---
            if "Verificação de temperatura corporal" in servicos_selecionados:
                dados_inspecao_atual["t_corporal"] = "sim"
                st.markdown("#### 🌡️ Especificações: Verificação de Temperatura Corporal")
                v_marca_tc = historico.get("marca_modelo_t_corporal", "")
                dados_inspecao_atual["marca_modelo_t_corporal"] = st.text_input(
                dados_inspecao_atual["marca_modelo_t_corporal"] = st.text_input(
                    "Marca e modelo do termômetro clínico utilizado:", value=str(v_marca_tc) if v_marca_tc and str(v_marca_tc) != "None" else "", key="ti_marca_modelo_t_corporal"
                )
                st.divider()
            else:
                dados_inspecao_atual["t_corporal"] = "não"

            # --- 3) PERFURAÇÃO DE LÓBULO AURICULAR ---
            if "Perfuração de lóbulo auricular" in servicos_selecionados:
                dados_inspecao_atual["perfuracao_orelha"] = "sim"
                st.markdown("#### 👂 Especificações: Perfuração de Lóbulo Auricular")
                col_po1, col_po2 = st.columns(2)
                with col_po1:
                    v_marca_po = historico.get("marca_modelo_perfuracao_orelha", "")
                    dados_inspecao_atual["marca_modelo_perfuracao_orelha"] = st.text_input("Marca e modelo da pistola de perfuração:", value=str(v_marca_po) if v_marca_po and str(v_marca_po) != "None" else "", key="ti_marca_modelo_perfuracao_orelha")
                with col_po2:
                    v_lote_po = historico.get("lote_perfuracao_orelha", "")
                    dados_inspecao_atual["lote_perfuracao_orelha"] = st.text_input("Lote do equipamento/brincos:", value=str(v_lote_po) if v_lote_po and str(v_lote_po) != "None" else "", key="ti_lote_perfuracao_orelha")
                st.divider()
            else:
                dados_inspecao_atual["perfuracao_orelha"] = "não"

            # --- 4) VACINAÇÃO ---
            if "Vacinação" in servicos_selecionados:
                dados_inspecao_atual["servico_vacinacao"] = "sim"
                st.markdown("#### 💉 Especificações: Serviço de Vacinação")
                col_v1, col_v2 = st.columns(2)
                
                with col_v1:
                    st.markdown("**Possui programa de capacitação sobre vacinação para os profissionais envolvidos?**")
                    v_cap_v = historico.get("capacitacao_vacinacao", "Não avaliado")
                    dados_inspecao_atual["capacitacao_vacinacao"] = st.radio(label="Cap Vacina", options=opcoes_status, index=mapeamento_status.get(v_cap_v, 0), key="rad_capacitacao_vacinacao", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Possui equipamento de refrigeração exclusivo para as vacinas?**")
                    v_ref_v = historico.get("refrigerador_vacinacao", "Não avaliado")
                    dados_inspecao_atual["refrigerador_vacinacao"] = st.radio(label="Ref Vacina", options=opcoes_status, index=mapeamento_status.get(v_ref_v, 0), key="rad_refrigerador_vacinacao", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Mantém prontuário individual com registro das vacinas aplicadas?**")
                    v_reg_v = historico.get("registro_vacinacao", "Não avaliado")
                    dados_inspecao_atual["registro_vacinacao"] = st.radio(label="Reg Vacina", options=opcoes_status, index=mapeamento_status.get(v_reg_v, 0), key="rad_registro_vacinacao", horizontal=True, label_visibility="collapsed")
                
                with col_v2:
                    st.markdown("**Possui comprovação de origem das vacinas?**")
                    v_orig_v = historico.get("origem_vacinacao", "Não avaliado")
                    dados_inspecao_atual["origem_vacinacao"] = st.radio(label="Orig Vacina", options=opcoes_status, index=mapeamento_status.get(v_orig_v, 0), key="rad_origem_vacinacao", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Notifica ocorrência de efeitos adversos e erros nos sistemas pertinentes?**")
                    v_ocor_v = historico.get("ocorrencias_vacinacao", "Não avaliado")
                    dados_inspecao_atual["ocorrencias_vacinacao"] = st.radio(label="Ocor Vacina", options=opcoes_status, index=mapeamento_status.get(v_ocor_v, 0), key="rad_ocorrencias_vacinacao", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Realiza emissão de CIVP, devidamente credenciado pela ANVISA?**")
                    v_civp_v = historico.get("civp_vacinacao", "Não avaliado")
                    dados_inspecao_atual["civp_vacinacao"] = st.radio(label="CIVP Vacina", options=opcoes_status, index=mapeamento_status.get(v_civp_v, 0), key="rad_civp_vacinacao", horizontal=True, label_visibility="collapsed")
                
                st.write("")
                st.markdown("**Possui lista de vacinas integrantes do Sistema Único de Saúde (SUS) exposta de forma visível?**")
                v_list_v = historico.get("lista_vacinacao", "Não avaliado")
                dados_inspecao_atual["lista_vacinacao"] = st.radio(label="Lista SUS", options=opcoes_status, index=mapeamento_status.get(v_list_v, 0), key="rad_lista_vacinacao", horizontal=True, label_visibility="collapsed")
                st.divider()
            else:
                dados_inspecao_atual["servico_vacinacao"] = "não"
            # --- 5) EXAMES DE ANÁLISES CLÍNICAS (EAC) I ---
            if "Exames de Análises Clínicas (EAC) I" in servicos_selecionados:
                dados_inspecao_atual["eac_i"] = "sim"
                st.markdown("#### 🔬 Especificações: Exames de Análises Clínicas (EAC) Tipo I")
                col_e1, col_e2 = st.columns(2)
                
                with col_e1:
                    st.markdown("**Realiza apenas coleta para testes Tipo I (punção capilar, saliva, swab nasal/orofaringe)?**")
                    v_t1 = historico.get("tipo_exames_eac_i", "Não avaliado")
                    dados_inspecao_atual["tipo_exames_eac_i"] = st.radio(label="Tipo EAC", options=opcoes_status, index=mapeamento_status.get(v_t1, 0), key="rad_tipo_exames_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Possui equipamento de refrigeração para diagnóstico que necessite dessas condições?**")
                    v_ref_e = historico.get("refrigerador_eac_i", "Não avaliado")
                    dados_inspecao_atual["refrigerador_eac_i"] = st.radio(label="Ref EAC", options=opcoes_status, index=mapeamento_status.get(v_ref_e, 0), key="rad_refrigerador_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Possui depósito adequado para equipamentos e materiais próprios?**")
                    v_dep_e = historico.get("deposito_eac_i", "Não avaliado")
                    dados_inspecao_atual["deposito_eac_i"] = st.radio(label="Dep EAC", options=opcoes_status, index=mapeamento_status.get(v_dep_e, 0), key="rad_deposito_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Possui adequado Programa de Garantia da Qualidade (PGQ) para todos os exames realizados?**")
                    v_pgq_e = historico.get("pgq_eac_i", "Não avaliado")
                    dados_inspecao_atual["pgq_eac_i"] = st.radio(label="PGQ EAC", options=opcoes_status, index=mapeamento_status.get(v_pgq_e, 0), key="rad_pgq_eac_i", horizontal=True, label_visibility="collapsed")
                
                with col_e2:
                    st.markdown("**Todos os produtos correlatos utilizados no serviço possuem regularização na ANVISA?**")
                    v_cor_e = historico.get("correlatos_eac_i", "Não avaliado")
                    dados_inspecao_atual["correlatos_eac_i"] = st.radio(label="Corr EAC", options=opcoes_status, index=mapeamento_status.get(v_cor_e, 0), key="rad_correlatos_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Os produtos e equipamentos são utilizados apenas para os fins a que se destinam?**")
                    v_ut_e = historico.get("utilizacao_correcta_eac_i", "Não avaliado") # mantendo o alias mapeado na tabela original
                    dados_inspecao_atual["utilizacao_correcta_eac_i"] = st.radio(label="Utiliz EAC", options=opcoes_status, index=mapeamento_status.get(v_ut_e, 0), key="rad_utilizacao_correcta_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Realiza devida manutenção preventiva ou calibração dos equipamentos utilizados?**")
                    v_man_e = historico.get("manutencao_equipamentos_eac_i", "Não avaliado")
                    dados_inspecao_atual["manutencao_equipamentos_eac_i"] = st.radio(label="Manut EAC", options=opcoes_status, index=mapeamento_status.get(v_man_e, 0), key="rad_manutencao_equipamentos_eac_i", horizontal=True, label_visibility="collapsed")
                    st.write("")
                    
                    st.markdown("**Possui programa de educação permanente com capacitações referentes aos EAC?**")
                    v_cap_e = historico.get("capacitacao_eac_i", "Não avaliado")
                    dados_inspecao_atual["capacitacao_eac_i"] = st.radio(label="Cap EAC", options=opcoes_status, index=mapeamento_status.get(v_cap_e, 0), key="rad_capacitacao_eac_i", horizontal=True, label_visibility="collapsed")
                
                st.divider()
            else:
                dados_inspecao_atual["eac_i"] = "não"

        # ==============================================================================
    # TAB 7: RECURSOS HUMANOS / FUNCIONÁRIOS
    # ==============================================================================
    with tab_funcionarios:
        st.markdown("## 👷 Seção: Recursos Humanos e Saúde do Trabalhador")
        st.markdown("### 📋 Controle de Pessoal, Proteção e Capacitações")
        
        # Criação do layout em duas colunas simétricas
        col_func1, col_func2 = st.columns(2)
        
        with col_func1:
            # 1. Número de funcionários (Tratamento seguro contra nulos ou NaNs do banco)
            val_num_func = historico.get("numero_funcionarios", 0)
            try:
                num_func_padrao = int(val_num_func) if val_num_func is not None and not pd.isna(val_num_func) else 0
            except ValueError:
                num_func_padrao = 0
                
            dados_inspecao_atual["numero_funcionarios"] = st.number_input(
                "Número total de funcionários / colaboradores ativos:",
                min_value=0,
                value=num_func_padrao,
                step=1,
                key="ni_numero_funcionarios"
            )
            st.write("")
            
            # 2. Uniformes adequados
            st.markdown("**Os funcionários possuem e utilizam uniformes limpos, adequados e em bom estado de conservação?**")
            val_uni_f = historico.get("uniformes_funcionarios", "Não avaliado")
            dados_inspecao_atual["uniformes_funcionarios"] = st.radio(
                label="Uniformes Funcionários",
                options=opcoes_status,
                index=mapeamento_status.get(val_uni_f, 0),
                key="rad_uniformes_funcionarios",
                horizontal=True,
                label_visibility="collapsed"
            )

        with col_func2:
            # 3. Equipamentos de Proteção Individual (EPIs)
            st.markdown("**Há Equipamentos de Proteção Individual (EPIs) adequados e em quantidade suficiente para os serviços realizados pelos colaboradores?**")
            val_epi_f = historico.get("epis_serviços_funcionarios", "Não avaliado")
            dados_inspecao_atual["epis_serviços_funcionarios"] = st.radio(
                label="EPIs Funcionários",
                options=opcoes_status,
                index=mapeamento_status.get(val_epi_f, 0),
                key="rad_epis_servicos_funcionarios",
                horizontal=True,
                label_visibility="collapsed"
            )
            st.write("")
            
            # 4. Treinamentos e Capacitação
            st.markdown("**Há a devida capacitação/educação continuada de todos os colaboradores para as suas funções, com registros mantidos?**")
            val_cap_f = historico.get("capacitacao_funcionarios", "Não avaliado")
            dados_inspecao_atual["capacitacao_funcionarios"] = st.radio(
                label="Capacitação Funcionários",
                options=opcoes_status,
                index=mapeamento_status.get(val_cap_f, 0),
                key="rad_capacitacao_funcionarios",
                horizontal=True,
                label_visibility="collapsed"
            )

    # ==============================================================================
    # 5. FINALIZAR RELATÓRIO E PERSISTIR NO SUPABASE (VERSÃO ATUALIZADA COM TABS 3 E 4)
    # ==============================================================================
    st.markdown("### 💾 Encerramento da Inspeção")
    
    if st.button("⚙️ Gravar Dados da Inspeção na Nuvem", type="primary", use_container_width=True, key="btn_salvar_supabase_mestre"):
        if conn_nuvem is None:
            st.error("❌ Não foi possível salvar: O aplicativo está operando em Modo Offline (Sem conexão com o banco).")
        elif not cnpj_estabelecimento:
            st.error("❌ Erro grave: O CNPJ do estabelecimento não foi identificado.")
        else:
            with st.spinner("Sincronizando dados e salvando histórico no Supabase..."):
                try:
                    import json

                    # 🚨 1. CAPTURA FORÇADA DE SEGURANÇA: CONTATOS (TAB 1) 🚨
                    lista_atualizada_salvamento = []
                    if "lista_pessoas_formulario" in st.session_state:
                        for idx in range(len(st.session_state["lista_pessoas_formulario"])):
                            nome_p = st.session_state.get(f"p_nome_{idx}", "").strip()
                            cpf_p = st.session_state.get(f"p_cpf_{idx}", "").strip()
                            cargo_p = st.session_state.get(f"p_cargo_{idx}", "").strip()
                            email_p = st.session_state.get(f"p_email_{idx}", "").strip()
                            if nome_p:
                                lista_atualizada_salvamento.append(f"{nome_p}/{cpf_p}/{cargo_p}/{email_p}")
                    
                    if lista_atualizada_salvamento:
                        dados_inspecao_atual["acompanhantes_inspecao"] = " | ".join(lista_atualizada_salvamento)

                    # 🚨 2. CAPTURA FORÇADA E CONVERSÃO DE SEGURANÇA: TERMÔMETROS (TAB 3) 🚨
                    if "locais_controle_geladeiras_json" not in dados_inspecao_atual:
                        dados_inspecao_atual["locais_controle_geladeiras_json"] = {
                            "lista_tu": lista_tu_atualizada if 'lista_tu_atualizada' in locals() else st.session_state.get("lista_tu_formulario", []),
                            "lista_t": lista_t_atualizada if 'lista_t_atualizada' in locals() else st.session_state.get("lista_t_formulario", [])
                        }
                    
                    if isinstance(dados_inspecao_atual["locais_controle_geladeiras_json"], dict):
                        dados_inspecao_atual["locais_controle_geladeiras_json"] = json.dumps(
                            dados_inspecao_atual["locais_controle_geladeiras_json"]
                        )

                    # 🚨 3. CAPTURA FORÇADA E CONVERSÃO DE SEGURANÇA: ESTOQUE CONTROLADOS E ANTIMICROBIANOS (TAB 4) 🚨
                    if "medicamentos_controlados_verificados_json" not in dados_inspecao_atual:
                        dados_inspecao_atual["medicamentos_controlados_verificados_json"] = lista_mc_atualizada if 'lista_mc_atualizada' in locals() else st.session_state.get("lista_mc_formulario", [])
                    
                    if "antimicrobianos_verificados_json" not in dados_inspecao_atual:
                        dados_inspecao_atual["antimicroresponsaveis_verificados_json"] = lista_am_atualizada if 'lista_am_atualizada' in locals() else st.session_state.get("lista_am_formulario", [])

                    # Serializa as tabelas da Tab 4 em strings JSON válidas antes do envio ao Supabase
                    for col_jsonb in ["medicamentos_controlados_verificados_json", "antimicrobianos_verificados_json"]:
                        if col_jsonb in dados_inspecao_atual:
                            if isinstance(dados_inspecao_atual[col_jsonb], list):
                                dados_inspecao_atual[col_jsonb] = json.dumps(dados_inspecao_atual[col_jsonb])

                    # 🚨 4. GERADOR DINÂMICO DE SQL E ENVIO 🚨
                    colunas_presentes = list(dados_inspecao_atual.keys())
                    placeholders = [f":{col}" for col in colunas_presentes]
                    
                    query_insert_dinamica = f"""
                        INSERT INTO inspecoes_db ({", ".join(colunas_presentes)})
                        VALUES ({", ".join(placeholders)});
                    """
                    
                    with conn_nuvem.session as session:
                        session.execute(text(query_insert_dinamica), dados_inspecao_atual)
                        session.commit()
                    
                    st.success("🟢 Sucesso! Os dados foram gravados de forma permanente na tabela 'inspecoes_db'.")
                    st.toast("Inspeção salva com sucesso!", icon="💾")
                    
                    st.session_state["historico_pre_preenchido"] = dados_inspecao_atual.copy()
                    
                except Exception as e_gravacao:
                    st.error(f"❌ Falha técnica ao tentar gravar no Supabase: {e_gravacao}")










   
        