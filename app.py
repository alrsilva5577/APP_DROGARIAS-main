import streamlit as st
from datetime import datetime
from docx import Document
import pandas as pd
import requests
import io

# 1. Configurações iniciais da página
st.set_page_config(page_title="Módulo de Inspeção", layout="wide")
st.title("💊 Sistema de Gestão de Inspeções - DROGARIAS")

# --- LEITURA DOS ARQUIVOS EXCEL LOCAIS (LEIS E FISCAIS) ---
@st.cache_data 
def carregar_dados_locais():
    try:
        df_inf = pd.read_excel("INFRACOES_DB.xlsx")
    except FileNotFoundError:
        st.error("❌ Erro: O arquivo 'INFRACOES_DB.xlsx' não foi encontrado.")
        df_inf = None
        
    try:
        df_fisc = pd.read_excel("FISCAIS_DB.xlsx")
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

# --- CRIAÇÃO AUTOMÁTICA DA TABELA NO SUPABASE SE NÃO EXISTIR (ATUALIZADA) ---
if conn_nuvem is not None:
    try:
        from sqlalchemy import text
        with conn_nuvem.session as s:
            s.execute(text("""
            CREATE TABLE IF NOT EXISTS inspecoes_db (
                id_inspecao TEXT PRIMARY KEY,
                id_renovacao TEXT,
                cnpj_estabelecimento TEXT,
                tipo_acao TEXT,
                data_procedimento TEXT,
                status_inspecao_itens TEXT,
                numero_colaboradores INTEGER,
                acompanhantes_inspecao TEXT
            );
            """))
            s.commit()
    except Exception as e:
        st.sidebar.warning(f"Aviso de tabela: {e}")


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
def carregar_historico_cnpj(cnpj_alvo):
    if conn_nuvem is not None:
        try:
            from sqlalchemy import text
            # Atualizamos o SELECT para buscar também o número de colaboradores e acompanhantes
            query = """
                SELECT status_inspecao_itens, numero_colaboradores, acompanhantes_inspecao 
                FROM inspecoes_db 
                WHERE cnpj_estabelecimento = :cnpj 
                ORDER BY id_inspecao DESC 
                LIMIT 1;
            """
            df_resultado = conn_nuvem.query(text(query), params={"cnpj": str(cnpj_alvo)}, ttl=0)
            
            if not df_resultado.empty:
                # 1. Recupera o histórico de respostas técnicos (C, NC, NA)
                string_status = str(df_resultado.iloc[0]["status_inspecao_itens"])
                dicionario_respostas = {}
                if string_status.strip() != "" and string_status != "nan":
                    pares = string_status.split(",")
                    for par in pares:
                        if ":" in par:
                            chave, valor = par.split(":")
                            dicionario_respostas[chave.strip()] = valor.strip()
                st.session_state["historico_pre_preenchido"] = dicionario_respostas
                
                # 2. Recupera o número de funcionários
                num_colab_banco = df_resultado.iloc[0]["numero_colaboradores"]
                st.session_state["num_colaboradores_banco"] = int(num_colab_banco) if pd.notna(num_colab_banco) else 0
                
                # 3. Recupera e reconstrói a lista de acompanhantes
                txt_acompanhantes = str(df_resultado.iloc[0]["acompanhantes_inspecao"])
                lista_reconstruida = []
                if txt_acompanhantes.strip() != "" and txt_acompanhantes != "nan":
                    pessoas_brutas = txt_acompanhantes.split(" | ")
                    for pess in pessoas_brutas:
                        if "/" in pess:
                            partes = pess.split("/")
                            lista_reconstruida.append({
                                "nome": partes[0] if len(partes) > 0 else "",
                                "cpf": partes[1] if len(partes) > 1 else "",
                                "cargo": partes[2] if len(partes) > 2 else "",
                                "email": partes[3] if len(partes) > 3 else ""
                            })
                
                if lista_reconstruida:
                    st.session_state["lista_pessoas"] = lista_reconstruida
                else:
                    st.session_state["lista_pessoas"] = [{"nome": "", "cpf": "", "cargo": "", "email": ""}]
                
                return True
        except Exception as e:
            print(f"Erro ao ler PostgreSQL: {e}")
            
    # Se for empresa nova, redefine os estados iniciais limpos
    st.session_state["historico_pre_preenchido"] = {}
    st.session_state["num_colaboradores_banco"] = 0
    st.session_state["lista_pessoas"] = [{"nome": "", "cpf": "", "cargo": "", "email": ""}]
    return False

# ==============================================================================
# TELA 1: FLUXO INICIAL
# ==============================================================================
if st.session_state["tela"] == "fluxo_inicial":
    st.subheader("1. Identificação da Inspeção")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Número da Solicitação WEB**")
        web = st.text_input("WEB Entrada", key="init_web", label_visibility="collapsed", placeholder="Ex: 2026-X812")
        
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
                                from sqlalchemy import text
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
        if st.button("🔍 Adicionar Inspeção Sanitária (IS)", use_container_width=True, type="primary"):
            st.session_state["tela"] = "inspecao_sanitaria"
            st.rerun()
        if st.button("📂 Adicionar Avaliação Documental (AD)", use_container_width=True, disabled=True):
            pass
            
    with col_menu2:
        st.markdown("### 📄 Relatórios e Documentos Oficiais")
        # O botão de adequações pendentes continua consultando o histórico normalmente...
        if st.button("📋 Gerar Lista de Adequações Pendentes", use_container_width=True):
            if conn_nuvem is not None:
                try:
                    df_historico_completo = conn_nuvem.read(worksheet="INSPECOES_DB", ttl="0s")
                    df_ult_inspecao = df_historico_completo[df_historico_completo["id_renovacao"].astype(str) == str(web_ativa)]
                    
                    if df_ult_inspecao.empty:
                        st.warning("⚠️ Nenhuma inspeção anterior registrada na nuvem para esta solicitação WEB.")
                    elif df_excel is None:
                        st.error("Planilha de Legislações ausente.")
                    else:
                        chaves_texto = str(df_ult_inspecao.iloc[-1]["lista_adequacoes_chaves"])
                        if chaves_texto.strip() == "" or chaves_texto == "nan":
                            st.success("🟢 Excelente! A última inspeção não registrou nenhuma inadequação pendente.")
                        else:
                            lista_chaves = chaves_texto.split(",")
                            doc_exigencias = Document()
                            doc_exigencias.add_heading("7- Orientações / Providências", level=1)
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
                            st.download_button(label="💾 Baixar Lista de Adequações (.docx)", data=buffer_exig, file_name=f"Lista_Adequacoes_{web_ativa}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"Erro ao acessar SQL: {e}")

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
                st.markdown("**CPF**")
                st.session_state["lista_pessoas"][indice]["cpf"] = st.text_input("CPF", value=p["cpf"], key=f"cpf_{indice}", label_visibility="collapsed")
            with col3:
                st.markdown("**Cargo**")
                st.session_state["lista_pessoas"][indice]["cargo"] = st.text_input("Cargo", value=p["cargo"], key=f"cargo_{indice}", label_visibility="collapsed")
            with col4:
                st.markdown("**E-mail**")
                st.session_state["lista_pessoas"][indice]["email"] = st.text_input("E-mail", value=p["email"], key=f"email_{indice}", label_visibility="collapsed")
                
        if st.button("➕ Adicionar Pessoa"):
            st.session_state["lista_pessoas"].append({"nome": "", "cpf": "", "cargo": "", "email": ""})
            st.rerun()
            
        st.divider()
        st.markdown("#### Informações")
        st.markdown("**Número de colaboradores:**")
        num_colaboradores = st.number_input("Número de colaboradores:", min_value=0, step=1, value=0, label_visibility="collapsed")
        
        st.divider()
        col_rt, col_botao = st.columns([3, 1])
        
        with col_rt:
            st.markdown("**Farmacêutico presente no ato da inspeção?**")
            
            # --- LÓGICA DE PRÉ-PREENCHIMENTO INTELIGENTE ---
            # Busca o status gravado no banco. Se não achar nada (empresa nova), o padrão é "Não avaliado"
            status_anterior_rt = st.session_state["historico_pre_preenchido"].get("RDC44_3", "Não avaliado")
            
            # Mapeia qual botão de rádio deve iniciar aceso
            mapeamento_indices = {"Não avaliado": 0, "C": 1, "NC": 2, "NA": 3}
            indice_padrao = mapeamento_indices.get(status_anterior_rt, 0)
            
            status_rt = st.radio(
                "RT Status", ["Não avaliado", "C", "NC", "NA"],
                index=indice_padrao,
                horizontal=True, 
                label_visibility="collapsed"
            )
            
        with col_botao:
            st.write("")
            if st.button("➕ Observação"):
                st.session_state["mostrar_observacao_rt"] = True
                st.rerun()
                
        obs_rt = ""
        if st.session_state["mostrar_observacao_rt"]:
            st.markdown("**Digite as observações sobre a situação do farmacêutico:**")
            obs_rt = st.text_area("Obs RT Texto", placeholder="Ex: Farmacêutico ausente...", label_visibility="collapsed")
            
        st.divider()
        st.markdown("### 💾 Finalizar Relatório")
        
        # 1. ETAPA DE PROCESSAMENTO (Grava no banco e prepara o Word na memória)
        if st.button("⚙️ Processar Relatório e Salvar na Nuvem", type="primary", key="btn_processar_geral"):
            if df_excel is None:
                st.error("Não é possível gerar o relatório sem carregar a planilha 'INFRACOES_DB.xlsx'.")
            else:
                with st.spinner("Sincronizando dados com o Google Sheets..."):
                    # Montagem do Word
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
                    lista_salvamento_banco = []
                    
                    if not linha_rt.empty:
                        if status_rt == "C":
                            frase_rt_final = str(linha_rt["frase_conforme"].values)
                            lista_salvamento_banco.append("RDC44_3:C")
                        elif status_rt == "NC":
                            frase_rt_final = str(linha_rt["frase_nao_conforme"].values)
                            lista_salvamento_banco.append("RDC44_3:NC")
                        elif status_rt == "NA":
                            lista_salvamento_banco.append("RDC44_3:NA")
                        elif status_rt == "Não avaliado":
                            lista_salvamento_banco.append("RDC44_3:Não avaliado")
                            
                    if obs_rt.strip() != "":
                        frase_rt_final += f" {obs_rt}" if frase_rt_final else obs_rt
                    if frase_rt_final:
                        doc.add_paragraph(frase_rt_final)
                        
                    doc.add_paragraph("")
                    doc.add_heading("1- Informações gerais:", level=1)
                    info_gerais = f"Estabelecimento que desenvolve atividade enquadrada no Agrupamento 28 – Comércio Varejista de Medicamentos – CNAE Fiscal 4771-7/01 – Comércio Varejista de Produtos Farmacêuticos sem Manipulação de Fórmulas.\nNúmero de colaboradores: {num_colaboradores}."
                    doc.add_paragraph(info_gerais)
                    
                    # Salva o arquivo Word na memória da sessão para não perder no reload
                    buffer = io.BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    st.session_state["buffer_word_pronto"] = buffer.getvalue()
                    
                    # --- GRAVAÇÃO AUTOMÁTICA NO BANCO SQL EM NUVEM (POSTGRESQL) ---
                                        # --- GRAVAÇÃO AUTOMÁTICA EM NUVEM (POSTGRESQL ATUALIZADA) ---
                    if conn_nuvem is not None:
                        timestamp_chave = datetime.now().strftime("%Y%m%d_%H%M%S")
                        chave_primaria_inspecao = f"{n_web}IS{timestamp_chave}"
                        texto_status_banco = ",".join(lista_salvamento_banco)
                        
                        # Transforma a lista dinâmica de pessoas em uma única linha de texto limpa para o banco
                        lista_salvamento_pessoas = []
                        for p in st.session_state["lista_pessoas"]:
                            if p["nome"].strip() != "":
                                # Guarda no formato: Nome/CPF/Cargo/Email
                                info_compacta = f"{p['nome'].strip()}/{p['cpf'].strip()}/{p['cargo'].strip()}/{p['email'].strip()}"
                                lista_salvamento_pessoas.append(info_compacta)
                        texto_pessoas_banco = " | ".join(lista_salvamento_pessoas)
                        
                        try:
                            from sqlalchemy import text
                            with conn_nuvem.session as session:
                                # Adicionamos os novos campos e parâmetros no comando INSERT do SQL
                                session.execute(
                                    text("""
                                    INSERT INTO inspecoes_db (id_inspecao, id_renovacao, cnpj_estabelecimento, tipo_acao, data_procedimento, status_inspecao_itens, numero_colaboradores, acompanhantes_inspecao)
                                    VALUES (:id_inspecao, :id_renovacao, :cnpj_estabelecimento, :tipo_acao, :data_procedimento, :status_inspecao_itens, :numero_colaboradores, :acompanhantes_inspecao);
                                    """),
                                    {
                                        "id_inspecao": chave_primaria_inspecao,
                                        "id_renovacao": n_web,
                                        "cnpj_estabelecimento": cnpj_estabelecimento,
                                        "tipo_acao": "IS",
                                        "data_procedimento": data_atual,
                                        "status_inspecao_itens": texto_status_banco,
                                        "numero_colaboradores": int(num_colaboradores),
                                        "acompanhantes_inspecao": texto_pessoas_banco
                                    }
                                )
                                session.commit()
                            st.session_state["gravou_nuvem_sucesso"] = True
                        except Exception as e:
                            st.session_state["erro_nuvem_mensagem"] = str(e)
                            st.session_state["gravou_nuvem_sucesso"] = False
                    st.rerun()


        # 2. ETAPA DE FEEDBACK E DOWNLOAD (Fora do botão anterior para evitar resets)
        if "gravou_nuvem_sucesso" in st.session_state:
            if st.session_state["gravou_nuvem_sucesso"]:
                st.success("✅ Dados salvos com sucesso na planilha do Google Sheets!")
            else:
                erro_txt = st.session_state.get("erro_nuvem_mensagem", "Erro desconhecido.")
                st.error(f"❌ O Word foi gerado, mas houve um erro ao salvar no Google Sheets: {erro_txt}")
                st.info("Verifique se as credenciais no Secrets da nuvem contêm a linha de 'scopes' corretamente.")

            # Exibe o botão de download estável se o arquivo estiver pronto na memória
            if "buffer_word_pronto" in st.session_state:
                st.download_button(
                    label="💾 Clique aqui para baixar o arquivo .docx", 
                    data=st.session_state["buffer_word_pronto"], 
                    file_name=f"Relatorio_Inspecao_{n_web}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="btn_download_estavel"
                )
