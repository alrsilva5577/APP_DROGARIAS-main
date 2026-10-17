import streamlit as st
from datetime import datetime
from docx import Document
import pandas as pd
import requests
import io
from sqlalchemy import text

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

# --- CRIAÇÃO AUTOMÁTICA DA TABELA NO SUPABASE COM TODOS OS CAMPOS ---
if conn_nuvem is not None:
    try:
        with conn_nuvem.session as s:
            s.execute(text("""
            CREATE TABLE IF NOT EXISTS inspecoes_db (
                -- Identificação e Dados Cadastrais
                id_inspecao TEXT PRIMARY KEY,
                cnpj TEXT,
                web TEXT,
                data_procedimento DATE,
                acompanhantes_inspecao TEXT,
                farmaceutico_presente TEXT,

                -- (documentacao) - AFE
                possui_afe TEXT,
                numero_afe TEXT,
                data_afe DATE,	
                status_afe TEXT,

                -- (documentacao) - CRT
                possui_crt TEXT,
                numero_crt TEXT,
                data_crt DATE,
                status_crt TEXT,
                rt_crt TEXT,

                -- Manuais
                possui_mbpf TEXT,
                possui_pops TEXT,

                -- Controle de Pragas
                possui_pragas TEXT,
                empresa_pragas TEXT,
                cnpj_pragas TEXT,
                data_pragas DATE,
                validade_pragas DATE,
                val_alvara_pragas DATE,

                -- Reservatório de Água
                possui_limpeza_reserv_agua TEXT,
                empresa_reserv_agua TEXT,
                cnpj_reserv_agua TEXT,
                data_reserv_agua DATE,

                -- Vistoria do Corpo de Bombeiros
                possui_vistoria_bombeiro TEXT,
                tipo_vistoria_bombeiro TEXT,
                num_vistoria_bombeiro TEXT,
                data_vistoria_bombeiro DATE,
                validade_vistoria_bombeiro DATE,

                -- PGRSS
                possui_pgrss TEXT,
                data_pgrss DATE,
                comp_coleta_pgrss TEXT,
                data_coleta_residuo DATE,

                -- PCMSO
                possui_pcmso TEXT,
                emissao_pcmso DATE,
                validade_pcmso DATE,
                empresa_pcmso TEXT,
                medico_pcmso TEXT,

                -- PGR
                possui_pgr TEXT,
                emissao_pgr DATE,
                validade_pgr DATE,
                responsavel_pgr TEXT,

                -- ASO
                possui_aso TEXT,
                empresa_aso TEXT,
                numero_aso TEXT,

                -- Lista Dinâmica: Locais de controle e Geladeiras (Salvo como Lista JSON)
                locais_controle_geladeiras_json JSONB,

                -- Comercialização de Produtos
                embalagem_produtos TEXT,
                permitidos_venda_produtos TEXT,

                -- Medicamentos Sujeitos a Controle Especial
                dispensa_controle_especial TEXT,
                local_armazenamento_controle_especial TEXT,
                area_segregada_inutilizados_mc TEXT,
                dispensa_retinoicos TEXT,
                data_autorizacao DATE,
                n_edital_visa_autorizacao TEXT,
                sistema_informatizado_autorizacao TEXT,
                tipo_escrituracao TEXT,
                dispensa_lista_a TEXT,
                data_protocolo_rmnra DATE,
                dispensa_lista_b2 TEXT,
                data_protocolo_rmnrb2 DATE,
                data_envio_bmpo_trimestral DATE,
                data_envio_bmpo_anual DATE,

                -- Listas Dinâmicas: Medicamentos e Antimicrobianos Verificados (Listas JSON)
                medicamentos_controlados_verificados_json JSONB,
                antimicrobianos_verificados_json JSONB,

                -- Inutilização e SNGPC
                possui_termo_inutilizacao_visa TEXT,
                data_termo_inutilizacao_visa DATE,
                faz_transmissao_sngpc TEXT,
                data__transmissao_sngpc DATE,
                periodo_transmissao_sngpc TEXT,
                validacao_transmissao_sngpc TEXT,

                -- Área Física: Administração
                possui_adm TEXT,
                condicoes_estruturais_adm TEXT,
                condicoes_higienicas_adm TEXT,
                protecao_pragas_adm TEXT,
                ventilacao_adm TEXT,
                iluminacao_adm TEXT,

                -- Área Física: Recebimento
                possui_recebimento TEXT,
                condicoes_estruturais_recebimento TEXT,
                condicoes_higienicas_recebimento TEXT,
                controle_recebimento TEXT,

                -- Área Física: Estoque
                possui_estoque TEXT,
                condicoes_estruturais_estoque TEXT,
                piso_estoque TEXT,
                parede_estoque TEXT,
                teto_estoque TEXT,
                condicoes_higienicas_estoque TEXT,
                protecao_pragas_estoque TEXT,
                ventilacao_estoque TEXT,
                iluminacao_estoque TEXT,
                afastamento_structure_estoque TEXT,
                produtos_nao_conformes_estoque TEXT,
                coleta_medicamentos_clientes_estoque TEXT,

                -- Área Física: DML
                possui_dml TEXT,
                condicoes_estruturais_dml TEXT,
                condicoes_higienicas_dml TEXT,
                produto_regularizados_dml TEXT,

                -- Área Física: Sanitários Clientes
                possui_sanitarios_clientes TEXT,
                condicoes_estruturais_sanitarios_clientes TEXT,
                condicoes_higienicas_sanitarios_clientes TEXT,
                lavatorio_sanitarios_clientes TEXT,
                papel_toalha_sanitarios_clientes TEXT,
                sabonete_sanitarios_clientes TEXT,
                lixeira_sanitarios_clientes TEXT,

                -- Área Física: Sanitários Funcionários
                possui_sanitarios_funcionarios TEXT,
                condicoes_estruturais_sanitarios_funcionarios TEXT,
                condicoes_higienicas_sanitarios_funcionarios TEXT,
                lavatorio_sanitarios_funcionarios TEXT,
                papel_toalha_sanitarios_funcionarios TEXT,
                sabonete_sanitarios_funcionarios TEXT,
                lixeira_sanitarios_funcionarios TEXT,
                local_pertences_sanitarios_funcionarios TEXT,

                -- Área Física: Copa
                possui_copa TEXT,
                separacao_copa TEXT,
                condicoes_estruturais_copa TEXT,
                condicoes_higienicas_copa TEXT,

                -- Área Física: Sala de Serviços
                possui_sala_servicos TEXT,
                servicos_realizados_sala_servicos TEXT,
                procedimentos_escritos_sala_servicos TEXT,
                separacao_sala_servicos TEXT,
                condicoes_estruturais_sala_servicos TEXT,
                piso_sala_servicos TEXT,
                parede_sala_servicos TEXT,
                teto_sala_servicos TEXT,
                condicoes_higienicas_sala_servicos TEXT,
                protecao_pragas_sala_servicos TEXT,
                ventilacao_sala_servicos TEXT,
                iluminacao_sala_servicos TEXT,
                pia_sala_servicos TEXT,
                sabonete_sala_servicos TEXT,
                papel_toalha_sala_servicos TEXT,
                caixa_perfurocortante_sala_servicos TEXT,
                lixeira_comum_sala_servicos TEXT,
                lixeira_infectantes_sala_servicos TEXT,
                alcool_sala_servicos TEXT,
                primeiros_socorros_sala_servicos TEXT,
                lista_estab_saude_sala_servicos TEXT,
                registro_limpeza_sala_servicos TEXT,
                declaracao_serviços_sala_servicos TEXT,

                -- Equipamento: Pressão Arterial
                pressao_arterial TEXT,
                marca_modelo_pressao_arterial TEXT,
                possui_calibracao_pressao_arterial TEXT,
                nota_aquisicao_pressao_arterial TEXT,
                data_aquisicao_pressao_arterial DATE,
                empresa_aquisicao_pressao_arterial TEXT,
                certificado_calibracao_pressao_arterial TEXT,
                empresa_calibracao_pressao_arterial TEXT,
                data_calibracao_pressao_arterial DATE,
                validade_calibracao_pressao_arterial DATE,

                -- Equipamento: Temperatura Corporal
                t_corporal TEXT,
                marca_modelo_t_corporal TEXT,

                -- Equipamento: Perfuração de Orelha
                perfuracao_orelha TEXT,
                marca_modelo_perfuracao_orelha TEXT,
                lote_perfuracao_orelha TEXT,

                -- Recursos Humanos / Funcionários
                numero_funcionarios INTEGER,
                uniformes_funcionarios TEXT,
                epis_serviços_funcionarios TEXT,
                capacitacao_funcionarios TEXT
            );
            """))
            s.commit()
    except Exception as e:
        st.sidebar.error(f"Erro crítico ao estruturar a tabela no Supabase: {e}")





























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
                st.session_state["historico_pre_preenchido"] = dict(resultado)
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
#   mapeamento_status = {"Não avaliado": 0, "C": 1, "NC": 2, "NA": 3}
#  opcoes_status = ["Não avaliado", "C", "NC", "NA"]
#
 #   mapeamento_sim_nao = {"Não avaliado": 0, "Sim": 1, "Não": 2}
 #   opcoes_sim_nao = ["Não avaliado", "Sim", "Não"]
 #   
    # Atalho para o dicionário do histórico mestre carregado na Tela 1
    historico = st.session_state.get("historico_pre_preenchido", {})
    
    # Dicionário temporário que vai coletar os dados modificados nesta tela
    dados_inspecao_atual = {
        "cnpj": cnpj_estabelecimento,
        "web": n_web,
        "data_procedimento": datetime.strptime(data_atual, "%d/%m/%Y").date()
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
        st.markdown("### 👥 Pessoas Contatadas / Acompanhantes da Inspeção")
        
        # --- MOTOR DE RECONSTRUÇÃO DO HISTÓRICO ---
        # Se o fiscal mudou de empresa ou iniciou o formulário agora, preparamos a lista de pessoas
        if "lista_pessoas_formulario" not in st.session_state or st.session_state.get("ultimo_cnpj_visto") != cnpj_estabelecimento:
            st.session_state["ultimo_cnpj_visto"] = cnpj_estabelecimento
            texto_banco_contatos = str(historico.get("acompanhantes_inspecao", ""))
            
            # Se existirem dados anteriores gravados na última inspeção, reconstrói a lista
            if texto_banco_contatos.strip() != "" and texto_banco_contatos != "None" and " | " in texto_banco_contatos:
                lista_reconstruida = []
                # Quebra os registros de cada pessoa separados pela barra vertical com espaços
                pessoas_registradas = texto_banco_contatos.split(" | ")
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

    # ==============================================================================
    # TAB 3: CONTROLES T E U
    # ==============================================================================
    with tab_controles:
        st.write("Aguardando inserção do loop dinâmico de termo-higrômetros...")

    # ==============================================================================
    # TAB 4: COMERCIALIZAÇÃO
    # ==============================================================================
    with tab_comercializacao:
        st.write("Aguardando inserção dos campos manuais de produtos e controle especial...")

    # ==============================================================================
    # TAB 5: ÁREAS FÍSICAS
    # ==============================================================================
    with tab_areas:
        st.write("Aguardando inserção das avaliações de estrutura física de cada setor...")

    # ==============================================================================
    # TAB 6: SALA DE SERVIÇOS
    # ==============================================================================
    with tab_sala:
        st.write("Aguardando inserção das regras de serviços farmacêuticos...")

    # ==============================================================================
    # TAB 7: FUNCIONÁRIOS
    # ==============================================================================
    with tab_funcionarios:
        st.write("Aguardando inserção dos dados de EPIs, uniformes e capacitações...")

    









   
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
