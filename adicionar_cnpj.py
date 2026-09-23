import streamlit as st
import pandas as pd
import requests
import time
from sqlalchemy import text

# Configuração da página do robô de automação
st.set_page_config(page_title="Robô de Carga - CNPJ", layout="wide")
st.title("🤖 Robô de Automação e Enriquecimento de Dados - CNPJ")
st.markdown("Insira a lista de CNPJs. O sistema buscará os dados na Receita Federal e salvará diretamente no Supabase.")

# --- CONEXÃO COM O SEU BANCO DE DADOS (SUPABASE) ---
conn_nuvem = None
try:
    # Utiliza as mesmas credenciais do seu arquivo secrets.toml
    conn_nuvem = st.connection("postgresql", type="sql")
except Exception as e:
    st.error(f"Erro ao conectar ao Supabase: {e}")

# Caixa de texto grande para colar os CNPJs (um por linha)
st.markdown("### 📋 Digite ou cole os CNPJs (um por linha):")
cnpjs_colados = st.text_area("Lista de CNPJs", height=200, label_visibility="collapsed", placeholder="Ex:\n27865757000102\n47508411000156")

if st.button("🚀 Iniciar Captura e Carga no Supabase", type="primary"):
    if not conn_nuvem:
        st.error("Banco de dados não conectado. Verifique suas credenciais.")
    elif not cnpjs_colados.strip():
        st.warning("Por favor, insira pelo menos um CNPJ para começar.")
    else:
        # Separa o texto por quebras de linha e limpa espaços e caracteres não numéricos
        lista_linhas = cnpjs_colados.split("\n")
        lista_cnpjs = []
        for linha in lista_linhas:
            cnpj_limpo = "".join(filter(str.isdigit, linha))
            if len(cnpj_limpo) == 14:
                lista_cnpjs.append(cnpj_limpo)
        
        if not lista_cnpjs:
            st.error("❌ Nenhum CNPJ válido com 14 dígitos foi encontrado na lista.")
        else:
            st.info(f"Total de {len(lista_cnpjs)} CNPJs válidos identificados. Iniciando processamento...")
            
            # Cria barras de progresso na tela do Streamlit
            barra_progresso = st.progress(0)
            status_texto = st.empty()
            
            sucessos = 0
            erros = 0
            
            for indice, cnpj in enumerate(lista_cnpjs):
                # Atualiza a interface visual para o usuário
                percentual = int((indice / len(lista_cnpjs)) * 100)
                barra_progresso.progress(percentual)
                status_texto.markdown(f"🔄 **Processando ({indice + 1}/{len(lista_cnpjs)}):** Consultando CNPJ {cnpj}...")
                
                try:
                    # Consulta a API pública da BrasilAPI (Sem limite agressivo de travas por minuto)
                    url_api = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
                    resposta = requests.get(url_api, timeout=10)
                    
                    if resposta.status_code == 200:
                        dados = resposta.json()
                        
                        # Extração e normalização dos dados cadastrais
                        razao_social = dados.get("razao_social") or dados.get("nome_fantasia") or "Não Informado"
                        
                        # Montagem amigável do endereço (Rua, Número, Complemento)
                        logradouro = dados.get("logradouro") or ""
                        numero = dados.get("numero") or ""
                        complemento = dados.get("complemento") or ""
                        endereco_completo = f"{logradouro}, Nº {numero}"
                        if complemento and complemento != "NONE":
                            endereco_completo += f" - {complemento}"
                            
                        bairro = dados.get("bairro") or "Não Informado"
                        cep = dados.get("cep") or "Não Informado"
                        cnae = str(dados.get("cnae_fiscal") or "Não Informado")
                        atividade = dados.get("cnae_fiscal_descricao") or "Comércio Varejista Farmacêutico"
                        
                        # Inserção direta e robusta no seu Supabase próprio
                        with conn_nuvem.session as session:
                            session.execute(
                                text("""
                                INSERT INTO empresas_db (cnpj, razao_social, endereco, bairro, cep, cnae, atividade)
                                VALUES (:cnpj, :razao_social, :endereco, :bairro, :cep, :cnae, :atividade)
                                ON CONFLICT (cnpj) DO UPDATE SET
                                    razao_social = EXCLUDED.razao_social,
                                    endereco = EXCLUDED.endereco,
                                    bairro = EXCLUDED.bairro,
                                    cep = EXCLUDED.cep,
                                    cnae = EXCLUDED.cnae,
                                    atividade = EXCLUDED.atividade;
                                """),
                                {
                                    "cnpj": cnpj,
                                    "razao_social": razao_social,
                                    "endereco": endereco_completo,
                                    "bairro": bairro,
                                    "cep": cep,
                                    "cnae": cnae,
                                    "atividade": atividade
                                }
                            )
                            session.commit()
                        
                        st.write(f"✅ CNPJ {cnpj} carregado com sucesso: **{razao_social}**")
                        sucessos += 1
                    else:
                        st.write(f"❌ Falha no CNPJ {cnpj}: Código de retorno da API {resposta.status_code}")
                        erros += 1
                        
                except Exception as e_api:
                    st.write(f"⚠️ Erro ao processar o CNPJ {cnpj}: {e_api}")
                    erros += 1
                
                # Pequena pausa de 1.5 segundos entre requisições para respeitar os servidores públicos
                time.sleep(1.5)
            
            # Finalização do processo
            barra_progresso.progress(100)
            status_texto.markdown("### 🏁 Processamento Concluído!")
            st.success(f"✨ Carga finalizada! {sucessos} empresas adicionadas/atualizadas no Supabase com sucesso. Total de erros: {erros}.")
