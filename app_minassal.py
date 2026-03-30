import streamlit as st
import pandas as pd
import os

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Coleta Minassal - Royal Canin", page_icon="👑", layout="wide")

st.title("📱 Portal de Coleta - Royal Canin")
st.markdown("Bem-vindo ao sistema de auditoria de preços.")

# --- Nomes dos arquivos fixos (Devem estar na exata mesma pasta do app) ---
ARQUIVO_FIXO_VENDAS = "Vendas.xlsx"
ARQUIVO_FIXO_MG = "Tabela_MG.xlsx"
ARQUIVO_FIXO_SP = "Tabela_SP.xlsx"

# --- FUNÇÕES DE LIMPEZA ---
def tratar_codigo(valor):
    try: return str(int(float(valor))).strip()
    except: return str(valor).strip()

@st.cache_data
def carregar_dados(file_or_path):
    if isinstance(file_or_path, str) and not os.path.exists(file_or_path):
        return pd.DataFrame()
    df = pd.read_excel(file_or_path)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def extrair_mapas(df_tabela):
    mapa_c, mapa_e = {}, {}
    if df_tabela.empty: return mapa_c, mapa_e
    
    if 'CODIGO' in df_tabela.columns:
        df_tabela['CODIGO'] = df_tabela['CODIGO'].apply(tratar_codigo)
        
    col_p = next((c for c in df_tabela.columns if "SUGESTÃO" in c or "RECOMENDADO" in c or "SUGESTAO" in c), None)
    col_e = next((c for c in df_tabela.columns if "EAN" in c or "BARRAS" in c), None)
    
    if col_p:
        if 'CODIGO' in df_tabela.columns:
            mapa_c = pd.Series(df_tabela[col_p].values, index=df_tabela['CODIGO']).to_dict()
        if col_e:
            mapa_e = pd.Series(df_tabela[col_p].values, index=df_tabela[col_e].apply(tratar_codigo)).to_dict()
            
    return mapa_c, mapa_e

# --- BARRA LATERAL (AGORA SÓ PARA EMERGÊNCIAS/TESTES) ---
with st.sidebar:
    st.header("⚙️ Atualização Manual")
    st.caption("O sistema já lê os arquivos da pasta sozinho. Use os botões abaixo APENAS se precisar sobrepor algum arquivo no meio do dia.")
    
    file_vendas_up = st.file_uploader("Sobrepor Vendas", type=["xlsx"])
    file_mg_up = st.file_uploader("Sobrepor Tabela MG", type=["xlsx"])
    file_sp_up = st.file_uploader("Sobrepor Tabela SP", type=["xlsx"])
    
    st.divider()
    st.subheader("⚙️ Regras de Exibição")
    mostrar_sem_preco = st.checkbox("Mostrar produtos que NÃO estão nas tabelas", value=False)

# --- CARREGAMENTO AUTOMÁTICO (OU MANUAL SE ENVIADO) ---
df_vendas = carregar_dados(file_vendas_up) if file_vendas_up else carregar_dados(ARQUIVO_FIXO_VENDAS)
df_tabela_mg = carregar_dados(file_mg_up) if file_mg_up else carregar_dados(ARQUIVO_FIXO_MG)
df_tabela_sp = carregar_dados(file_sp_up) if file_sp_up else carregar_dados(ARQUIVO_FIXO_SP)

# Painel visual mostrando o que o sistema conseguiu ler sozinho da pasta
st.write("### 📂 Status da Base de Dados")
cols_status = st.columns(3)
cols_status[0].info("✅ Vendas Carregadas" if not df_vendas.empty else "❌ Arquivo Vendas.xlsx Ausente")
cols_status[1].success("✅ Tabela MG Ativa" if not df_tabela_mg.empty else "❌ Tabela_MG.xlsx Ausente")
cols_status[2].success("✅ Tabela SP Ativa" if not df_tabela_sp.empty else "❌ Tabela_SP.xlsx Ausente")

# --- MOTOR PRINCIPAL ---
if not df_vendas.empty:
    if 'PRODUTO CODIGO' in df_vendas.columns:
        df_vendas['PRODUTO CODIGO'] = df_vendas['PRODUTO CODIGO'].apply(tratar_codigo)

    col_ean_vendas = next((c for c in df_vendas.columns if "EAN" in c or "BARRAS" in c), None)
    clientes = sorted(df_vendas['CLIENTE NOME'].dropna().unique())

    # --- ÁREA DO PROMOTOR ---
    st.divider()
    col1, col2 = st.columns([1, 2])
    with col1:
        estado_selecionado = st.radio("📍 Região da Loja:", ["Minas Gerais (MG)", "São Paulo (SP)"])
        cliente_selecionado = st.selectbox("🏪 Selecione a Loja para Auditoria:", ["-- Selecione --"] + clientes)

    mapa_cod_ativo, mapa_ean_ativo = {}, {}
    if estado_selecionado == "Minas Gerais (MG)":
        mapa_cod_ativo, mapa_ean_ativo = extrair_mapas(df_tabela_mg)
    else:
        mapa_cod_ativo, mapa_ean_ativo = extrair_mapas(df_tabela_sp)

    if cliente_selecionado and cliente_selecionado != "-- Selecione --":
        st.subheader(f"🛒 Produtos Auditados: {cliente_selecionado}")

        df_cliente = df_vendas[df_vendas['CLIENTE NOME'] == cliente_selecionado].drop_duplicates(subset=['PRODUTO CODIGO'])

        dados_tela = []
        for _, linha in df_cliente.iterrows():
            sku = tratar_codigo(linha.get('PRODUTO CODIGO', ''))
            ean_venda = tratar_codigo(linha.get(col_ean_vendas, '')) if col_ean_vendas else ''
            nome_p = str(linha.get('PRODUTO NOME', ''))

            valor = mapa_cod_ativo.get(sku)
            if valor is None and ean_venda:
                valor = mapa_ean_ativo.get(ean_venda)

            if valor is not None:
                try:
                    if float(valor) <= 0:
                        if not mostrar_sem_preco: continue
                        txt_preco = "Preço Zero na Tabela"
                    else:
                        txt_preco = f"R$ {float(valor):.2f}"
                except: txt_preco = str(valor)
            else:
                if not mostrar_sem_preco: continue
                txt_preco = "Fora da Tabela"

            dados_tela.append({
                "CÓDIGO": sku,
                "PRODUTO": nome_p,
                "SUGERIDO": txt_preco,
                "PREÇO NA LOJA": None 
            })

        df_tela = pd.DataFrame(dados_tela)

        if not df_tela.empty:
            df_editado = st.data_editor(
                df_tela,
                use_container_width=True,
                hide_index=True,
                disabled=["CÓDIGO", "PRODUTO", "SUGERIDO"],
                column_config={
                    "PREÇO NA LOJA": st.column_config.NumberColumn("PREÇO NA LOJA", format="R$ %.2f", min_value=0.0)
                }
            )
            
            if st.button("💾 ENVIAR COLETA", type="primary"):
                st.success(f"✅ Coleta da loja {cliente_selecionado} enviada com sucesso!")
        else:
            st.warning("⚠️ Nenhum produto encontrado. Verifique se a tabela de preços está correta ou ative a opção de mostrar produtos fora da tabela na barra lateral.")
else:
    st.info("👈 Por favor, coloque o arquivo 'Vendas.xlsx' na mesma pasta do sistema para começar.")