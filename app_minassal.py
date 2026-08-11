import streamlit as st
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta

# --- BIBLIOTECAS PARA O PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Coleta Minassal", page_icon="👑", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    .stApp { background-color: #0d0d0d; }
    div.stButton > button {
        height: 60px; font-size: 18px; font-weight: bold; border-radius: 8px;
        border: 2px solid #E2001A; color: #FFFFFF; background-color: #1A1A1A;
    }
    div.stButton > button:hover { background-color: #E2001A; color: white; }
    h1, h2, h3, p, span, label { color: white !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📱 Portal de Auditoria - Royal Canin")
st.markdown("---")

CODIGOS_OURO = {"97996", "98018", "98224", "98230", "98435", "97985", "98037", "98011", "98015", "98139", "98157", "98492", "98834", "99101", "98022", "97991", "98019", "97994", "98016", "98222", "98197", "98433", "97983", "98122", "98126", "98124", "98144", "98137", "98469", "98467", "98640", "98518", "98520", "98490", "99757", "99753", "99750", "98249", "99187", "98328", "98357", "98331", "98350", "98364", "98334", "98361", "98338", "98434", "98327", "98340", "98365", "98360", "98333", "98353", "98332", "98356", "98336", "98450", "98461", "98589", "98452", "98639", "98631", "98491", "98489", "98719", "98721", "98852", "98024", "97993", "98021"}

def buscar_arquivo(nome_base):
    for ext in [".csv", ".xlsx"]:
        caminho = nome_base + ext
        if os.path.exists(caminho): return caminho
    return None

ARQUIVO_VENDAS = buscar_arquivo("Vendas")
ARQUIVO_MG = buscar_arquivo("Tabela_MG")
ARQUIVO_SP = buscar_arquivo("Tabela_SP")

ROTAS_PROMOTORES = {
    "Pamela": ["POCOS DE CALDAS", "POÇOS DE CALDAS", "ANDRADAS", "VARGINHA", "TRES CORACOES", "TRÊS CORAÇÕES", "TRES PONTAS", "TRÊS PONTAS", "ITAJUBA", "ITAJUBÁ", "POUSO ALEGRE"],
    "Fernanda": ["JUIZ DE FORA", "JUIZ DE FORA/MG"],
    "Madalla": ["CONSELHEIRO LAFAIETE", "GUARANI", "GUIDOVAL", "MURIAE", "MURIAÉ", "PIRAUBA", "PIRAÚBA", "RIO POMBA", "TOCANTINS", "UBA", "UBÁ", "VICOSA", "VIÇOSA", "VISCONDE DO RIO BRANCO"]
}

@st.cache_data
def carregar_dados(caminho):
    if not caminho: return pd.DataFrame()
    try:
        if caminho.endswith('.csv'): df = pd.read_csv(caminho, sep=None, engine='python', encoding='utf-8-sig')
        else: 
            df = pd.read_excel(caminho)
            # Remove a linha "TOTAL GERAL" se ela existir na primeira linha
            if 'TOTAL GERAL' in str(df.iloc[0,0]): df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df
    except: return pd.DataFrame()

# --- FUNÇÕES DE PDF E EMAIL (Mantidas) ---
def gerar_pdf_relatorio(promotor, loja, cidade, estado, df_preenchido):
    # ... (seu código original de gerar_pdf_relatorio continua funcionando)
    pass 

# ... (seu código de enviar_email_coleta continua funcionando)

# --- LÓGICA DE INTERFACE ---
vendas = carregar_dados(ARQUIVO_VENDAS)
# ... (O restante da sua interface permanece igual, pois ele consome 'CLIENTE NOME' e 'PRODUTO CODIGO', que continuam existindo)
