import streamlit as st
import pandas as pd
import os
import smtplib
import unicodedata
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta

# --- BIBLIOTECAS PARA O PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONFIGURAÇÃO DA PÁGINA E DESIGN DARK ---
st.set_page_config(page_title="Coleta Minassal", page_icon="👑", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
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

# --- LISTA MESTRA: SELEÇÃO DE OURO ---
CODIGOS_OURO = {"97996", "98018", "98224", "98230", "98435", "97985", "98037", "98011", "98015", "98139", "98157", "98492", "98834", "99101", "98022", "97991", "98019", "97994", "98016", "98222", "98197", "98433", "97983", "98122", "98126", "98124", "98144", "98137", "98469", "98467", "98640", "98518", "98520", "98490", "99757", "99753", "99750", "98249", "99187", "98328", "98357", "98331", "98350", "98364", "98334", "98361", "98338", "98434", "98327", "98340", "98365", "98360", "98333", "98353", "98332", "98356", "98336", "98450", "98461", "98589", "98452", "98639", "98631", "98491", "98489", "98719", "98721", "98852", "98024", "97993", "98021"}

# --- FUNÇÃO PARA BUSCAR ARQUIVOS ---
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
    "Fernanda": ["JUIZ DE FORA", "JUIZ DE FORA/MG"]
}

# --- FUNÇÃO DE GERAÇÃO DE PDF ---
def gerar_pdf_relatorio(promotor, loja, cidade, estado, df_preenchido):
    hora_brasil = datetime.now() - timedelta(hours=3)
    data_str = hora_brasil.strftime('%d/%m/%Y %H:%M')
    
    caminho_pdf = f"Auditoria_{loja[:10].replace(' ', '_')}.pdf"
    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4)
    estilos = getSampleStyleSheet()
    elementos = []
    
    elementos.append(Paragraph(f"<b>RELATÓRIO DE AUDITORIA - ROYAL CANIN</b>", estilos['Title']))
    elementos.append(Paragraph(f"<b>LOJA:</b> {loja}", estilos['Heading2']))
    elementos.append(Paragraph(f"<b>PROMOTOR(A):</b> {promotor} | <b>CIDADE:</b> {cidade} | <b>DATA:</b> {data_str}", estilos['Normal']))
    elementos.append(Spacer(1, 15))
    
    data = [["PRODUTO", "CÓDIGO", "PREÇO SUG.", "PREÇO LOJA", "SITUAÇÃO"]]
    estilo_tabela = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2001A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTSIZE', (0,0), (-1,-1), 8)
    ]

    def limpar_valor(v):
        if pd.isna(v) or v == "" or str(v).lower() == "none": return 0.0
        s = str(v).replace("R$", "").replace(" ", "").strip()
        
        # LÓGICA PARA MANTER PONTO OU VÍRGULA SEM MULTIPLICAR
        if "," in s and "." in s: # Caso tenha os dois (ex: 1.250,50)
            s = s.replace(".", "").replace(",", ".")
        elif "," in s: # Caso tenha apenas vírgula (ex: 61,11)
            s = s.replace(",", ".")
        # Se tiver apenas ponto, o float(s) já resolve corretamente (ex: 61.11)
        
        try: return float(s)
        except: return 0.0

    for i, linha in enumerate(df_preenchido.itertuples()):
        idx = i + 1
        nome = str(linha.PRODUTO).replace("⭐ ", "")[:35]
        cod = str(linha.CÓDIGO)
        
        # Limpeza do sugerido
        p_sug = limpar_valor(linha.SUGERIDO)
        p_loja = limpar_valor(getattr(linha, "_4"))

        if p_loja <= 0:
            sit, cor = "OPORTUNIDADE", colors.orange
        elif p_loja <= (p_sug + 0.05):
            sit, cor = "CORRETO", colors.green
        else:
            diff = ((p_loja / p_sug) - 1) * 100
            sit, cor = (f"ACIMA {diff:.0f}%", colors.red) if diff >= 1 else ("CORRETO", colors.green)
                
        data.append([nome, cod, f"R$ {p_sug:.2f}", f"R$ {p_loja:.2f}" if p_loja > 0 else "--", sit])
        estilo_tabela.append(('TEXTCOLOR', (4, idx), (4, idx), cor))
        if cod in CODIGOS_OURO: 
            estilo_tabela.append(('BACKGROUND', (0, idx), (0, idx), colors.HexColor("#FEF3C7")))

    t = Table(data, colWidths=[200, 60, 80, 80, 100])
    t.setStyle(TableStyle(estilo_tabela))
    elementos.append(t)
    doc.build(elementos)
    return caminho_pdf

# --- ENVIO DE EMAIL ---
def enviar_email_coleta(promotor, loja, cidade, estado, df_editado, feedback):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = ["benedito.bandola@minassal.com.br"]

    caminho_pdf = gerar_pdf_relatorio(promotor, loja, cidade, estado, df_editado)
    
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, ", ".join(destino), f"✅ Auditoria PDV - {loja} ({promotor})"
    corpo = f"<html><body><h2 style='color: #E2001A;'>Auditoria Royal Canin Recebida</h2><p><b>Loja:</b> {loja}<br><b>Promotor:</b> {promotor}<br><b>Cidade:</b> {cidade}</p><hr><p><b>Observações:</b><br>{feedback if feedback.strip() else 'Sem comentários.'}</p></body></html>"
    msg.attach(MIMEText(corpo, 'html'))

    try:
        with open(caminho_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=os.path.basename(caminho_pdf))
            msg.attach(anexo)
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls(); s.login(remetente, senha); s.sendmail(remetente, destino, msg.as_string()); s.quit()
        os.remove(caminho_pdf)
        return True, "Relatório enviado com sucesso!"
    except Exception as e: return False, f"Erro: {e}"

# --- CARREGAR DADOS ---
@st.cache_data
def carregar_dados(caminho):
    if not caminho: return pd.DataFrame()
    try:
        if caminho.endswith('.csv'): df = pd.read_csv(caminho, sep=None, engine='python', encoding='utf-8-sig')
        else: df = pd.read_excel(caminho)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df
    except: return pd.DataFrame()

# --- INTERFACE ---
if 'promotor_logado' not in st.session_state: st.session_state.promotor_logado = None

vendas = carregar_dados(ARQUIVO_VENDAS)
tab_mg = carregar_dados(ARQUIVO_MG)
tab_sp = carregar_dados(ARQUIVO_SP)

if not vendas.empty:
    if st.session_state.promotor_logado is None:
        st.subheader("Selecione o seu Perfil")
        c1, c2 = st.columns(2)
        if c1.button("👩‍💼 PAMELA", use_container_width=True): st.session_state.promotor_logado = "Pamela"; st.rerun()
        if c2.button("👩‍💼 FERNANDA", use_container_width=True): st.session_state.promotor_logado = "Fernanda"; st.rerun()
    else:
        promotor = st.session_state.promotor_logado
        st.sidebar.info(f"👤 {promotor}")
        if st.sidebar.button("Sair"): st.session_state.promotor_logado = None; st.rerun()

        cidades_autorizadas = ROTAS_PROMOTORES[promotor]
        vendas['CIDADE_LIMPA'] = vendas['CIDADE'].astype(str).str.upper().str.strip()
        df_f = vendas[vendas['CIDADE_LIMPA'].isin(cidades_autorizadas)]
        
        loja_sel = st.selectbox("🏪 Selecione a Loja:", ["-- Selecione --"] + sorted(df_f['CLIENTE NOME'].unique()))
        regiao = st.radio("Região:", ["Minas Gerais (MG)", "São Paulo (SP)"], horizontal=True)

        tab_ativa = tab_mg if "Minas" in regiao else tab_sp
        mapa_precos = {}
        if not tab_ativa.empty:
            col_preco = next((c for c in tab_ativa.columns if "SUGEST" in c or "RECOMEN" in c), None)
            mapa_precos = pd.Series(tab_ativa[col_preco].values, index=tab_ativa['CODIGO'].astype(str).str.strip()).to_dict()

        if loja_sel != "-- Selecione --":
            df_loja = df_f[df_f['CLIENTE NOME'] == loja_sel].drop_duplicates(subset=['PRODUTO CODIGO'])
            dados_tabela = []
            for _, r in df_loja.iterrows():
                cod = str(r['PRODUTO CODIGO']).strip()
                p_sug = mapa_precos.get(cod, 0.0)
                if p_sug > 0:
                    dados_tabela.append({"CÓDIGO": cod, "PRODUTO": ("⭐ " if cod in CODIGOS_OURO else "") + str(r['PRODUTO NOME']), "SUGERIDO": f"R$ {float(p_sug):.2f}", "PREÇO NA LOJA": 0.0})

            if dados_tabela:
                df_editor = st.data_editor(pd.DataFrame(dados_tabela), use_container_width=True, hide_index=True, disabled=["CÓDIGO", "PRODUTO", "SUGERIDO"])
                obs = st.text_area("Inteligência de Campo:")
                if st.button("🚀 ENVIAR AUDITORIA", use_container_width=True):
                    cid_final = df_f[df_f['CLIENTE NOME'] == loja_sel]['CIDADE'].iloc[0]
                    ok, res = enviar_email_coleta(promotor, loja_sel, cid_final, regiao, df_editor, obs)
                    if ok: st.success(res); st.balloons()
                    else: st.error(res)
else:
    st.error("Arquivo 'Vendas' não encontrado.")
