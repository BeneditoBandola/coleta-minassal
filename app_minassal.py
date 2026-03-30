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

# --- CONFIGURAÇÃO DA PÁGINA E DESIGN DARK (PRETO) ---
st.set_page_config(page_title="Coleta Minassal", page_icon="👑", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    div.stButton > button {
        height: 60px; font-size: 18px; font-weight: bold; border-radius: 8px;
        border: 2px solid #E2001A; color: #FFFFFF; background-color: #1A1A1A;
    }
    div.stButton > button:hover { background-color: #E2001A; color: white; }
    </style>
""", unsafe_allow_html=True)

st.title("📱 Portal de Auditoria")
st.markdown("---")

# --- LISTA MESTRA: SELEÇÃO DE OURO ---
CODIGOS_OURO = {"97996", "98018", "98224", "98230", "98435", "97985", "98037", "98011", "98015", "98139", "98157", "98492", "98834", "99101", "98022", "97991", "98019", "97994", "98016", "98222", "98197", "98433", "97983", "98122", "98126", "98124", "98144", "98137", "98469", "98467", "98640", "98518", "98520", "98490", "99757", "99753", "99750", "98249", "99187", "98328", "98357", "98331", "98350", "98364", "98334", "98361", "98338", "98434", "98327", "98340", "98365", "98360", "98333", "98353", "98332", "98356", "98336", "98450", "98461", "98589", "98452", "98639", "98631", "98491", "98489", "98719", "98721", "98852", "98024", "97993", "98021"}

# --- NOMES DOS ARQUIVOS (BUSCA INTELIGENTE) ---
def buscar_arquivo(nome_base):
    for ext in [".xlsx", ".csv", ".xlsx.xlsx"]:
        caminho = nome_base + ext
        if os.path.exists(caminho): return caminho
    return None

ARQUIVO_VENDAS = buscar_arquivo("Vendas")
ARQUIVO_MG = buscar_arquivo("Tabela_MG")
ARQUIVO_SP = buscar_arquivo("Tabela_SP")

ROTAS_PROMOTORES = {
    "Pamela": ["POÇOS DE CALDAS", "POCOS DE CALDAS", "ANDRADAS", "VARGINHA", "TRÊS CORAÇÕES", "TRES CORACOES", "TRÊS PONTAS", "TRES PONTAS", "ITAJUBÁ", "ITAJUBA", "POUSO ALEGRE"],
    "Fernanda": ["JUIZ DE FORA", "JUIZ DE FORA/MG"]
}

# --- FUNÇÃO DE GERAÇÃO DE PDF COM HORÁRIO BRASIL ---
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
    
    data = [["PRODUTO", "CÓDIGO", "PREÇO SUG.", "PREÇO LOJA", "RESULTADO"]]
    estilo_tabela = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2001A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
    ]

    for i, linha in enumerate(df_preenchido.itertuples()):
        idx = i + 1
        nome = str(linha.PRODUTO).replace("⭐ ", "")[:40]
        cod = str(linha.CÓDIGO)
        
        # Limpa o preço sugerido (remove R$ e trata decimal)
        try: 
            sug_limpo = str(linha.SUGERIDO).replace("R$ ", "").replace(".", "").replace(",", ".")
            p_sug = float(sug_limpo)
        except: p_sug = 0.0
        
        # Limpa o preço da loja (trata a vírgula digitada)
        try: 
            loja_limpo = str(getattr(linha, "_4")).replace(",", ".")
            p_loja = float(loja_limpo)
        except: p_loja = 0.0

        if p_loja <= 0:
            texto_res, cor_texto, txt_loja = "OPORTUNIDADE", colors.orange, "--"
        elif p_loja <= p_sug:
            texto_res, cor_texto = "CORRETO", colors.green
        else:
            diff = ((p_loja / p_sug) - 1) * 100
            texto_res, cor_texto = (f"ACIMA {diff:.0f}%", colors.red) if diff >= 1 else ("CORRETO", colors.green)
                
        data.append([nome, cod, f"R$ {p_sug:.2f}", f"R$ {p_loja:.2f}" if p_loja > 0 else "--", texto_res])
        estilo_tabela.append(('TEXTCOLOR', (4, idx), (4, idx), cor_texto))
        if cod in CODIGOS_OURO: estilo_tabela.append(('BACKGROUND', (0, idx), (3, idx), colors.HexColor("#FEF3C7")))

    t = Table(data, colWidths=[230, 60, 75, 75, 80])
    t.setStyle(TableStyle(estilo_tabela))
    elementos.append(t)
    doc.build(elementos)
    return caminho_pdf

# --- ENVIO DE EMAIL CONFIGURADO ---
def enviar_email_coleta(promotor, loja, cidade, estado, df_editado, feedback_promotor):
    email_remetente = "beneditobandola@gmail.com" 
    senha_remetente = "kfih ccqx cskn oito" 
    emails_destino = ["benedito.bandola@minassal.com.br"] 

    df_preenchido = df_editado[df_editado["PREÇO NA LOJA"].notna()]
    if df_preenchido.empty: return False, "Nenhum preço preenchido."

    caminho_pdf = gerar_pdf_relatorio(promotor, loja, cidade, estado, df_preenchido)
    feedback_txt = f"\"{feedback_promotor}\"" if feedback_promotor.strip() != "" else "<i>Sem comentários.</i>"

    corpo_email = f"<html><body><h2 style='color: #E2001A;'>Nova Auditoria Recebida</h2><p><b>Loja:</b> {loja}</p><p><b>Promotor:</b> {promotor}</p><hr><h3>💬 Inteligência de Campo:</h3><p>{feedback_txt}</p></body></html>"

    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = email_remetente, ", ".join(emails_destino), f"✅ Auditoria PDV - {loja} ({promotor})"
    msg.attach(MIMEText(corpo_email, 'html'))

    try:
        with open(caminho_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=caminho_pdf)
            msg.attach(anexo)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, emails_destino, msg.as_string())
        server.quit()
        os.remove(caminho_pdf)
        return True, "Relatório enviado com sucesso!"
    except Exception as e: return False, f"Erro ao enviar: {e}"

# --- CARREGAR DADOS ---
@st.cache_data
def carregar_dados(caminho):
    if not caminho: return pd.DataFrame()
    if caminho.endswith('.csv'): df = pd.read_csv(caminho, sep=None, engine='python', encoding='utf-8-sig')
    else: df = pd.read_excel(caminho)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def extrair_mapas(df_tab):
    if df_tab.empty: return {}, {}
    if 'CODIGO' in df_tab.columns: df_tab['CODIGO'] = df_tab['CODIGO'].astype(str).str.strip()
    col_p = next((c for c in df_tab.columns if "SUGEST" in c or "RECOMEN" in c), None)
    return pd.Series(df_tab[col_p].values, index=df_tab['CODIGO']).to_dict(), {}

# --- LÓGICA DO APP ---
if 'promotor_logado' not in st.session_state: st.session_state.promotor_logado = None

df_vendas = carregar_dados(ARQUIVO_VENDAS)
df_tab_mg = carregar_dados(ARQUIVO_MG)
df_tab_sp = carregar_dados(ARQUIVO_SP)

if not df_vendas.empty:
    if st.session_state.promotor_logado is None:
        st.subheader("Selecione o seu Perfil")
        c1, c2 = st.columns(2)
        if c1.button("👩‍💼 PAMELA", use_container_width=True): st.session_state.promotor_logado = "Pamela"; st.rerun()
        if c2.button("👩‍💼 FERNANDA", use_container_width=True): st.session_state.promotor_logado = "Fernanda"; st.rerun()
    else:
        promotor = st.session_state.promotor_logado
        col_n, col_s = st.columns([4, 1])
        col_n.info(f"👤 Logado como: **{promotor}**")
        if col_s.button("Sair"): st.session_state.promotor_logado = None; st.rerun()
        
        cidades = ROTAS_PROMOTORES[promotor]
        df_vendas['CIDADE_BUSCA'] = df_vendas['CIDADE'].astype(str).str.upper().str.strip()
        df_f = df_vendas[df_vendas['CIDADE_BUSCA'].isin(cidades)]
        clientes = sorted(df_f['CLIENTE NOME'].dropna().unique())

        colA, colB = st.columns([1, 2])
        est = colA.radio("Região:", ["Minas Gerais (MG)", "São Paulo (SP)"])
        loja = colB.selectbox("🏪 Selecione a Loja:", ["-- Selecione --"] + clientes)

        mapa_p, _ = extrair_mapas(df_tab_mg if "Minas" in est else df_tab_sp)

        if loja != "-- Selecione --":
            st.divider()
            df_c = df_f[df_f['CLIENTE NOME'] == loja].drop_duplicates(subset=['PRODUTO CODIGO'])
            dados = []
            for _, r in df_c.iterrows():
                sku = str(r.get('PRODUTO CODIGO', '')).strip()
                nome = ("⭐ " if sku in CODIGOS_OURO else "") + str(r.get('PRODUTO NOME', ''))
                val = mapa_p.get(sku)
                if val: dados.append({"CÓDIGO": sku, "PRODUTO": nome, "SUGERIDO": f"R$ {float(val):.2f}", "PREÇO NA LOJA": None})

            if dados:
                df_ed = st.data_editor(pd.DataFrame(dados), use_container_width=True, hide_index=True, disabled=["CÓDIGO", "PRODUTO", "SUGERIDO"])
                st.markdown("### 💬 Inteligência de Campo")
                feedback = st.text_area("Na sua opinião, o que poderíamos fazer para melhorar nossa participação neste cliente?", height=100)
                if st.button("🚀 ENVIAR AUDITORIA", type="primary", use_container_width=True):
                    cid_loja = df_f[df_f['CLIENTE NOME'] == loja]['CIDADE'].iloc[0]
                    ok, msg = enviar_email_coleta(promotor, loja, cid_loja, est, df_ed, feedback)
                    if ok: st.success(msg); st.balloons()
                    else: st.error(msg)
            else: st.warning("Nenhum produto da tabela encontrado para este cliente.")
else: st.error("Arquivo de Vendas não encontrado.")
