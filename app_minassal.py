import streamlit as st
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

# --- BIBLIOTECAS PARA O PDF ---
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# --- CONFIGURAÇÃO DA PÁGINA E DESIGN DARK (PRETO) ---
st.set_page_config(page_title="Coleta Minassal", page_icon="👑", layout="wide", initial_sidebar_state="collapsed")

# CSS para forçar um visual limpo, esconder menus e adaptar ao tema escuro
st.markdown("""
    <style>
    /* Esconde o menu superior direito e o rodapé do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Botões corporativos para tema escuro */
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
CODIGOS_OURO = {
    "97996", "98018", "98224", "98230", "98435", "97985", "98037", "98011", 
    "98015", "98139", "98157", "98492", "98834", "99101", "98022", "97991", 
    "98019", "97994", "98016", "98222", "98197", "98433", "97983", "98122", 
    "98126", "98124", "98144", "98137", "98469", "98467", "98640", "98518", 
    "98520", "98490", "99757", "99753", "99750", "98249", "99187", "98328", 
    "98357", "98331", "98350", "98364", "98334", "98361", "98338", "98434", 
    "98327", "98340", "98365", "98360", "98333", "98353", "98332", "98356", 
    "98336", "98450", "98461", "98589", "98452", "98639", "98631", "98491", 
    "98489", "98719", "98721", "98852", "98024", "97993", "98021"
}

# --- Nomes dos arquivos fixos (Devem estar na mesma pasta) ---
ARQUIVO_FIXO_VENDAS = "Vendas.xlsx"
ARQUIVO_FIXO_MG = "Tabela_MG.xlsx"
ARQUIVO_FIXO_SP = "Tabela_SP.xlsx"

ROTAS_PROMOTORES = {
    "Pamela": [
        "POÇOS DE CALDAS", "POCOS DE CALDAS", "ANDRADAS", "VARGINHA", 
        "TRÊS CORAÇÕES", "TRES CORACOES", "TRÊS PONTAS", "TRES PONTAS", 
        "ITAJUBÁ", "ITAJUBA", "POUSO ALEGRE"
    ],
    "Fernanda": [
        "JUIZ DE FORA", "JUIZ DE FORA/MG"
    ]
}

# --- FUNÇÃO DE GERAÇÃO DE PDF (CRÍTICAS) ---
def gerar_pdf_relatorio(promotor, loja, cidade, estado, df_preenchido):
    caminho_pdf = f"Auditoria_{loja[:10].replace(' ', '_')}.pdf"
    
    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4)
    estilos = getSampleStyleSheet()
    elementos = []
    
    elementos.append(Paragraph(f"<b>RELATÓRIO DE AUDITORIA - ROYAL CANIN</b>", estilos['Title']))
    elementos.append(Paragraph(f"<b>LOJA:</b> {loja}", estilos['Heading2']))
    elementos.append(Paragraph(f"<b>PROMOTOR(A):</b> {promotor} | <b>CIDADE:</b> {cidade} | <b>DATA:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilos['Normal']))
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
        
        try: p_sug = float(str(linha.SUGERIDO).replace("R$ ", "").replace(",", "."))
        except: p_sug = 0.0
        
        try: p_loja = float(getattr(linha, "_4"))
        except: p_loja = 0.0

        txt_loja = f"R$ {p_loja:.2f}"
        
        if p_loja <= 0:
            texto_res, cor_texto, txt_loja = "NÃO TEM", colors.black, "--"
        elif p_loja <= p_sug:
            texto_res, cor_texto = "CORRETO", colors.green
        else:
            diff = ((p_loja / p_sug) - 1) * 100
            if diff < 1:
                texto_res, cor_texto = "CORRETO", colors.green
            else:
                texto_res, cor_texto = f"ACIMA {diff:.0f}%", colors.red
                
        data.append([nome, cod, f"R$ {p_sug:.2f}", txt_loja, texto_res])
        estilo_tabela.append(('TEXTCOLOR', (4, idx), (4, idx), cor_texto))
        if texto_res == "CORRETO":
            estilo_tabela.append(('FONTNAME', (4, idx), (4, idx), 'Helvetica-Bold'))
            
        if cod in CODIGOS_OURO:
            estilo_tabela.append(('BACKGROUND', (0, idx), (3, idx), colors.HexColor("#FEF3C7")))

    t = Table(data, colWidths=[230, 60, 75, 75, 80])
    t.setStyle(TableStyle(estilo_tabela))
    elementos.append(t)
    doc.build(elementos)
    
    return caminho_pdf

# --- FUNÇÃO DE ENVIO DE E-MAIL COM ANEXO E FEEDBACK ---
def enviar_email_coleta(promotor, loja, cidade, estado, df_editado, feedback_promotor):
    email_remetente = "beneditobandola@gmail.com" # ⚠️ Coloque o seu e-mail robô aqui
    senha_remetente = "fjhy tgih iypx zpsf" # ⚠️ Coloque a sua senha de app aqui
    emails_destino = ["benedito.bandola@minassal.com.br"] 

    df_preenchido = df_editado[df_editado["PREÇO NA LOJA"].notna()]
    if df_preenchido.empty: return False, "Nenhum preço foi preenchido. E-mail não enviado."

    caminho_pdf = gerar_pdf_relatorio(promotor, loja, cidade, estado, df_preenchido)

    if feedback_promotor.strip() == "":
        texto_feedback = "<i>Nenhum comentário ou sugestão enviado pelo promotor.</i>"
    else:
        texto_feedback = f"<span style='color: #059669; font-size: 16px;'><b>\"{feedback_promotor}\"</b></span>"

    corpo_email = f"""
    <html>
      <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #E2001A;">Nova Auditoria Recebida</h2>
        <p><b>Promotor(a):</b> {promotor}</p>
        <p><b>Loja Auditada:</b> {loja} - {cidade}</p>
        <hr>
        <h3>💬 Opinião do Promotor:</h3>
        <p>O que poderíamos fazer para melhorar nossa participação neste cliente?</p>
        <p>{texto_feedback}</p>
        <hr>
        <p>Encontra-se em anexo o arquivo PDF com a tabela detalhada e as críticas de preços (Correto, Acima, Não Tem).</p>
      </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['From'] = email_remetente
    msg['To'] = ", ".join(emails_destino)
    msg['Subject'] = f"✅ Auditoria PDV - {loja} ({promotor})"
    msg.attach(MIMEText(corpo_email, 'html'))

    try:
        with open(caminho_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=caminho_pdf)
            msg.attach(anexo)
    except Exception as e:
        return False, f"Erro ao anexar PDF: {e}"

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(email_remetente, senha_remetente)
        server.sendmail(email_remetente, emails_destino, msg.as_string())
        server.quit()
        os.remove(caminho_pdf) 
        return True, "Coleta registrada e PDF enviado para o seu e-mail com sucesso!"
    except Exception as e:
        return False, f"Erro ao enviar e-mail: {e}"

# --- FUNÇÕES DE DADOS ---
def tratar_codigo(valor):
    try: return str(int(float(valor))).strip()
    except: return str(valor).strip()

@st.cache_data
def carregar_dados(file_or_path):
    if file_or_path is None or (isinstance(file_or_path, str) and not os.path.exists(file_or_path)):
        return pd.DataFrame()
    df = pd.read_excel(file_or_path)
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def extrair_mapas(df_tabela):
    mapa_c, mapa_e = {}, {}
    if df_tabela.empty: return mapa_c, mapa_e
    if 'CODIGO' in df_tabela.columns: df_tabela['CODIGO'] = df_tabela['CODIGO'].apply(tratar_codigo)
    col_p = next((c for c in df_tabela.columns if "SUGESTÃO" in c or "RECOMENDADO" in c or "SUGESTAO" in c), None)
    col_e = next((c for c in df_tabela.columns if "EAN" in c or "BARRA" in c), None)
    if col_p:
        if 'CODIGO' in df_tabela.columns: mapa_c = pd.Series(df_tabela[col_p].values, index=df_tabela['CODIGO']).to_dict()
        if col_e: mapa_e = pd.Series(df_tabela[col_p].values, index=df_tabela[col_e].apply(tratar_codigo)).to_dict()
    return mapa_c, mapa_e

def pintar_linha_ouro(row):
    # Cores otimizadas para leitura no modo escuro
    if row['CÓDIGO'] in CODIGOS_OURO:
        return ['background-color: #D4AF37; color: black; font-weight: bold'] * len(row)
    return [''] * len(row)

# --- MEMÓRIA DO APLICATIVO ---
if 'promotor_logado' not in st.session_state:
    st.session_state.promotor_logado = None

# Barra lateral em modo "Administrador" (escondida por padrão)
with st.sidebar:
    st.header("🛠️ Modo Administrador")
    st.caption("Apenas para atualizações manuais e emergenciais.")
    with st.expander("Atualizar Bases via Upload"):
        file_vendas_up = st.file_uploader("Substituir Vendas", type=["xlsx"])
        file_mg_up = st.file_uploader("Substituir Tabela MG", type=["xlsx"])
        file_sp_up = st.file_uploader("Substituir Tabela SP", type=["xlsx"])

# O sistema entra lendo os arquivos locais automaticamente!
df_vendas = carregar_dados(file_vendas_up) if file_vendas_up else carregar_dados(ARQUIVO_FIXO_VENDAS)
df_tabela_mg = carregar_dados(file_mg_up) if file_mg_up else carregar_dados(ARQUIVO_FIXO_MG)
df_tabela_sp = carregar_dados(file_sp_up) if file_sp_up else carregar_dados(ARQUIVO_FIXO_SP)

if not df_vendas.empty:
    if 'PRODUTO CODIGO' in df_vendas.columns: df_vendas['PRODUTO CODIGO'] = df_vendas['PRODUTO CODIGO'].apply(tratar_codigo)
    col_ean_vendas = next((c for c in df_vendas.columns if "EAN" in c or "BARRA" in c), None)

    if st.session_state.promotor_logado is None:
        st.subheader("Selecione o seu Perfil")
        col1, col2 = st.columns(2)
        if col1.button("👩‍💼 PAMELA", use_container_width=True):
            st.session_state.promotor_logado = "Pamela"
            st.rerun()
        if col2.button("👩‍💼 FERNANDA", use_container_width=True):
            st.session_state.promotor_logado = "Fernanda"
            st.rerun()
            
    else:
        promotor = st.session_state.promotor_logado
        col_nome, col_sair = st.columns([4, 1])
        col_nome.info(f"👤 Em serviço: **{promotor}**")
        if col_sair.button("Sair", use_container_width=True):
            st.session_state.promotor_logado = None
            st.rerun()

        cidades_da_rota = ROTAS_PROMOTORES[promotor]
        if 'CIDADE' in df_vendas.columns:
            df_vendas['CIDADE_BUSCA'] = df_vendas['CIDADE'].astype(str).str.upper().str.strip()
            df_filtrado = df_vendas[df_vendas['CIDADE_BUSCA'].isin(cidades_da_rota)]
        else:
            df_filtrado = df_vendas

        clientes = sorted(df_filtrado['CLIENTE NOME'].dropna().unique())

        st.subheader("📍 Rota de Coleta")
        colA, colB = st.columns([1, 2])
        with colA: estado_selecionado = st.radio("Região da Loja:", ["Minas Gerais (MG)", "São Paulo (SP)"])
        with colB:
            if len(clientes) == 0: st.warning("Nenhuma rota hoje."); cliente_selecionado = "-- Selecione --"
            else: cliente_selecionado = st.selectbox("🏪 Selecione a Loja:", ["-- Selecione --"] + clientes)

        mapa_cod_ativo, mapa_ean_ativo = extrair_mapas(df_tabela_mg) if estado_selecionado == "Minas Gerais (MG)" else extrair_mapas(df_tabela_sp)

        if cliente_selecionado and cliente_selecionado != "-- Selecione --":
            st.divider()
            st.subheader(f"🛒 Produtos Auditados: {cliente_selecionado}")
            cidade_loja = df_filtrado[df_filtrado['CLIENTE NOME'] == cliente_selecionado]['CIDADE'].iloc[0] if 'CIDADE' in df_filtrado.columns else "N/A"

            df_cliente = df_filtrado[df_filtrado['CLIENTE NOME'] == cliente_selecionado].drop_duplicates(subset=['PRODUTO CODIGO'])

            dados_tela = []
            for _, linha in df_cliente.iterrows():
                sku = tratar_codigo(linha.get('PRODUTO CODIGO', ''))
                ean_venda = tratar_codigo(linha.get(col_ean_vendas, '')) if col_ean_vendas else ''
                nome_p = str(linha.get('PRODUTO NOME', ''))
                
                if sku in CODIGOS_OURO: nome_p = "⭐ " + nome_p

                valor = mapa_cod_ativo.get(sku)
                if valor is None and ean_venda: valor = mapa_ean_ativo.get(ean_venda)

                if valor is not None:
                    try:
                        if float(valor) <= 0: continue 
                        else: txt_preco = f"R$ {float(valor):.2f}"
                    except: txt_preco = str(valor)
                else: continue 

                dados_tela.append({
                    "CÓDIGO": sku, "PRODUTO": nome_p, "SUGERIDO": txt_preco, "PREÇO NA LOJA": None 
                })

            df_tela = pd.DataFrame(dados_tela)

            if not df_tela.empty:
                df_styled = df_tela.style.apply(pintar_linha_ouro, axis=1)
                
                df_editado = st.data_editor(
                    df_styled, use_container_width=True, hide_index=True, disabled=["CÓDIGO", "PRODUTO", "SUGERIDO"],
                    column_config={"PREÇO NA LOJA": st.column_config.NumberColumn("PREÇO NA LOJA", format="R$ %.2f", min_value=0.0)}
                )
                
                st.divider()
                st.markdown("### 💬 Inteligência de Campo")
                feedback = st.text_area("Na sua opinião, o que poderíamos fazer para melhorar nossa participação neste cliente? (Opcional)", 
                                        placeholder="Digite aqui a sua sugestão ou observação de gôndola...",
                                        height=100)
                
                st.write("") 
                
                if st.button("🚀 ENVIAR AUDITORIA", type="primary", use_container_width=True):
                    with st.spinner("A gerar o PDF e a enviar o e-mail..."):
                        sucesso, mensagem = enviar_email_coleta(promotor, cliente_selecionado, cidade_loja, estado_selecionado, df_editado, feedback)
                    
                    if sucesso: 
                        st.success(mensagem)
                        st.balloons()
                    else: 
                        st.error(mensagem)
            else:
                st.warning("⚠️ Nenhum produto da tabela atual foi comprado por este cliente.")
else:
    st.error("❌ Arquivo 'Vendas.xlsx' não encontrado na pasta do sistema.")