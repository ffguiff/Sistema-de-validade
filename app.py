import os
import sys

# Auto-instalador do openpyxl se ele sumir do ambiente do Streamlit
try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Validador de Insumos - Pesagem", page_icon="⚖️", layout="centered")

# Estilização visual para os alertas da produção
st.markdown("""
    <style>
    .alerta-critico {
        background-color: #ffcccc;
        padding: 15px;
        border-radius: 5px;
        border: 2px solid #ff0000;
        color: #990000;
        font-size: 15px;
        margin-bottom: 10px;
    }
    .sucesso-liberado {
        background-color: #d4edda;
        padding: 12px;
        border-radius: 5px;
        border: 2px solid #28a745;
        color: #155724;
        font-size: 15px;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ Controle de Validades - Setor de Pesagem")
st.write("Insira os códigos dos materiais da OP (separados por vírgula) para checar lotes críticos de 3 meses.")

# 2. CARREGAMENTO DA BASE DO SAP (LX02)
@st.cache_data
def carregar_sap():
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    caminho_excel = os.path.join(pasta_atual, "lx02_export.xlsx")
    
    # Lê forçando os códigos a virem como texto
    df = pd.read_excel(caminho_excel, dtype={'Codigo_Pr': str})
    
    # Limpa linhas vazias na data e converte o formato de pontos do SAP
    df = df.dropna(subset=['Data_Validade'])
    df['Data_Validade'] = pd.to_datetime(df['Data_Validade'], format='%d.%m.%Y', errors='coerce')
    df = df.dropna(subset=['Data_Validade'])
    
    # Define limite de 90 dias (3 meses a partir de hoje)
    limite_data = datetime.now() + timedelta(days=180)
    criticos = df[df['Data_Validade'] <= limite_data]
    
    resultado = {}
    for _, linha in criticos.iterrows():
        cod = str(linha['Codigo_Pr']).strip()
        nome = str(linha['Texto_bre']).strip() if 'Texto_bre' in df.columns else "Não Informado"
        lote = str(linha['Lote']).strip() if 'Lote' in df.columns else "N/A"
        validade_str = linha['Data_Validade'].strftime('%d/%m/%Y')
        
        if cod not in resultado:
            resultado[cod] = []
        resultado[cod].append({"nome": nome, "lote": lote, "validade": validade_str})
        
    return resultado, criticos

try:
    mapa_critico, df_criticos_puro = carregar_sap()
    st.sidebar.success("📊 Base SAP LX02 carregada!")
    
    # EXIBIR PAINEL DE LOTES CRÍTICOS NA BARRA LATERAL
    st.sidebar.markdown("### 🚨 Lotes Críticos no Estoque")
    if not df_criticos_puro.empty:
        # Formata o visual da tabela na lateral para o operador consultar de relance
        for _, reg in df_criticos_puro.iterrows():
            validade_formatada = reg['Data_Validade'].strftime('%d/%m/%Y')
            st.sidebar.markdown(
                f"**Cód:** {reg['Codigo_Pr']} | **Lote:** {reg['Lote']}\n"
                f"*{reg['Texto_bre']}*\n"
                f"Vence em: **{validade_formatada}**\n"
                "---"
            )
    else:
        st.sidebar.info("Nenhum lote crítico encontrado para os próximos 3 meses.")
        
except Exception as e:
    st.error("❌ ERRO ao processar as colunas do Excel.")
    st.info("Verifique se as colunas estão como: 'Codigo_Pr', 'Texto_bre', 'Lote' and 'Data_Validade'.")
    st.code(f"Detalhe técnico: {str(e)}")
    st.stop()

# 3. CAMPO DE ENTRADA DO OPERADOR
st.subheader("Consulta de Insumos da OP")
entrada_usuario = st.text_area(
    "Digite ou cole os códigos dos materiais (separe por vírgula):",
    placeholder="Exemplo: 100001, 100002"
)

if st.button("🔍 Verificar Validades", use_container_width=True):
    if entrada_usuario:
        codigos_verificar = [c.strip() for c in entrada_usuario.split(",") if c.strip()]
        
        st.write("### Resultado da Análise da OP:")
        
        algum_critico_encontrado = False
        
        for codigo in codigos_verificar:
            # SÓ MOSTRA SE ESTIVER CRÍTICO
            if codigo in mapa_critico:
                algum_critico_encontrado = True
                for info_lote in mapa_critico[codigo]:
                    st.markdown(f"""
                        <div class="alerta-critico">
                            ⚠️ <b>PRODUTO CRÍTICO DETECTADO!</b><br>
                            <b>Nome:</b> {info_lote['nome']}<br>
                            <b>Código:</b> {codigo}<br>
                            <b>Lote:</b> {info_lote['lote']}<br>
                            <b>Data de Vencimento:</b> {info_lote['validade']}
                        </div>
                    """, unsafe_allow_html=True)
        
        # Se nenhum dos códigos digitados tinha lote crítico, solta a mensagem verde tranquila
        if not algum_critico_encontrado:
            st.markdown("""
                <div class="sucesso-liberado">
                    ✅ <b>OP LIBERADA:</b> Nenhum dos materiais inseridos possui lotes críticos próximos ao vencimento.
                </div>
            """, unsafe_allow_html=True)
            
    else:
        st.warning("Por favor, insira pelo menos um código de material para verificar.")