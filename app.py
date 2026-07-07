import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Validador de Insumos - Pesagem", page_icon="⚖️", layout="centered")

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
st.write("Carregue a planilha do SAP e insira os códigos das matérias-primas para checar lotes críticos de 6 meses (180 dias).")

# 2. ÁREA DE IMPORTAÇÃO DA BASE DE DADOS (DINÂMICA)
st.subheader("📊 1. Carregar Base de Dados do SAP")
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione o arquivo Excel exportado do SAP (.xlsx):", 
    type=["xlsx"]
)

def processar_sap(arquivo):
    # Lê o arquivo tratando a coluna Material como texto
    df = pd.read_excel(arquivo, dtype={'Material': str}, engine='openpyxl')
    
    # Limpa espaços em branco dos nomes das colunas
    df.columns = [str(c).strip() for c in df.columns]
    
    # Validação das colunas obrigatórias
    colunas_obrigatorias = ['Material', 'Texto breve material', 'Lote', 'Validade', 'Tipo de material']
    for col in colunas_obrigatorias:
        if col not in df.columns:
            df[col] = "N/A" if col != 'Validade' else pd.NaT
            
    # FILTRO: Mantém estritamente o tipo 'ROH' (Matéria-Prima)
    df['Tipo de material'] = df['Tipo de material'].astype(str).str.strip().str.upper()
    df = df[df['Tipo de material'] == 'ROH']
            
    # CORREÇÃO DA DATA: Força o Pandas a entender que o DIA vem primeiro (formato BR)
    df['Validade'] = pd.to_datetime(df['Validade'], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Validade'])
    
    # MODIFICAÇÃO: Altera a linha de corte para 180 dias (6 meses)
    limite_data = datetime.now() + timedelta(days=180)
    criticos = df[df['Validade'] <= limite_data]
    
    resultado = {}
    for _, linha in criticos.iterrows():
        cod = str(linha['Material']).strip()
        nome = str(linha['Texto breve material']).strip()
        lote = str(linha['Lote']).strip()
        validade_str = linha['Validade'].strftime('%d/%m/%Y')
        
        if cod not in resultado:
            resultado[cod] = []
        resultado[cod].append({"nome": nome, "lote": lote, "validade": validade_str})
        
    return resultado, criticos

# Só libera o resto da tela se o operador tiver carregado um arquivo
if arquivo_carregado is not None:
    try:
        mapa_critico, df_criticos_puro = processar_sap(arquivo_carregado)
        st.sidebar.success("📊 MPs (ROH) processadas com sucesso!")
        
        # EXIBIR PAINEL DE LOTES CRÍTICOS NA BARRA LATERAL (AGORA PARA 6 MESES)
        st.sidebar.markdown(f"### 🚨 Lotes Críticos de MP ({len(df_criticos_puro)})")
        if not df_criticos_puro.empty:
            # Ordena do mais próximo do vencimento para o mais distante
            df_criticos_puro = df_criticos_puro.sort_values(by='Validade')
            for _, reg in df_criticos_puro.iterrows():
                validade_formatada = reg['Validade'].strftime('%d/%m/%Y')
                st.sidebar.markdown(
                    f"**Cód:** {reg['Material']} | **Lote:** {reg['Lote']}\n"
                    f"*{reg['Texto breve material']}*\n"
                    f"Vence em: **{validade_formatada}**\n"
                    "---"
                )
        else:
            st.sidebar.info("Nenhuma matéria-prima crítica para os próximos 6 meses.")
            
        # 3. CAMPO DE ENTRADA DO OPERADOR
        st.subheader("🔍 2. Consulta de Insumos da OP")
        entrada_usuario = st.text_area(
            "Digite ou cole os códigos das matérias-primas (separe por vírgula):",
            placeholder="Exemplo: 100001, 100002"
        )

        if st.button("Verificar Validades", use_container_width=True):
            if entrada_usuario:
                codigos_verificar = [c.strip() for c in entrada_usuario.split(",") if c.strip()]
                
                st.write("### Resultado da Análise da OP:")
                algum_critico_encontrado = False
                
                for codigo in codigos_verificar:
                    if codigo in mapa_critico:
                        algum_critico_encontrado = True
                        for info_lote in mapa_critico[codigo]:
                            st.markdown(f"""
                                <div class="alerta-critico">
                                    ⚠️ <b>MATÉRIA-PRIMA CRÍTICA!</b><br>
                                    <b>Nome:</b> {info_lote['nome']}<br>
                                    <b>Código:</b> {codigo}<br>
                                    <b>Lote:</b> {info_lote['lote']}<br>
                                    <b>Data de Vencimento:</b> {info_lote['validade']}
                                </div>
                            """, unsafe_allow_html=True)
                
                if not algum_critico_encontrado:
                    st.markdown("""
                        <div class="sucesso-liberado">
                            ✅ <b>OP LIBERADA:</b> Nenhuma das matérias-primas inseridas possui lotes críticos.
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.warning("Por favor, insira pelo menos um código de material para verificar.")

    except Exception as e:
        st.error("❌ Erro ao ler a estrutura desse arquivo Excel.")
        st.code(f"Detalhe técnico: {str(e)}")
else:
    st.info("💡 Por favor, carregue o arquivo Excel exportado do SAP para ativar o validador.")
