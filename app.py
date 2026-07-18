import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Configurações globais da interface do Streamlit
st.set_page_config(page_title="Validador & Dashboard - Pesagem", page_icon="⚖️", layout="centered")

# Injeção de CSS personalizado para estilização profissional dos alertas e cartões
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
    .card-metrica {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 5px;
        border-left: 5px solid #28a745;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True) 

st.title("⚖️ Controle de Validades & Dashboard Gerencial")
st.write("Sistema de alta performance focado estritamente no escoamento de Matérias-Primas (MP).")

PASTA_DATA = "data"
CAMINHO_LOCAL = os.path.join(PASTA_DATA, "sap_atual.xlsx")

# Cria a pasta de dados local se não existir no servidor para a persistência
if not os.path.exists(PASTA_DATA):
    os.makedirs(PASTA_DATA)

st.subheader("📊 1. Carregar Base de Dados do SAP")
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione o arquivo Excel exportado do SAP (.xlsx):", 
    type=["xlsx"]
)

planilha_para_processar = None

# Lógica inteligente para blindar o sistema contra re-uploads constantes (Timeouts)
if arquivo_carregado is not None:
    with open(CAMINHO_LOCAL, "wb") as f:
        f.write(arquivo_carregado.getbuffer())
    planilha_para_processar = CAMINHO_LOCAL
    st.success("💾 Nova base de dados carregada e salva com sucesso no servidor!")
elif os.path.exists(CAMINHO_LOCAL):
    planilha_para_processar = CAMINHO_LOCAL
    st.info("⚡ Utilizando a última base de dados salva no servidor (Auto-recuperação ativa).")

@st.cache_data(ttl=3600)
def processar_sap_alta_performance(caminho_arquivo):
    df_header_check = pd.read_excel(caminho_arquivo, nrows=25, header=None, engine='openpyxl')
    
    linha_titulos = 8
    for idx, row in df_header_check.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notna(v)]
        if 'MATERIAL' in valores:
            linha_titulos = idx
            break
            
    df = pd.read_excel(caminho_arquivo, skiprows=linha_titulos, dtype=str, engine='openpyxl')
    df.columns = df.columns.str.strip()
    
    c_cod = 'Material'
    c_nom = 'Texto breve material'
    c_lot = 'Lote'
    c_val = 'Data venc.'
    
    if c_cod not in df.columns or c_val not in df.columns:
        st.error(f"❌ Layout incompatível detectado.")
        st.stop()

    df = df.dropna(subset=[c_cod, c_val, c_lot, c_nom])
    df[c_cod] = df[c_cod].str.strip()
    df[c_lot] = df[c_lot].str.strip()
    df[c_nom] = df[c_nom].str.strip()
    df['cod_limpo'] = df[c_cod].str.lstrip('0')
    
    # Filtro rígido e inclusivo para manter unicamente lotes que começam com 'MP'
    df = df[df[c_lot].str.upper().str.startswith('MP', na=False)]
    
    # Conversão vetorizada das datas para evitar lentidões
    df[c_val] = df[c_val].str.split().str[0]
    df['Data_Processada'] = pd.to_datetime(df[c_val], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Data_Processada'])
    
    # Filtro temporal: De ontem até 180 dias futuros
    data_inicio = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
    limite_data = data_inicio + timedelta(days=181)
    df_tempo = df[(df['Data_Processada'] >= data_inicio) & (df['Data_Processada'] <= limite_data)].copy()
    
    # Eliminação rigorosa de duplicidades decorrentes de múltiplas posições físicas de estoque (WM)
    criticos = df_tempo.drop_duplicates(subset=[c_cod, c_lot, 'Data_Processada'])
    
    resultado = {}
    for _, linha in criticos.iterrows():
        cod_orig = linha[c_cod]
        cod_limp = linha['cod_limpo']
        nome = str(linha[c_nom]).strip()
        lote = linha[c_lot]
        validade_str = linha['Data_Processada'].strftime('%d/%m/%Y')
        
        dados_lote = {"nome": nome, "lote": lote, "validade": validade_str}
        
        if cod_orig not in resultado: 
            resultado[cod_orig] = []
        resultado[cod_orig].append(dados_lote)
        
        if cod_limp not in resultado: 
            resultado[cod_limp] = []
        resultado[cod_limp].append(dados_lote)
        
    return resultado, criticos, c_cod, c_nom, c_lot, 'Data_Processada'

if planilha_para_processar is not None:
    try:
        mapa_critico, df_criticos_puro, c_cod, c_nom, c_lot, c_val = processar_sap_alta_performance(planilha_para_processar)
        st.sidebar.success("⚡ Processamento concluído!")
        
        st.sidebar.markdown(f"### 🚨 Lotes MP a Vencer ({len(df_criticos_puro)})")
        if not df_criticos_puro.empty:
            df_ordenado_lateral = df_criticos_puro.sort_values(by=c_val)
            for _, reg in df_ordenado_lateral.head(30).iterrows():
                validade_formatada = reg[c_val].strftime('%d/%m/%Y')
                st.sidebar.markdown(
                    f"**Cód:** {reg[c_cod]} | **Lote:** {reg[c_lot]}\n"
                    f"*{reg[c_nom]}*\n"
                    f"Vence em: **{validade_formatada}**\n"
                    "---"
                )
        else:
            st.sidebar.info("Nenhum lote de MP crítica para os próximos 6 meses.")
            
        aba_consulta, aba_dashboard = st.tabs(["🔍 Consulta de OPs", "📊 Dashboard de Perdas & Projeções"])
        
        with aba_consulta:
            st.subheader("Consulta de Insumos da Ordem de Processo")
            entrada_usuario = st.text_area(
                "Digite ou cole os códigos das matérias-primas da OP (separe por vírgula):",
                placeholder="Exemplo: 100389, 102108"
            )

            if st.button("Verificar Validades", use_container_width=True):
                if entrada_usuario:
                    codigos_verificar = [c.strip() for c in entrada_usuario.split(",") if c.strip()]
                    codigos_verificar_limpos = [c.lstrip('0') for c in codigos_verificar]
                    todos_codigos_busca = list(set(codigos_verificar + codigos_verificar_limpos))
                    
                    st.write("### Resultado da Análise da OP:")
                    algum_critico_encontrado = False
                    codigos_exibidos = set()
                    
                    for codigo in todos_codigos_busca:
                        if codigo in mapa_critico:
                            for info_lote in mapa_critico[codigo]:
                                chave_alerta = f"{codigo}_{info_lote['lote']}"
                                if chave_alerta not in codigos_exibidos:
                                    algum_critico_encontrado = True
                                    codigos_exibidos.add(chave_alerta)
                                    st.markdown(f"""
                                        <div class="alerta-critico">
                                            ⚠️ <b>MATÉRIA-PRIMA CRÍTICA DETECTADA!</b><br>
                                            <b>Nome:</b> {info_lote['nome']}<br>
                                            <b>Código:</b> {codigo}<br>
                                            <b>Lote:</b> {info_lote['lote']}<br>
                                            <b>Data de Vencimento:</b> {info_lote['validade']}
                                        </div>
                                    """, unsafe_allow_html=True)
                    
                    if not algum_critico_encontrado:
                        st.markdown("""
                            <div class="sucesso-liberado">
                                ✅ <b>OP LIBERADA:</b> Nenhuma das matérias-primas inseridas possui lotes críticos no estoque.
                            </div>
                        """, unsafe_allow_html=True)

        with aba_dashboard:
            st.subheader("Projeção Preditiva de Escoamento (Próximos 6 Meses)")
            
            if not df_criticos_puro.empty:
                # Ordena por data real antes de formatar para manter a integridade da série temporal
                df_criticos_puro = df_criticos_puro.sort_values(by=c_val)
                
                # Usamos o formato 'YYYY-MM' (Ex: 2026-07) para garantir que a ordenação alfabética do gráfico seja 100% cronológica
                df_criticos_puro['Mês Vencimento'] = df_criticos_puro[c_val].dt.strftime('%Y-%m')
                meses_ordenados = df_criticos_puro['Mês Vencimento'].unique()
                
                st.markdown(f"""
                    <div class="card-metrica">
                        🎯 <b>POTENCIAL DE EVITABILIDADE DE PERDAS (FUTURAS):</b><br>
                        Identificados <b>{len(df_criticos_puro)} lotes de MP únicos</b> com vencimento crítico mapeado para os próximos 6 meses.<br>
                        <i>Otimização: Linhas duplicadas por posições físicas de palete no estoque WM foram unificadas.</i>
                    </div>
                """, unsafe_allow_html=True)
                
                st.write("#### Distribuição de Lotes de MP por Mês de Vencimento:")
                grafico_data = df_criticos_puro.groupby('Mês Vencimento').size()
                grafico_data = grafico_data.reindex(meses_ordenados)
                st.bar_chart(grafico_data, color="#ff4b4b")
                
                st.write("#### 🔍 Detalhar Lotes por Mês de Vencimento")
                mes_selecionado = st.selectbox("Selecione um mês abaixo para listar as MPs:", options=meses_ordenados)
                
                if mes_selecionado:
                    df_filtrado_mes = df_criticos_puro[df_criticos_puro['Mês Vencimento'] == mes_selecionado].copy()
                    df_filtrado_mes['Vencimento'] = df_filtrado_mes[c_val].dt.strftime('%d/%m/%Y')
                    
                    tabela_exibicao = df_filtrado_mes[[c_cod, c_nom, c_lot, 'Vencimento']]
                    tabela_exibicao.columns = ['Código MP', 'Descrição do Insumo', 'Lote', 'Data Vencimento']
                    st.dataframe(tabela_exibicao, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum lote com a sigla 'MP' encontrado a vencer nos critérios de 6 meses futuros.")

    except Exception as e:
        st.error("❌ Erro no processamento do Dashboard.")
        st.code(f"Detalhe técnico: {str(e)}")
else:
    st.info("💡 Carregue a planilha para ativar o validador.")
