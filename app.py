import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA (STREAMLIT)
st.set_page_config(page_title="Validador Ultra Rápido - Pesagem", page_icon="⚖️", layout="centered")

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

st.title("⚖️ Controle de Validades Flash - Setor de Pesagem")
st.write("Processamento de alta performance para liberação diária de OPs.")

# 2. ÁREA DE IMPORTAÇÃO
st.subheader("📊 1. Carregar Base de Dados do SAP")
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione o arquivo Excel exportado do SAP (.xlsx):", 
    type=["xlsx"]
)

@st.cache_data(ttl=3600)  # Otimização Streamlit: Mantém em memória se o mesmo arquivo for usado na mesma sessão
def processar_sap_alta_performance(arquivo):
    # 1. Carga única e rápida: lê as primeiras 25 linhas apenas para achar o cabeçalho
    df_header_check = pd.read_excel(arquivo, nrows=25, header=None, engine='openpyxl')
    
    linha_titulos = 8  # Fallback padrão (linha 9 do Excel)
    for idx, row in df_header_check.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notna(v)]
        if 'MATERIAL' in valores:
            linha_titulos = idx
            break
            
    # 2. Carrega a planilha real pulando o topo de forma ultra rápida
    arquivo.seek(0)
    df = pd.read_excel(arquivo, skiprows=linha_titulos, dtype=str, engine='openpyxl')
    
    # Remove espaços vazios dos nomes das colunas de forma vetorizada
    df.columns = df.columns.str.strip()
    
    c_cod = 'Material'
    c_nom = 'Texto breve material'
    c_lot = 'Lote'
    c_val = 'Data venc.'
    
    # Validação expressa de colunas estruturais
    if c_cod not in df.columns or c_val not in df.columns:
        st.error(f"❌ Layout incompatível detectado na linha {linha_titulos + 1}.")
        st.stop()
        
    # 3. Limpeza Expressa de Nulos e Espaços
    df = df.dropna(subset=[c_cod, c_val, c_lot])
    df[c_cod] = df[c_cod].str.strip()
    df[c_lot] = df[c_lot].str.strip()
    
    # 4. FILTRAGEM VETORIZADA: Drop rápido de lotes que começam com 'ME' (Materiais de Embalagem)
    # O operador ~ inverte a máscara booleana feita diretamente em C
    df = df[~df[c_lot].str.upper().str.startswith('ME', na=False)]
    
    # Cria os códigos limpos de forma vetorizada (muito mais rápido que loops)
    df['cod_limpo'] = df[c_cod].str.lstrip('0')
    
    # 5. CONVERSÃO VETORIZADA DE DATA: Remove resíduos de hora e converte o bloco inteiro de uma vez
    df[c_val] = df[c_val].str.split().str[0]
    df['Data_Processada'] = pd.to_datetime(df[c_val], dayfirst=True, errors='coerce')
    df = df.dropna(subset=['Data_Processada'])
    
    # Janela de corte estipulada para 180 dias
    limite_data = datetime.now() + timedelta(days=180)
    criticos = df[df['Data_Processada'] <= limite_data]
    
    # 6. Mapeamento indexado em dicionário para busca instantânea O(1)
    resultado = {}
    for _, linha in criticos.iterrows():
        cod_orig = linha[c_cod]
        cod_limp = linha['cod_limpo']
        nome = str(linha[c_nom]).strip()
        lote = linha[c_lot]
        validade_str = linha['Data_Processada'].strftime('%d/%m/%Y')
        
        dados_lote = {"nome": nome, "lote": lote, "validade": validade_str}
        
        # Aloca memória indexada para os formatos com e sem zero
        if cod_orig not in resultado: resultado[cod_orig] = []
        resultado[cod_orig].append(dados_lote)
        
        if cod_limp not in resultado: resultado[cod_limp] = []
        resultado[cod_limp].append(dados_lote)
        
    return resultado, criticos, c_cod, c_nom, c_lot, 'Data_Processada'

if arquivo_carregado is not None:
    try:
        # Executa o processamento otimizado
        mapa_critico, df_criticos_puro, c_cod, c_nom, c_lot, c_val = processar_sap_alta_performance(arquivo_carregado)
        st.sidebar.success("⚡ Processamento concluído em milissegundos!")
        
        # PAINEL LATERAL DE MATÉRIAS-PRIMAS CRÍTICAS
        st.sidebar.markdown(f"### 🚨 Lotes Críticos de MP ({len(df_criticos_puro)})")
        if not df_criticos_puro.empty:
            df_criticos_puro = df_criticos_puro.sort_values(by=c_val)
            for _, reg in df_criticos_puro.iterrows():
                validade_formatada = reg[c_val].strftime('%d/%m/%Y')
                st.sidebar.markdown(
                    f"**Cód:** {reg[c_cod]} | **Lote:** {reg[c_lot]}\n"
                    f"*{reg[c_nom]}*\n"
                    f"Vence em: **{validade_formatada}**\n"
                    "---"
                )
        else:
            st.sidebar.info("Nenhuma MP crítica para os próximos 6 meses.")
            
        # 3. CAMPO DE ENTRADA DO OPERADOR
        st.subheader("🔍 2. Consulta de Insumos da OP")
        entrada_usuario = st.text_area(
            "Digite ou cole os códigos das matérias-primas (separe por vírgula):",
            placeholder="Exemplo: 506847, 507074"
        )

        if st.button("Verificar Validades", use_container_width=True):
            if entrada_usuario:
                # Separação e limpeza instantânea de strings
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
            else:
                st.warning("Por favor, insira pelo menos um código de material para verificar.")

    except Exception as e:
        st.error("❌ Erro no processamento de alta performance.")
        st.code(f"Detalhe técnico: {str(e)}")
else:
    st.info("💡 Carregue a planilha para ativar o validador rápido.")
