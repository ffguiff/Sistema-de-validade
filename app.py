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
st.write("Carregue a planilha do SAP e insira os códigos das matérias-primas para validação (6 meses).")

# 2. ÁREA DE IMPORTAÇÃO
st.subheader("📊 1. Carregar Base de Dados do SAP")
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione o arquivo Excel exportado do SAP (.xlsx):", 
    type=["xlsx"]
)

def processar_sap_bruto(arquivo):
    # Carrega o arquivo sem cabeçalho para mapear manualmente as linhas
    df_raw = pd.read_excel(arquivo, header=None, dtype=str, engine='openpyxl')
    
    linha_titulos = None
    for idx, row in df_raw.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notna(v)]
        if 'MATERIAL' in valores:
            linha_titulos = idx
            break
            
    if linha_titulos is None:
        st.error("❌ Não foi possível encontrar a coluna 'Material' nas primeiras linhas do arquivo.")
        st.stop()
        
    # Reconstrói o DataFrame usando a linha correta como cabeçalho
    df_dados = df_raw.iloc[linha_titulos+1:].copy()
    df_dados.columns = [str(c).strip() for c in df_raw.iloc[linha_titulos].values]
    
    c_cod = 'Material'
    c_nom = 'Texto breve material'
    c_lot = 'Lote'
    c_val = 'Data venc.'
    
    # Remove linhas onde as colunas essenciais estão nulas
    df_dados = df_dados.dropna(subset=[c_cod, c_val])
    
    # Padroniza textos, códigos e lotes
    df_dados[c_cod] = df_dados[c_cod].astype(str).str.strip()
    df_dados[c_nom] = df_dados[c_nom].astype(str).str.strip()
    df_dados[c_lot] = df_dados[c_lot].astype(str).str.strip()
    
    # MELHORIA CRUCIAL: Drop de Materiais de Embalagem com base na coluna LOTE
    # Remove qualquer linha onde o Lote comece com "ME" (ignora maiúsculas/minúsculas)
    df_dados = df_dados[~df_dados[c_lot].str.upper().str.startswith('ME', na=False)]
    
    # Cria a coluna de código limpo (sem zeros à esquerda) para a busca flexível
    df_dados['cod_limpo'] = df_dados[c_cod].str.lstrip('0')
    
    # Conversão robusta de datas
    datas_convertidas = []
    for val in df_dados[c_val]:
        val_str = str(val).strip().split()[0]
        try:
            dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        except:
            dt = pd.NaT
        datas_convertidas.append(dt)
        
    df_dados['Data_Processada'] = datas_convertidas
    df_dados = df_dados.dropna(subset=['Data_Processada'])
    
    # Define a janela de corte de 6 meses (180 dias)
    limite_data = datetime.now() + timedelta(days=180)
    criticos = df_dados[df_dados['Data_Processada'] <= limite_data]
    
    resultado = {}
    for _, linha in criticos.iterrows():
        cod_orig = str(linha[c_cod]).strip()
        cod_limp = str(linha['cod_limpo']).strip()
        nome = str(linha[c_nom]).strip()
        lote = str(linha[c_lot]).strip()
        validade_str = linha['Data_Processada'].strftime('%d/%m/%Y')
        
        dados_lote = {"nome": nome, "lote": lote, "validade": validade_str}
        
        # Registra o lote crítico nas duas opções de busca (com e sem zero à esquerda)
        if cod_orig not in resultado: resultado[cod_orig] = []
        resultado[cod_orig].append(dados_lote)
        
        if cod_limp not in resultado: resultado[cod_limp] = []
        resultado[cod_limp].append(dados_lote)
        
    return resultado, criticos, c_cod, c_nom, c_lot, 'Data_Processada'

if arquivo_carregado is not None:
    try:
        mapa_critico, df_criticos_puro, c_cod, c_nom, c_lot, c_val = processar_sap_bruto(arquivo_carregado)
        st.sidebar.success("📊 Planilha processada!")
        
        # EXIBIR PAINEL DE LOTES CRÍTICOS NA BARRA LATERAL (APENAS MPs)
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
            st.sidebar.info("Nenhuma matéria-prima crítica encontrada para os próximos 6 meses.")
            
        # 3. CAMPO DE ENTRADA DO OPERADOR
        st.subheader("🔍 2. Consulta de Insumos da OP")
        entrada_usuario = st.text_area(
            "Digite ou cole os códigos das matérias-primas (separe por vírgula):",
            placeholder="Exemplo: 506847, 507074"
        )

        if st.button("Verificar Validades", use_container_width=True):
            if entrada_usuario:
                codigos_verificar = [c.strip() for c in entrada_usuario.split(",") if c.strip()]
                
                # Monta uma lista limpa retirando os zeros à esquerda para cruzar os dados
                codigos_verificar_limpos = [c.lstrip('0') for c in codigos_verificar]
                todos_codigos_busca = list(set(codigos_verificar + codigos_verificar_limpos))
                
                st.write("### Resultado da Análise da OP:")
                algum_critico_encontrado = False
                codigos_exibidos = set() # Evita duplicar o alerta caso o código bata duas vezes
                
                for codigo in todos_codigos_busca:
                    if codigo in mapa_critico:
                        for info_lote in mapa_critico[codigo]:
                            # Identificador único do alerta para não repetir
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
        st.error("❌ Erro ao processar o arquivo Excel.")
        st.code(f"Detalhe técnico: {str(e)}")
else:
    st.info("💡 Por favor, carregue o arquivo Excel exportado do SAP para ativar o validador.")
