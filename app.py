import streamlit as st
import pandas as pd
import io
import re


def tratar_planilha(file, incoterm_valor):
    # Carrega a planilha de origem mantendo a primeira linha como cabeçalho (header=0)
    df_origem = pd.read_excel(file, header=0)

    # Remove espaços extras no início/fim dos nomes das colunas da origem para evitar erros de digitação
    df_origem.columns = [str(col).strip() for col in df_origem.columns]

    # --- DICIONÁRIO DE MAPEAMENTO (Cabeçalhos da Planilha de Origem) ---
    C_PARTNUMBER = "CÓDIGO PRINCIPAL"
    C_QUANTIDADE = "QUANTIDADE"
    C_DESCRICAO = "DESCRICAO PORTUGUES"
    C_PRECO_UNIT = "VALOR UNITARIO ITEM"
    C_PESO_UNIT = "PESO LIQUIDO UNITÁRIO"
    C_FATURA = "FATURA"
    C_GTIN_EAN = "CODIGO(GTIN / EAN)"
    C_ORDEM_COMPRA = "ORDEM DE COMPRA"

    # Lista de colunas obrigatórias para verificar se a planilha de origem está correta
    colunas_obrigatorias = [
        C_PARTNUMBER, C_QUANTIDADE, C_DESCRICAO, C_PRECO_UNIT,
        C_PESO_UNIT, C_FATURA, C_GTIN_EAN, C_ORDEM_COMPRA
    ]

    # Verifica se algum cabeçalho está faltando na planilha de origem antes de processar
    colunas_faltantes = [col for col in colunas_obrigatorias if col not in df_origem.columns]
    if colunas_faltantes:
        raise ValueError(
            f"Os seguintes cabeçalhos não foram encontrados na planilha de origem: {', '.join(colunas_faltantes)}")

    # Criando o DataFrame final com a estrutura exata solicitada (Case Sensitive)
    colunas_finais = [
        'PARTNUMBER', 'QUANTIDADE', 'UNIDADE', 'PRECOTOTAL', 'PESOTOTAL',
        'INCOTERMS', 'MOEDA', 'FATURA', 'OUTRAS REFERENCIAS', 'nrlote', 'expedicao'
    ]
    df_final = pd.DataFrame(columns=colunas_finais)

    # A: PARTNUMBER ➔ "CÓDIGO PRINCIPAL"
    df_final['PARTNUMBER'] = df_origem[C_PARTNUMBER]

    # B: QUANTIDADE ➔ "QUANTIDADE"
    df_final['QUANTIDADE'] = df_origem[C_QUANTIDADE]

    # C: UNIDADE ➔ Lógica Tênis/Sapato/Mocassim baseada em "DESCRICAO PORTUGUES"
    def verificar_unidade(valor):
        valor_str = str(valor).upper()
        palavras_pares = ["TENIS", "TÊNIS", "SAPATO", "MOCASSIM", "SANDALIA"]
        if any(p in valor_str for p in palavras_pares):
            return "PARES"
        return "PECA"

    df_final['UNIDADE'] = df_origem[C_DESCRICAO].apply(verificar_unidade)

    # Converte colunas numéricas de forma segura para evitar erros matemáticos
    qtd_num = pd.to_numeric(df_origem[C_QUANTIDADE], errors='coerce').fillna(0)
    preco_num = pd.to_numeric(df_origem[C_PRECO_UNIT], errors='coerce').fillna(0)
    peso_num = pd.to_numeric(df_origem[C_PESO_UNIT], errors='coerce').fillna(0)

    # D: PRECOTOTAL ➔ "VALOR UNITARIO ITEM" * "QUANTIDADE"
    df_final['PRECOTOTAL'] = preco_num * qtd_num

    # E: PESOTOTAL ➔ ("PESO LIQUIDO UNITÁRIO" * "QUANTIDADE") convertido de gramas para KG (/ 1000)
    df_final['PESOTOTAL'] = (peso_num * qtd_num) / 1000

    # F: INCOTERMS ➔ Valor da interface
    df_final['INCOTERMS'] = incoterm_valor

    # G: MOEDA ➔ Sempre '790'
    df_final['MOEDA'] = '790'

    # H: FATURA ➔ "FATURA"
    df_final['FATURA'] = df_origem[C_FATURA]

    # H: FATURA ➔ "FATURA"
    df_final['nrlote'] = df_origem[C_ORDEM_COMPRA]

    # H: FATURA ➔ "FATURA"
    df_final['expedicao'] = df_origem[C_GTIN_EAN]

    # I: OUTRAS REFERENCIAS ➔ "CODIGO(GTIN / EAN)" (Tratado como Texto Puro)
    def limpar_referencia(valor):
        if pd.isna(valor): return ""
        val_str = str(valor).strip()
        if val_str.endswith('.0'): val_str = val_str[:-2]
        return val_str

    df_final['OUTRAS REFERENCIAS'] = df_origem[C_GTIN_EAN].apply(limpar_referencia)

    return df_final


# --- CONFIGURAÇÃO DA INTERFACE WEB (STREAMLIT) ---
st.set_page_config(page_title="Tratador de Planilhas", layout="centered")
st.title("📂 Tratamento planilha CHNL")
st.markdown("Insira os dados abaixo para gerar a nova planilha tratada.")

incoterm_input = st.text_input("Informe o INCOTERMS:", placeholder="Ex: FOB, CIF, EXW...")
uploaded_file = st.file_uploader("Selecione o arquivo Excel de origem (.xlsx)", type=["xlsx"])

if uploaded_file and incoterm_input:
    if st.button("Processar Planilha", use_container_width=True):
        try:
            with st.spinner("Processando dados pelos cabeçalhos..."):
                resultado = tratar_planilha(uploaded_file, incoterm_input)
                output = io.BytesIO()

                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    resultado.to_excel(writer, index=False, sheet_name='Planilha Tratada')
                    workbook = writer.book
                    worksheet = writer.sheets['Planilha Tratada']

                    # Força a formatação de texto puro do Excel nas colunas I, J e K (índices 8, 9, 10)
                    fmt_txt = workbook.add_format({'num_format': '@'})
                    worksheet.set_column(8, 10, None, fmt_txt)

                st.success("Planilha processada com sucesso!")
                st.download_button(
                    label="📥 Clique aqui para baixar a Planilha Tratada",
                    data=output.getvalue(),
                    file_name="planilha_tratada.xlsx",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")

elif uploaded_file and not incoterm_input:
    st.warning("⚠️ Por favor, digite o INCOTERMS no campo de texto para liberar o processamento.")
