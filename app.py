import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor de Pedidos - Campo Doce",
    page_icon="📊",
    layout="wide",
)

st.title("📄 Conversor de Pedidos PDF -> Excel (Com Base Campo Doce)")
st.markdown(
    "1. Faça o upload do **PDF do Pedido (Consinco)**.\n2. Faça o upload da"
    " sua **Base - Campo Doce (.xlsx ou .csv)** para cruzar com a coluna"
    " **CÓDIGO CD**."
)

# 1. Carregar o PDF do Pedido
uploaded_pdf = st.file_uploader(
    "Selecione o ficheiro PDF do pedido", type=["pdf"]
)

# 2. Carregar a Base de Dados Auxiliar
uploaded_aux = st.file_uploader(
    "Selecione a base de códigos (Base - Campo Doce)", type=["xlsx", "csv"]
)

if uploaded_pdf is not None:
  with st.spinner("A processar o PDF..."):
    texto_completo = ""
    with pdfplumber.open(uploaded_pdf) as pdf:
      for pagina in pdf.pages:
        txt = pagina.extract_text()
        if txt:
          texto_completo += txt + "\n"

    linhas = texto_completo.split("\n")
    dados_tabela = []

    for linha in linhas:
      match = re.search(
          r"^(\d{4,8})\s*(.*?)\s+(?:CX|UN|PC|KG|FD)\s*\d*\s+(\d+[\.,]\d{2})",
          linha.strip(),
      )
      if match:
        codigo = match.group(1)
        descricao = match.group(2)
        quantidade = match.group(3)
        dados_tabela.append({
            "Código (SEQ)": codigo,
            "Descrição do PDF": descricao,
            "Qtde": quantidade,
        })

    if dados_tabela:
      df_pedido = pd.DataFrame(dados_tabela)
      df_pedido = df_pedido.drop_duplicates()

      # Se a Base - Campo Doce foi carregada, efetuamos o PROCV pela coluna 'CÓDIGO CD'
      if uploaded_aux is not None:
        if uploaded_aux.name.endswith(".csv"):
          df_aux = pd.read_csv(uploaded_aux)
        else:
          df_aux = pd.read_excel(uploaded_aux)

        st.success("Base - Campo Doce carregada com sucesso!")
        st.subheader("Pré-visualização da Base Auxiliar:")
        st.dataframe(df_aux.head(3), use_container_width=True)

        try:
          df_pedido["Código (SEQ)"] = df_pedido["Código (SEQ)"].astype(str)

          # Procura especificamente pela coluna 'CÓDIGO CD' (ou variações) na base auxiliar
          col_aux_chave = None
          for col in df_aux.columns:
            col_limpa = str(col).strip().upper()
            if "CÓDIGO CD" in col_limpa or "CODIGO CD" in col_limpa:
              col_aux_chave = col
              break

          # Se não encontrar exatamente, procura por colunas que contenham "CD" ou "SEQ"
          if not col_aux_chave:
            for col in df_aux.columns:
              if any(
                  termo in str(col).upper() for termo in ["CD", "SEQ", "CÓDIGO"]
              ):
                col_aux_chave = col
                break

          # Fallback final para a 3ª coluna (onde o CÓDIGO CD costuma estar na imagem)
          if not col_aux_chave and len(df_aux.columns) >= 3:
            col_aux_chave = df_aux.columns[2]
          elif not col_aux_chave:
            col_aux_chave = df_aux.columns[0]

          df_aux[col_aux_chave] = df_aux[col_aux_chave].astype(str)

          # Realiza o PROCV (Left Merge) utilizando o Código do PDF e a coluna CÓDIGO CD da base
          df_final = pd.merge(
              df_pedido,
              df_aux,
              left_on="Código (SEQ)",
              right_on=col_aux_chave,
              how="left",
          )
          st.success(
              f"Cruzamento efetuado com sucesso usando a coluna '{col_aux_chave}'"
              " da Base Campo Doce!"
          )
        except Exception as e:
          df_final = df_pedido
          st.warning(
              f"Erro ao cruzar com a base auxiliar: {e}. A mostrar apenas os"
              " dados do PDF."
          )
      else:
        df_final = df_pedido
        st.info(
            "💡 Dica: Carregue o ficheiro da **Base - Campo Doce** para cruzar"
            " com os códigos CD automaticamente."
        )

      st.subheader("Resultado Final:")
      st.dataframe(df_final, use_container_width=True)

      # Botão para descarregar o Excel final
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, header=True)
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Descarregar Planilha Completa (.xlsx)",
          data=excel_data,
          file_name="pedido_com_codigo_cd.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.error("Não foram encontrados itens válidos no PDF.")
