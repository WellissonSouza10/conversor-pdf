import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor de Pedidos para Excel", page_icon="📊", layout="wide"
)

st.title("📄 Conversor de Pedidos PDF -> Excel (Focado em Código e Qtde)")
st.markdown(
    "Faça o upload do PDF. O sistema vai extrair o **Código (SEQ)** e a"
    " **Quantidade** com base no layout do Consinco."
)

uploaded_file = st.file_uploader(
    "Selecione o ficheiro PDF do pedido", type=["pdf"]
)

if uploaded_file is not None:
  with st.spinner("A ler o texto e a extrair os produtos..."):
    texto_completo = ""

    # Lê o texto de todas as páginas do PDF
    with pdfplumber.open(uploaded_file) as pdf:
      for pagina in pdf.pages:
        txt = pagina.extract_text()
        if txt:
          texto_completo += txt + "\n"

    # Quebra o texto por linhas para processamento individual
    linhas = texto_completo.split("\n")
    dados_tabela = []

    # Expressão regular para detetar linhas de produtos do Consinco:
    # Começa com um código de 4 a 8 dígitos, seguido do nome e da quantidade no formato X,XX ou XX,XX
    for linha in linhas:
      # Exemplo de linha do PDF: "98967MASSA TALHARIM MEZZANI 500G CX 12 3,00 137,8100 413,43..."
      # Procuramos o código no início e a quantidade logo após as unidades de embalagem (CX, UN, PC)
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
            "Descrição do Produto": descricao,
            "Qtde": quantidade,
        })
      else:
        # Padrão alternativo caso o espaçamento varie ligeiramente
        match_alt = re.search(r"^(\d{4,8})\s+(.*)", linha.strip())
        if match_alt:
          # Tenta isolar números que pareçam quantidades na mesma linha
          partes = linha.split()
          if len(partes) >= 4:
            # O código é o primeiro elemento e a quantidade costuma estar antes dos valores monetários
            codigo = partes[0]
            if len(codigo) >= 4 and codigo.isdigit():
              # Procura um valor decimal curto na linha que sirva de quantidade
              for p in partes:
                if (
                    "," in p
                    and len(p) <= 6
                    and p.replace(",", "").isdigit()
                    and p != codigo
                ):
                  quantidade = p
                  # Descrição fica no meio
                  desc = " ".join(
                      [
                          x
                          for x in partes[1:]
                          if x != p
                          and not "137" in x
                          and not "247" in x
                          and not "0,00" in x
                      ]
                  )
                  dados_tabela.append({
                      "Código (SEQ)": codigo,
                      "Descrição do Produto": desc[:50],
                      "Qtde": quantidade,
                  })
                  break

    # Se a extração automática por linhas avançadas trouxer dados, geramos o DataFrame
    if dados_tabela:
      df = pd.DataFrame(dados_tabela)
      # Remove duplicados caso o regex apanhe a mesma linha duas vezes
      df = df.drop_duplicates()

      st.success(
          f"Foram extraídos {len(df)} itens com sucesso do documento!"
      )
      st.dataframe(df, use_container_width=True)

      # Botão para descarregar o Excel limpo
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=True)
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Descarregar Planilha Filtrada em Excel (.xlsx)",
          data=excel_data,
          file_name="pedido_seq_qtde.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      # Plano de contingência: mostra o texto bruto para conferência se o formato exato mudar
      st.warning(
          "Não foi possível aplicar o filtro automático exate para este"
          " layout. A apresentar o texto extraído:"
      )
      st.text_area("Texto Bruto do PDF", texto_completo, height=300)