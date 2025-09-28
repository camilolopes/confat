import io
import pandas as pd
import streamlit as st
from processor import build_processed_workbook

st.set_page_config(page_title="C6 Fatura Processor", page_icon="📊", layout="centered")
st.title("📊 Processador de Fatura C6")
st.caption("Envie o arquivo Excel (.xlsx) e receba a planilha consolidada pronta para análise.")

with st.expander("📌 Instruções rápidas", expanded=False):
    st.markdown(
        """**Como usar**
1. Clique em **Selecionar arquivo** e faça **upload do .xlsx** da fatura.
2. Clique em **Processar**.
3. Baixe o arquivo final gerado pelo app.

**O que o app faz**
- Consolida gastos por **cartão**, **estabelecimento** e **categoria por cartão** (apenas valores positivos).
- Gera **abas por cartão** com gráficos **pizza (Top 3 + Outras)** e título com **final do cartão + nome do portador**.
- Cria **Índice** com links para navegação e **oculta** a aba **Transações Originais**.
- Inclui **Devoluções** com valores negativos e **Resumo da Fatura**.
"""
    )

uploaded = st.file_uploader("Selecione o arquivo .xlsx", type=["xlsx"])

if uploaded is not None:
    st.write("Arquivo recebido:", uploaded.name)
    if st.button("▶️ Processar", type="primary"):
        try:
            output_bytes = build_processed_workbook(uploaded.read())
            st.success("Processamento concluído!")
            st.download_button(
                label="⬇️ Baixar planilha processada",
                data=output_bytes,
                file_name="fatura_c6_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
