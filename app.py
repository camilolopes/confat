
import io
import pandas as pd
import streamlit as st
from processor import build_processed_workbook_c6, build_processed_workbook_nubank

st.set_page_config(page_title="Faturas Cartão - Processor", page_icon="💳", layout="centered")

st.title("💳 Processador de Faturas (C6 & Nubank)")
st.caption("Escolha o banco e envie a fatura no formato correto para gerar a planilha consolidada.")

with st.expander("📌 Como funciona", expanded=False):
    st.markdown(
        """**Fluxo**
1. Selecione o **Banco**.
2. Envie o arquivo no **formato exigido**:
   - **C6** -> .xlsx
   - **Nubank** -> .pdf
3. Clique em **Processar** e baixe o resultado em Excel.

**O que o app gera**
- Índice navegável
- Consolidados: **Cartão**, **Estabelecimento**, **Categoria por Cartão**
- **Devoluções** (valores negativos) e **Resumo da Fatura**
- Abas por cartão com **pizza (Top 3 + Outras)** e título com **final do cartão + portador**
- Aba **Transações Originais** (oculta)
"""
    )

bank = st.selectbox("Banco", ["C6 (Excel .xlsx)", "Nubank (PDF .pdf)"])

if bank.startswith("C6"):
    uploaded = st.file_uploader("Envie o arquivo .xlsx do C6", type=["xlsx"])
elif bank.startswith("Nubank"):
    uploaded = st.file_uploader("Envie a fatura do Nubank em .pdf", type=["pdf"])
else:
    uploaded = None

if uploaded is not None:
    st.write("Arquivo recebido:", uploaded.name)
    if st.button("▶️ Processar", type="primary"):
        try:
            if bank.startswith("C6"):
                output_bytes = build_processed_workbook_c6(uploaded.read())
            else:
                output_bytes = build_processed_workbook_nubank(uploaded.read())
            st.success("Processamento concluído!")
            st.download_button(
                label="⬇️ Baixar planilha processada",
                data=output_bytes,
                file_name="fatura_processada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Erro ao processar: {e}")
