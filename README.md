# Processador de Faturas (C6 & Nubank)

Nova versão com seleção de banco e suporte a PDF do Nubank.
- C6 -> .xlsx
- Nubank -> .pdf
O app bloqueia extensões incompatíveis por banco.

## Rodar localmente
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Cloud / Railway / Render)
Use o Procfile:
```
web: streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```
