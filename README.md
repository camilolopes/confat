# C6 Fatura Processor (Streamlit)

App para processar faturas C6 e gerar uma planilha consolidada.

## Rodar localmente
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Community Cloud)
1. Suba estes arquivos para um repositório Git.
2. No Streamlit Cloud, conecte o repositório e selecione `app.py`.
3. Deploy.

## Deploy (Railway/Render)
Use o `Procfile` incluído:
```
web: streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```
