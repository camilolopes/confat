# v19 — 28/09/2025 (Estável)

## Principais mudanças
- C6: sanitização e normalização da coluna **Parcela** (datas → `d/m`, `Única` → `1/1`, `2 de 4` → `2/4`).
- C6: não reescreve mais `Parcela` a partir de `Descrição` quando a coluna já existe na planilha.
- C6: aba **Parcelas Ativas** sempre presente e populada quando houver parcelas em aberto.
- Resumo: **Compromissos Futuros** + quebra por cartão/portador.
- Nubank: fluxo inalterado (independente).

## Como atualizar
1. Faça o merge na `main` com uma mensagem contendo `#minor` (ou `#major`/`#patch` conforme desejar).  
2. O workflow unificado **gera tag + Release** e **atualiza o VERSIONS.md** automaticamente.

> Caso prefira publicar manualmente: crie a Release `v19.0.0` e use este texto como notas.
