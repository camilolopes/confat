# 📘 VERSIONS.md — Faturas Processor App

## 🔖 Identificação
- **Produto**: Faturas Processor App (C6 + Nubank)  
- **Última Versão**: **v12 (28/09/2025)**  

---

## 📈 Histórico de Versões

### v19 — 28/09/2025
**Status:** Estável

**Correções e melhorias (C6)**:
- **Sanitização da coluna `Parcela`** quando o Excel converte `2/4` em data (ex.: `2025-04-02` → `2/4`).  
- **Respeito à coluna `Parcela` nativa** (não sobrescreve mais a partir de `Descrição` se já existir).  
- **Normalização**: `Única` → `1/1`; `2 de 4`, `2 / 4`, `02/04` → `2/4`.  
- **Aba `Parcelas Ativas` garantida** (sempre criada, populada quando houver parcelas em aberto).  
- **Resumo** com **Compromissos Futuros** e quebra por **cartão/portador** quando aplicável.  

**Nubank**: sem alterações (fluxo independente mantido).
---

### v13 — 28/09/2025
**Fixes:**
- Correção de import no app: exposta a função pública `build_processed_workbook_nubank` no `processor.py` para evitar `ImportError` no `app.py`.
- Sem alterações de comportamento no processamento; apenas ajuste de empacotamento e API do módulo.

---

### v12 — 28/09/2025
**Novidades:**
- Suporte a **parcelamentos**:
  - Extração automática de `Parcela 4/5`, `4 de 5`, `4/05` etc.
  - Novas colunas: **Parcela Nº**, **Qtde Parcelas**, **Restantes**, **É Última?**, **Término Estimado**.
- Nova aba **Parcelas Ativas**:
  - Lista todas transações parceladas **ainda não finalizadas**.
  - Coluna **Compromisso Futuro (R$)** = Valor Parcela × Restantes.
- **Resumo da Fatura**:
  - Inclui **Total Compromissos Futuros (Parcelas)**.
  - Quebra por **cartão (final + portador)**.

---

### v11 — 27/09/2025
**Correções:**
- Função `_extract_holder_candidates_from_pages` não encontrada → fix na ordem de definição.
- Remoção de duplicação da lógica de `holder`.
- Correção na linha `full = "\n".join(texts)`.

---

### v10 — 27/09/2025
**Novidades:**
- Captura robusta do **portador no Nubank** via cabeçalho em **CAIXA ALTA**.
- Se várias páginas → votação para determinar o nome.
- Fallback: heurística textual.

---

### v9 — 27/09/2025
**Melhoria:**
- Filtro para não capturar saudações como nome (ex.: “olá camilo”).
- Heurística só aceita **nomes plausíveis de pessoas** (2–6 palavras, ≥80% letras, sem dígitos).

---

### v8 — 27/09/2025
**Novidades:**
- Parser Nubank ajustado para:
  - Datas `dd MMM` em pt-BR.
  - Nome do portador (heurística).
  - Separação de parcelas no fim da descrição.
  - Categorização básica (alimentação, transporte, seguro etc.).

---

### v7 — 27/09/2025
**Novidade:**  
- Suporte a múltiplos bancos:
  - **C6** (Excel .xlsx).
  - **Nubank** (PDF .pdf).
- Upload só aceita a extensão correta conforme banco escolhido.

---

### v1 (Base Estável) — 26/09/2025
**Funcionalidades:**
- Processamento de faturas do **C6 (Excel)**.
- Saída Excel com:
  - Índice navegável.
  - Consolidado Cartão, Estabelecimento, Categoria por Cartão.
  - Resumo da Fatura.
  - Devoluções.
  - Abas por Cartão (pizza com Top 3 + Outras).
  - Transações Originais (oculta).

---

## 🔄 Como atualizar este documento
1. **Na próxima versão**, adicionar nova seção no topo (`### v13 — [data]`) com:
   - **Novidades**.
   - **Correções**.
   - **Melhorias**.
2. Manter sempre o histórico para rastrear evolução.
3. Opcional: automatizar geração de changelog a partir dos commits (GitHub Actions).  
