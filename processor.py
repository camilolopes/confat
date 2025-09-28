
import io, re, unicodedata
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

def _autosize(ws):
    for c in range(1, ws.max_column + 1):
        max_len = 0
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(c)].width = max_len + 2

def _set_currency(ws, header_text="Valor BRL", header_row=1):
    headers = {ws.cell(row=header_row, column=i).value: i for i in range(1, ws.max_column + 1)}
    if header_text in headers:
        col = headers[header_text]
        for r in range(header_row + 1, ws.max_row + 1):
            ws.cell(row=r, column=col).number_format = u'R$ #,##0.00'

def _write_df(ws, df, start_row=1, start_col=1):
    for j, col in enumerate(df.columns, start=start_col):
        ws.cell(row=start_row, column=j, value=str(col))
    for i in range(len(df)):
        for j, col in enumerate(df.columns, start=start_col):
            val = df.iloc[i][col]
            ws.cell(row=start_row + 1 + i, column=j, value=(None if pd.isna(val) else val))

def _normalize_header(s):
    if s is None:
        return ""
    import re, unicodedata
    t = unicodedata.normalize("NFKD", str(s)).encode("ascii","ignore").decode("ascii")
    t = t.lower()
    t = re.sub(r"r\$|\\(r\\$?\\)|currency|valor\\s*\\(.*\\)", "valor", t)
    t = re.sub(r"[^a-z0-9]+", " ", t)
    t = re.sub(r"\\s+", " ", t).strip()
    return t

def _best_sheet(excel_io):
    # Try to pick the sheet with the most overlap with expected columns
    xl = pd.ExcelFile(excel_io)
    candidates = []
    expected_tokens = [
        ("nome no cartao", {"nome","cartao"}),
        ("final do cartao", {"final","cartao"}),
        ("categoria", {"categoria"}),
        ("descricao", {"descricao","estabelecimento","loja","merchant"}),
        ("valor brl", {"valor"}),
    ]
    for sheet in xl.sheet_names:
        try:
            df_tmp = xl.parse(sheet, nrows=3)
            norm = [_normalize_header(c) for c in df_tmp.columns]
            score = 0
            for label, tokens in expected_tokens:
                if any(all(tok in h for tok in tokens) for h in norm):
                    score += 1
            candidates.append((score, sheet))
        except Exception:
            continue
    candidates.sort(reverse=True)
    return candidates[0][1] if candidates else None

def _coerce_brl(x):
    if pd.isna(x): return None
    s = str(x)
    # remove currency and spaces
    s = s.replace("R$", "").replace(" ", "")
    # keep only digits , . -
    import re
    s = re.sub(r"[^0-9,.-]", "", s)
    # if there is comma and dot, assume comma as decimal if comma is last separator
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "")
        s = s.replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        try:
            return pd.to_numeric(s, errors="coerce")
        except:
            return None

def _generate_pie_image(series_df, title, img_path, text_fontsize=8, title_fontsize=11):
    total = series_df["Valor BRL"].sum()
    labels = [
        f"{cat}\\n{val/total:.1%} • R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        for cat, val in zip(series_df["Categoria"], series_df["Valor BRL"])
    ]
    plt.figure(figsize=(8, 8))
    plt.pie(
        series_df["Valor BRL"],
        labels=labels,
        startangle=90,
        colors=plt.cm.Set2.colors,
        textprops={"fontsize": 8},
    )
    plt.title(title, fontsize=11)
    plt.tight_layout()
    plt.savefig(img_path, bbox_inches="tight")
    plt.close()

def build_processed_workbook(file_bytes: bytes) -> bytes:
    excel_io = io.BytesIO(file_bytes)
    # Pick best sheet
    try:
        df = pd.read_excel(excel_io, sheet_name="Transações Originais")
    except Exception:
        excel_io.seek(0)
        best = _best_sheet(excel_io)
        excel_io.seek(0)
        df = pd.read_excel(excel_io, sheet_name=best if best is not None else 0)

    # Normalize headers
    original_cols = list(df.columns)
    norm_map = {_normalize_header(c): c for c in df.columns}

    def find_col(*tokens_sets):
        # tokens_sets: list of sets of tokens that should all appear
        for norm, orig in norm_map.items():
            for tokens in tokens_sets:
                if all(tok in norm for tok in tokens):
                    return orig
        return None

    col_nome = find_col({"nome","cartao"}, {"portador"}, {"titular"})
    col_final = find_col({"final","cartao"}, {"final","****"}, {"cartao","final"})
    col_categoria = find_col({"categoria"})
    col_descricao = find_col({"descricao"}, {"estabelecimento"}, {"loja"}, {"merchant"})
    col_valor = find_col({"valor"})
    col_data = find_col({"data"})  # optional

    required = {
        "Nome no Cartão": col_nome,
        "Final do Cartão": col_final,
        "Categoria": col_categoria,
        "Descrição": col_descricao,
        "Valor BRL": col_valor
    }
    missing = [k for k,v in required.items() if v is None]
    if missing:
        # include debug info to help user
        raise ValueError(f"Não encontrei colunas: {missing}. Colunas do arquivo: {original_cols}")

    # Rename to canonical
    df = df.rename(columns={
        col_nome: "Nome no Cartão",
        col_final: "Final do Cartão",
        col_categoria: "Categoria",
        col_descricao: "Descrição",
        col_valor: "Valor BRL",
        **({col_data: "Data"} if col_data else {})
    })

    # Coerce numeric BRL
    df["Valor BRL"] = df["Valor BRL"].apply(_coerce_brl)

    # Dates
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    # Ensure Final do Cartão is last 4 digits if masked
    def last4(x):
        s = str(x)
        m = re.findall(r"(\\d{4})", s)
        return m[-1] if m else s[-4:]
    df["Final do Cartão"] = df["Final do Cartão"].apply(last4)

    # Positive/Negative splits
    df_pos = df[df["Valor BRL"] > 0].copy()
    df_neg = df[df["Valor BRL"] < 0].copy()

    # Consolidations
    consol_cartao = (
        df_pos.groupby(["Final do Cartão", "Nome no Cartão", "Descrição"], as_index=False)["Valor BRL"]
        .sum()
        .sort_values(["Final do Cartão", "Valor BRL"], ascending=[True, False])
        .rename(columns={"Nome no Cartão": "Nome do Portador"})
    )
    consol_estab = (
        df_pos.groupby(["Nome no Cartão", "Final do Cartão", "Descrição"], as_index=False)["Valor BRL"]
        .sum()
        .sort_values(["Nome no Cartão", "Final do Cartão", "Valor BRL"], ascending=[True, True, False])
        .rename(columns={"Nome no Cartão": "Nome do Portador"})
    )
    consol_cat_cartao = (
        df_pos.groupby(["Final do Cartão", "Nome no Cartão", "Categoria"], as_index=False)["Valor BRL"]
        .sum()
        .sort_values(["Final do Cartão", "Valor BRL"], ascending=[True, False])
        .rename(columns={"Nome no Cartão": "Nome do Portador"})
    )
    resumo = pd.DataFrame({
        "Total Fatura (R$)": [df["Valor BRL"].sum()],
        "Total Sem Devoluções (R$)": [df_pos["Valor BRL"].sum()],
        "Total Devoluções (R$)": [df_neg["Valor BRL"].sum()],
    })

    # Build workbook
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    def _write_sheet_consol(name, data, header_row=1):
        ws = wb.create_sheet(name)
        # write
        for j, col in enumerate(data.columns, start=1):
            ws.cell(row=header_row, column=j, value=str(col))
        for i in range(len(data)):
            for j, col in enumerate(data.columns, start=1):
                ws.cell(row=header_row + 1 + i, column=j, value=None if pd.isna(data.iloc[i][col]) else data.iloc[i][col])
        # format
        headers_idx = {ws.cell(row=header_row, column=i).value: i for i in range(1, data.shape[1] + 1)}
        if "Valor BRL" in headers_idx:
            c = headers_idx["Valor BRL"]
            for r in range(header_row + 1, header_row + 1 + len(data)):
                ws.cell(row=r, column=c).number_format = u'R$ #,##0.00'
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
        _autosize(ws)
        return ws

    _write_sheet_consol("Consolidado Cartão", consol_cartao)
    # Estabelecimento com nota simples
    ws_ce = _write_sheet_consol("Consolidado Estabelecimento", consol_estab, header_row=2)
    ws_ce.insert_rows(1)
    ws_ce["A1"] = "NOTA: 'Final do Cartão' = últimos 4 dígitos; 'Nome do Portador' = nome impresso. Somente valores positivos."
    ws_ce.freeze_panes = "A3"

    _write_sheet_consol("Consolidado Cat por Cartão", consol_cat_cartao)

    # Devoluções
    cols_dev = ["Data","Nome no Cartão","Final do Cartão","Categoria","Descrição","Parcela","Valor BRL"]
    cols_dev_present = [c for c in cols_dev if c in df_neg.columns]
    ws_dev = wb.create_sheet("Devoluções")
    # write headers
    for j, col in enumerate(cols_dev_present, start=1):
        ws_dev.cell(row=1, column=j, value=str(col if col != "Nome no Cartão" else "Nome do Portador"))
    # write rows
    for i in range(len(df_neg)):
        for j, col in enumerate(cols_dev_present, start=1):
            val = df_neg.iloc[i][col]
            ws_dev.cell(row=2 + i, column=j, value=None if pd.isna(val) else val)
    # currency
    if "Valor BRL" in cols_dev_present:
        col_idx = cols_dev_present.index("Valor BRL") + 1
        for r in range(2, ws_dev.max_row + 1):
            ws_dev.cell(row=r, column=col_idx).number_format = u'R$ #,##0.00'
    ws_dev.auto_filter.ref = f"A1:{get_column_letter(ws_dev.max_column)}{ws_dev.max_row}"
    _autosize(ws_dev)

    # Resumo
    ws_rf = wb.create_sheet("Resumo Fatura")
    ws_rf.cell(row=1, column=1, value="Total Fatura (R$)")
    ws_rf.cell(row=1, column=2, value=resumo.iloc[0,0])
    ws_rf.cell(row=2, column=1, value="Total Sem Devoluções (R$)")
    ws_rf.cell(row=2, column=2, value=resumo.iloc[0,1])
    ws_rf.cell(row=3, column=1, value="Total Devoluções (R$)")
    ws_rf.cell(row=3, column=2, value=resumo.iloc[0,2])
    for r in range(1,4):
        ws_rf.cell(row=r, column=2).number_format = u'R$ #,##0.00'
    _autosize(ws_rf)

    # Transações Originais (oculta)
    ws_to = wb.create_sheet("Transações Originais")
    for j, col in enumerate(df.columns, start=1):
        ws_to.cell(row=1, column=j, value=str(col))
    for i in range(len(df)):
        for j, col in enumerate(df.columns, start=1):
            ws_to.cell(row=2+i, column=j, value=None if pd.isna(df.iloc[i][col]) else df.iloc[i][col])
    ws_to.sheet_state = "hidden"

    # Índice
    ws_idx = wb.create_sheet("Índice", 0)
    ws_idx["A1"] = "📑 Índice de Navegação"
    row = 3
    for name in ["Consolidado Cartão","Consolidado Estabelecimento","Consolidado Cat por Cartão","Resumo Fatura","Devoluções"]:
        ws_idx.cell(row=row, column=1, value=name)
        ws_idx.cell(row=row, column=1).hyperlink = f"#{name}!A1"
        ws_idx.cell(row=row, column=1).style = "Hyperlink"
        row += 1
    ws_idx.cell(row=row, column=1, value="---"); row += 1
    ws_idx.cell(row=row, column=1, value="Cartões (Mapa de Calor – Top 3 + Outras):"); row += 1

    # Holder map
    portador_map = (
        df_pos.groupby(["Final do Cartão", "Nome no Cartão"])["Valor BRL"].sum()
        .reset_index()
        .sort_values(["Final do Cartão", "Valor BRL"], ascending=[True, False])
        .drop_duplicates(subset=["Final do Cartão"])
        .set_index("Final do Cartão")["Nome no Cartão"].to_dict()
    )

    # Card tabs with pies (Top3 + Outras), hide if <=2 categories
    cats_por_cartao = df_pos.groupby("Final do Cartão")["Categoria"].nunique().to_dict()
    gastos_por_cartao_cat = (
        df_pos.groupby(["Final do Cartão", "Categoria"], as_index=False)["Valor BRL"].sum()
    )
    for final_cartao, grupo in gastos_por_cartao_cat.groupby("Final do Cartão"):
        if grupo.shape[0] == 0:
            continue
        sheet_name = f"Cartão {final_cartao}"
        ws_card = wb.create_sheet(sheet_name)
        ws_card["A1"] = f"Mapa de Calor - Cartão {final_cartao} (Top 3 + Outras)"

        tabela = grupo.sort_values("Valor BRL", ascending=False).reset_index(drop=True)
        if tabela.shape[0] > 3:
            top3 = tabela.head(3).copy()
            outras_val = float(tabela["Valor BRL"].sum() - top3["Valor BRL"].sum())
            if outras_val > 0:
                top3 = pd.concat([top3, pd.DataFrame([{"Categoria": "Outras", "Valor BRL": outras_val}])], ignore_index=True)
            tabela = top3

        holder = portador_map.get(str(final_cartao), "")
        chart_title = f"Distribuição de Gastos – Cartão {final_cartao}"
        if holder:
            chart_title += f" – {holder}"

        total = tabela["Valor BRL"].sum()
        labels = [f"{cat}\\n{val/total:.1%} • R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                  for cat, val in zip(tabela["Categoria"], tabela["Valor BRL"])]
        plt.figure(figsize=(8, 8))
        plt.pie(tabela["Valor BRL"], labels=labels, startangle=90, colors=plt.cm.Set2.colors, textprops={"fontsize": 8})
        plt.title(chart_title, fontsize=11)
        plt.tight_layout()
        img_path = f"/mnt/data/pizza_cartao_runtime_{final_cartao}.png"
        plt.savefig(img_path, bbox_inches="tight")
        plt.close()
        ws_card.add_image(XLImage(img_path), "A3")

        if cats_por_cartao.get(final_cartao, 0) <= 2:
            ws_card.sheet_state = "hidden"

        display_name = sheet_name + (f" – {holder}" if holder else "")
        ws_idx.cell(row=row, column=1, value=display_name)
        ws_idx.cell(row=row, column=1).hyperlink = f"#{sheet_name}!A1"
        ws_idx.cell(row=row, column=1).style = "Hyperlink"
        row += 1

    ws_idx.column_dimensions["A"].width = 65

    # Reorder after index
    desired_after_index = ["Consolidado Cartão", "Consolidado Estabelecimento", "Consolidado Cat por Cartão"]
    current = wb.sheetnames
    ordered = ["Índice"] + [s for s in desired_after_index if s in current]
    for s in current:
        if s not in ordered:
            ordered.append(s)
    wb._sheets = [wb[s] for s in ordered]

    out_io = io.BytesIO()
    wb.save(out_io)
    out_io.seek(0)
    return out_io.getvalue()
