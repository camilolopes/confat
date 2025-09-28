import io
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

def _generate_pie_image(series_df, title, img_path, text_fontsize=8, title_fontsize=11):
    total = series_df["Valor BRL"].sum()
    labels = [
        f"{cat}\n{val/total:.1%} • R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
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
    try:
        df = pd.read_excel(excel_io, sheet_name="Transações Originais")
    except Exception:
        excel_io.seek(0)
        df = pd.read_excel(excel_io, sheet_name=0)

    df["Valor BRL"] = pd.to_numeric(df.get("Valor BRL"), errors="coerce")
    df["Data"] = pd.to_datetime(df.get("Data"), errors="coerce")

    required_cols = ["Nome no Cartão","Final do Cartão","Categoria","Descrição","Valor BRL"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Colunas obrigatórias ausentes: {missing}. Verifique o layout do arquivo.")

    df_pos = df[df["Valor BRL"] > 0].copy()
    df_neg = df[df["Valor BRL"] < 0].copy()

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

    wb = Workbook()
    default = wb.active
    wb.remove(default)

    ws_cc = wb.create_sheet("Consolidado Cartão")
    _write_df(ws_cc, consol_cartao)
    _set_currency(ws_cc, "Valor BRL", header_row=1)
    ws_cc.auto_filter.ref = f"A1:{get_column_letter(ws_cc.max_column)}{ws_cc.max_row}"
    _autosize(ws_cc)

    ws_ce = wb.create_sheet("Consolidado Estabelecimento")
    ws_ce.append(["NOTA: 'Final do Cartão' = últimos 4 dígitos; 'Nome do Portador' = nome impresso. Somente valores positivos."])
    for j, col in enumerate(consol_estab.columns, start=1):
        ws_ce.cell(row=2, column=j, value=str(col))
    for i in range(len(consol_estab)):
        for j, col in enumerate(consol_estab.columns, start=1):
            ws_ce.cell(row=3 + i, column=j, value=None if pd.isna(consol_estab.iloc[i][col]) else consol_estab.iloc[i][col])
    _set_currency(ws_ce, "Valor BRL", header_row=2)
    ws_ce.auto_filter.ref = f"A2:{get_column_letter(ws_ce.max_column)}{ws_ce.max_row}"
    ws_ce.freeze_panes = "A3"
    _autosize(ws_ce)

    ws_cpc = wb.create_sheet("Consolidado Cat por Cartão")
    _write_df(ws_cpc, consol_cat_cartao)
    _set_currency(ws_cpc, "Valor BRL", header_row=1)
    ws_cpc.auto_filter.ref = f"A1:{get_column_letter(ws_cpc.max_column)}{ws_cpc.max_row}"
    _autosize(ws_cpc)

    ws_dev = wb.create_sheet("Devoluções")
    cols_dev = ["Data","Nome no Cartão","Final do Cartão","Categoria","Descrição","Parcela","Valor BRL"]
    df_dev = df_neg[cols_dev].rename(columns={"Nome no Cartão":"Nome do Portador"})
    _write_df(ws_dev, df_dev)
    _set_currency(ws_dev, "Valor BRL", header_row=1)
    ws_dev.auto_filter.ref = f"A1:{get_column_letter(ws_dev.max_column)}{ws_dev.max_row}"
    _autosize(ws_dev)

    ws_rf = wb.create_sheet("Resumo Fatura")
    _write_df(ws_rf, resumo)
    for c in range(1, ws_rf.max_column + 1):
        for r in range(2, ws_rf.max_row + 1):
            val = ws_rf.cell(row=r, column=c).value
            if isinstance(val, (int, float)):
                ws_rf.cell(row=r, column=c).number_format = u'R$ #,##0.00'
    _autosize(ws_rf)

    ws_to = wb.create_sheet("Transações Originais")
    _write_df(ws_to, df)
    ws_to.sheet_state = "hidden"

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

    portador_por_final = (
        df_pos.groupby(["Final do Cartão", "Nome no Cartão"])["Valor BRL"].sum()
        .reset_index()
        .sort_values(["Final do Cartão", "Valor BRL"], ascending=[True, False])
        .drop_duplicates(subset=["Final do Cartão"])
        .set_index("Final do Cartão")["Nome no Cartão"].to_dict()
    )

    cats_por_cartao = df_pos.groupby("Final do Cartão")["Categoria"].nunique().to_dict()
    gastos_por_cartao_cat = (
        df_pos.groupby(["Final do Cartão", "Categoria"], as_index=False)["Valor BRL"].sum()
    )

    img_paths = []
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

        holder = portador_por_final.get(str(final_cartao), "")
        chart_title = f"Distribuição de Gastos – Cartão {final_cartao}"
        if holder:
            chart_title += f" – {holder}"

        img_path = f"/mnt/data/pizza_cartao_runtime_{final_cartao}.png"
        total = tabela["Valor BRL"].sum()
        labels = [
            f"{cat}\n{val/total:.1%} • R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            for cat, val in zip(tabela["Categoria"], tabela["Valor BRL"])
        ]
        plt.figure(figsize=(8, 8))
        plt.pie(
            tabela["Valor BRL"],
            labels=labels,
            startangle=90,
            colors=plt.cm.Set2.colors,
            textprops={"fontsize": 8},
        )
        plt.title(chart_title, fontsize=11)
        plt.tight_layout()
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
