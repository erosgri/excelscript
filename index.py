import sys
import os
import time

import pandas as pd
import numpy as np
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.units import mm

# =========================
# CONFIG
# =========================
ARQUIVO = "divergencia.xlsx"

COLUNAS = [
    "Nota Fiscal", "Modelo", "Data", "CNPJ/CPF",
    "CCE", "UF", "CFOP", "CST / Mercadoria",
    "Categoria", "Ct Sefaz", "Ct Contrib",
    "Dif CT", "Valor Produto", "Dif. Icms",
]

GRUPOS = [
    ("Documento", 3),
    ("Destino", 4),
    ("Produto", 2),
    ("Tributação", 5),
]

COR_AZUL = colors.HexColor("#4472C4")
COR_AZUL_CLARO = colors.HexColor("#D9E2F3")

STYLE_TITULO = ParagraphStyle(
    "titulo", fontSize=11, leading=13, alignment=TA_CENTER,
    fontName="Helvetica-Bold",
)
STYLE_INFO = ParagraphStyle(
    "info", fontSize=8, leading=10, alignment=TA_LEFT,
)
STYLE_HEADER = ParagraphStyle(
    "header", fontSize=7, leading=9, alignment=TA_CENTER,
    textColor=colors.whitesmoke, fontName="Helvetica-Bold",
)
STYLE_CELL = ParagraphStyle(
    "cell", fontSize=6.5, leading=8, alignment=TA_LEFT,
)


def _safe(val) -> str:
    s = str(val)
    if s in ("nan", "NaT", "None"):
        return ""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def extrair_header(caminho: str) -> dict:
    """Extrai informações do cabeçalho institucional da primeira aba."""
    df = pd.read_excel(caminho, sheet_name=0, header=None, nrows=6)
    info = {}

    row0 = df.iloc[0].fillna("").astype(str)
    info["orgao"] = row0.iloc[0] if row0.iloc[0] != "" else "ESTADO DE GOIAS SECRETARIA DA FAZENDA"
    titulo_parts = [v for v in row0 if "Diverg" in v]
    info["titulo"] = titulo_parts[0] if titulo_parts else "Divergencias de Carga Tributaria Informada e Calculada - Nota Fiscal"

    row1 = df.iloc[1].fillna("").astype(str)
    info_parts = [v for v in row1 if v.strip()]
    info["empresa"] = "  |  ".join(info_parts)

    row2 = df.iloc[2].fillna("").astype(str)
    datas = []
    for v in row2:
        v = v.strip()
        if v and v not in ("Período :", "a", "nan", "NaT", ""):
            try:
                dt = pd.to_datetime(v)
                datas.append(dt.strftime("%d/%m/%Y"))
            except Exception:
                pass
    if len(datas) >= 2:
        info["periodo"] = f"Período: {datas[0]}  a  {datas[1]}"
    else:
        info["periodo"] = "Período: -"

    return info


import re

MARCA_MES = "__MES__"
MARCA_SUBTOTAL = "__SUB__"
_RE_MES = re.compile(r"^\d{1,2}/\d{4}$")


def _find_mes(row_values) -> str | None:
    """Procura um separador de mês (M/AAAA) em qualquer coluna da linha."""
    for v in row_values:
        s = str(v).strip()
        if _RE_MES.match(s):
            return s
    return None


def _find_subtotal(row_values) -> str | None:
    """Detecta linha de subtotal: quase tudo NaN com um valor numérico."""
    strs = [str(v).strip() for v in row_values]
    non_empty = [s for s in strs if s not in ("", "nan", "NaT", "None")]
    if len(non_empty) == 1:
        try:
            float(non_empty[0])
            return non_empty[0]
        except ValueError:
            pass
    return None


def ler_dados(caminho: str) -> pd.DataFrame:
    """Lê todas as abas, preserva separadores de mês e subtotais."""
    if not os.path.isfile(caminho):
        print(f"ERRO: arquivo '{caminho}' nao encontrado.")
        sys.exit(1)

    print("  Carregando abas...")
    t0 = time.time()
    sheets: dict[str, pd.DataFrame] = pd.read_excel(
        caminho, sheet_name=None, header=None, engine="openpyxl",
    )
    print(f"  {len(sheets)} abas lidas em {time.time() - t0:.1f}s")

    n_target = len(COLUNAS)
    all_rows: list[list] = []

    for df in sheets.values():
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            vals = row.values
            first = str(vals[0]).strip()

            if first in ("Documento", "Nota Fiscal", "Período :"):
                continue
            if first.startswith("ESTADO") or first.startswith("Raz"):
                continue

            mes = _find_mes(vals)
            if mes:
                mes_row = [MARCA_MES + mes] + [""] * (n_target - 1)
                all_rows.append(mes_row)
                continue

            sub = _find_subtotal(vals)
            if sub:
                sub_row = [MARCA_SUBTOTAL] + [""] * (n_target - 2) + [sub]
                all_rows.append(sub_row)
                continue

            if first in ("", "nan", "NaT", "None"):
                continue

            try:
                int(float(first))
            except (ValueError, OverflowError):
                continue

            non_nan = [v for v in vals if not (isinstance(v, float) and np.isnan(v))
                       and str(v) not in ("nan", "NaT", "None", "")]
            if len(non_nan) >= 5:
                trimmed = non_nan[:n_target]
                while len(trimmed) < n_target:
                    trimmed.append("")
                all_rows.append(trimmed)

    if not all_rows:
        print("ERRO: nenhum dado encontrado.")
        sys.exit(1)

    df_final = pd.DataFrame(all_rows, columns=COLUNAS)
    return df_final


COR_MES = colors.HexColor("#E2EFDA")
STYLE_MES = ParagraphStyle(
    "mes", fontSize=8, leading=10, alignment=TA_LEFT,
    fontName="Helvetica-Bold",
)


def gerar_pdf(df: pd.DataFrame, saida: str, header_info: dict) -> None:
    df_str = df.fillna("").astype(str).replace("nan", "")
    n_rows, n_cols = df_str.shape
    print(f"  Montando PDF ({n_rows} linhas x {n_cols} colunas)...")
    t0 = time.time()

    # --- Cabeçalho institucional ---
    elements = []
    elements.append(Paragraph(_safe(header_info.get("titulo", "")), STYLE_TITULO))
    elements.append(Spacer(1, 3 * mm))
    elements.append(Paragraph(_safe(header_info.get("orgao", "")), STYLE_INFO))
    elements.append(Spacer(1, 1.5 * mm))
    elements.append(Paragraph(_safe(header_info.get("empresa", "")), STYLE_INFO))
    elements.append(Spacer(1, 1.5 * mm))
    elements.append(Paragraph(_safe(header_info.get("periodo", "")), STYLE_INFO))
    elements.append(Spacer(1, 4 * mm))

    # --- Linha de grupo ---
    grupo_row = []
    for nome, span in GRUPOS:
        grupo_row.append(Paragraph(f"<b>{_safe(nome)}</b>", STYLE_HEADER))
        for _ in range(span - 1):
            grupo_row.append("")
    while len(grupo_row) < n_cols:
        grupo_row.append("")
    grupo_row = grupo_row[:n_cols]

    # --- Linha de colunas ---
    col_row = [
        Paragraph(f"<b>{_safe(c)}</b>", STYLE_HEADER)
        for c in df_str.columns
    ]

    # --- Dados (preservando meses e subtotais) ---
    values = df_str.values.tolist()
    data = [grupo_row, col_row]
    mes_row_indices = []
    sub_row_indices = []

    for row in values:
        first = str(row[0])
        if first.startswith(MARCA_MES):
            mes_label = first.replace(MARCA_MES, "")
            mes_cells = [Paragraph(f"<b>{_safe(mes_label)}</b>", STYLE_MES)]
            mes_cells += [""] * (n_cols - 1)
            data.append(mes_cells)
            mes_row_indices.append(len(data) - 1)
        elif first.startswith(MARCA_SUBTOTAL):
            sub_cells = [""] * (n_cols - 1)
            sub_cells.append(Paragraph(f"<b>{_safe(str(row[-1]))}</b>", STYLE_MES))
            data.append(sub_cells)
            sub_row_indices.append(len(data) - 1)
        else:
            data.append([_safe(str(cell)) for cell in row])

    print(f"  Tabela montada em {time.time() - t0:.1f}s")
    print(f"  {len(mes_row_indices)} separadores de mes, {len(sub_row_indices)} subtotais")

    # --- Larguras ---
    CELL_PAD = 2
    MIN_COL_W = 20
    page_width, page_height = landscape(A3)
    MARGIN = 15
    usable_width = page_width - (MARGIN * 2)

    mask = ~(df_str.iloc[:, 0].str.startswith(MARCA_MES) |
             df_str.iloc[:, 0].str.startswith(MARCA_SUBTOTAL))
    df_data = df_str[mask]
    max_lens = df_data.apply(lambda c: c.map(len).max()).values
    header_lens = [len(str(c)) for c in df_str.columns]
    col_widths = [
        max(min(max(ml, hl) * 4.5, 250), MIN_COL_W)
        for ml, hl in zip(max_lens, header_lens)
    ]
    total = sum(col_widths)
    if total > usable_width:
        scale = usable_width / total
        col_widths = [max(w * scale, MIN_COL_W) for w in col_widths]

    # --- Estilos ---
    style_cmds = [
        ("GRID",       (0, 0), (-1, -1), 0.25, colors.black),
        ("BACKGROUND", (0, 0), (-1, 1),  COR_AZUL),
        ("TEXTCOLOR",  (0, 0), (-1, 1),  colors.whitesmoke),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE",   (0, 2), (-1, -1), 6.5),
        ("ALIGN",      (0, 0), (-1, 1),  "CENTER"),
        ("LEFTPADDING",   (0, 0), (-1, -1), CELL_PAD),
        ("RIGHTPADDING",  (0, 0), (-1, -1), CELL_PAD),
        ("TOPPADDING",    (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]

    # Spans do grupo (linha 0)
    col_idx = 0
    for _, span in GRUPOS:
        if span > 1 and col_idx + span - 1 < n_cols:
            style_cmds.append(
                ("SPAN", (col_idx, 0), (col_idx + span - 1, 0))
            )
        col_idx += span

    # Estilo das linhas de mês e subtotal
    special = set(mes_row_indices) | set(sub_row_indices)

    for mes_idx in mes_row_indices:
        style_cmds.append(("SPAN", (0, mes_idx), (n_cols - 1, mes_idx)))
        style_cmds.append(("BACKGROUND", (0, mes_idx), (-1, mes_idx), COR_MES))
        style_cmds.append(("ALIGN", (0, mes_idx), (-1, mes_idx), "LEFT"))

    for sub_idx in sub_row_indices:
        style_cmds.append(("BACKGROUND", (0, sub_idx), (-1, sub_idx), colors.HexColor("#FFF2CC")))
        style_cmds.append(("ALIGN", (n_cols - 1, sub_idx), (n_cols - 1, sub_idx), "RIGHT"))

    # Linhas de dados alternadas (ignorando especiais)
    data_counter = 0
    for r in range(2, len(data)):
        if r in special:
            continue
        bg = colors.white if data_counter % 2 == 0 else COR_AZUL_CLARO
        style_cmds.append(("BACKGROUND", (0, r), (-1, r), bg))
        data_counter += 1

    doc = SimpleDocTemplate(
        saida,
        pagesize=landscape(A3),
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN,
    )

    table = Table(data, colWidths=col_widths, repeatRows=2)
    table.setStyle(TableStyle(style_cmds))

    elements.append(table)

    print("  Gerando PDF...")
    t0 = time.time()
    doc.build(elements)
    print(f"  PDF gerado em {time.time() - t0:.1f}s")


def main():
    t_total = time.time()
    nome_base = os.path.splitext(ARQUIVO)[0]
    saida_excel = f"{nome_base}_unificado.xlsx"
    saida_pdf = f"{nome_base}.pdf"

    print(f"Lendo '{ARQUIVO}'...")

    header_info = extrair_header(ARQUIVO)
    print(f"  Titulo: {header_info.get('titulo', '-')}")
    print(f"  {header_info.get('periodo', '-')}")

    df_final = ler_dados(ARQUIVO)
    print(f"  -> {len(df_final)} linhas | {len(df_final.columns)} colunas")

    xls = pd.ExcelFile(ARQUIVO)
    if len(xls.sheet_names) > 1:
        print("Salvando Excel unificado...")
        mask = ~(df_final.iloc[:, 0].astype(str).str.startswith(MARCA_MES) |
                 df_final.iloc[:, 0].astype(str).str.startswith(MARCA_SUBTOTAL))
        df_excel = df_final[mask].copy()
        df_excel.to_excel(saida_excel, index=False)
        print(f"  {saida_excel} salvo ({len(df_excel)} linhas de dados).")
    else:
        print("  Apenas 1 aba encontrada, pulando unificacao do Excel.")

    gerar_pdf(df_final, saida_pdf, header_info)
    print(f"  {saida_pdf} salvo.")

    print(f"Concluido em {time.time() - t_total:.1f}s")


if __name__ == "__main__":
    main()
