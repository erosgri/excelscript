"""
Dupla verificação: compara o PDF gerado com o Excel unificado (.xlsx)
e com o arquivo original, célula a célula e separador a separador.

Uso: python verificar_pdf.py [nome_base]
  Exemplo: python verificar_pdf.py divergencia2
  Compara divergencia2.pdf  ↔  divergencia2_unificado.xlsx  ↔  divergencia2.xlsx
"""

import sys
import os
import re
import time

import pandas as pd
import pdfplumber


COLUNAS = [
    "Nota Fiscal", "Modelo", "Data", "CNPJ/CPF",
    "CCE", "UF", "CFOP", "CST / Mercadoria",
    "Categoria", "Ct Sefaz", "Ct Contrib",
    "Dif CT", "Valor Produto", "Dif. Icms",
]

_RE_MES = re.compile(r"^\d{1,2}/\d{4}$")


def _normalizar(val: str) -> str:
    """Normaliza valor para comparação: remove espaços, .0 de inteiros, formata datas."""
    s = str(val).strip()
    if s in ("nan", "NaT", "None", "NaN", ""):
        return ""

    s = s.replace("\n", " ").replace("\r", "")
    s = re.sub(r"\s+", " ", s).strip()

    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        try:
            dt = pd.to_datetime(s)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass

    try:
        f = float(s)
        if f == int(f) and "." not in s.rstrip("0").rstrip("."):
            return str(int(f))
        return f"{f:g}"
    except (ValueError, OverflowError):
        pass

    return s


def extrair_pdf(caminho_pdf: str) -> list[list[str]]:
    """Extrai todas as linhas da tabela principal do PDF."""
    print(f"  Extraindo dados do PDF '{caminho_pdf}'...")
    t0 = time.time()

    all_rows = []
    header_found = False
    n_cols = len(COLUNAS)

    with pdfplumber.open(caminho_pdf) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance": 5,
            })

            if not tables:
                continue

            for table in tables:
                for row in table:
                    if row is None:
                        continue

                    cells = [str(c).strip() if c else "" for c in row]

                    if not header_found:
                        if any("Nota Fiscal" in c for c in cells):
                            header_found = True
                        continue

                    if any("Nota Fiscal" in c for c in cells):
                        continue

                    grupo = any(c in ("Documento", "Destino", "Produto", "Tributação")
                                for c in cells if c)
                    if grupo:
                        continue

                    if len(cells) < n_cols:
                        cells += [""] * (n_cols - len(cells))
                    cells = cells[:n_cols]

                    all_rows.append(cells)

            if page_num % 50 == 0:
                print(f"    ... pagina {page_num}/{total_pages}")

    print(f"  {len(all_rows)} linhas extraidas do PDF em {time.time() - t0:.1f}s")
    return all_rows


def carregar_excel(caminho_xlsx: str) -> pd.DataFrame:
    """Carrega o Excel unificado."""
    print(f"  Carregando Excel '{caminho_xlsx}'...")
    df = pd.read_excel(caminho_xlsx)
    print(f"  {len(df)} linhas no Excel")
    return df


def comparar(df_excel: pd.DataFrame, pdf_rows: list[list[str]]) -> dict:
    """Compara Excel vs PDF célula a célula."""
    resultado = {
        "linhas_excel": len(df_excel),
        "linhas_pdf": 0,
        "linhas_comparadas": 0,
        "celulas_ok": 0,
        "celulas_divergentes": 0,
        "divergencias": [],
        "pdf_meses": {},
    }

    pdf_data = []
    for row in pdf_rows:
        first = row[0]
        if _RE_MES.match(first) or first == "RESUMO":
            total = row[-1].strip() if row[-1] else ""
            if total:
                try:
                    resultado["pdf_meses"][first] = float(total)
                except ValueError:
                    resultado["pdf_meses"][first] = total
            else:
                resultado["pdf_meses"][first] = None
            continue
        if first == "" and all(c == "" for c in row[:-1]) and row[-1] != "":
            continue
        pdf_data.append(row)

    resultado["linhas_pdf"] = len(pdf_data)
    n_compare = min(len(df_excel), len(pdf_data))
    resultado["linhas_comparadas"] = n_compare
    n_cols = len(COLUNAS)

    max_divergencias = 100

    for i in range(n_compare):
        excel_row = df_excel.iloc[i]
        pdf_row = pdf_data[i]

        for j in range(n_cols):
            col_name = COLUNAS[j]
            val_excel = _normalizar(excel_row.iloc[j])
            val_pdf = _normalizar(pdf_row[j])

            if val_excel == val_pdf:
                resultado["celulas_ok"] += 1
                continue

            try:
                fe = float(val_excel) if val_excel else None
                fp = float(val_pdf) if val_pdf else None
                if fe is not None and fp is not None and abs(fe - fp) < 0.01:
                    resultado["celulas_ok"] += 1
                    continue
            except (ValueError, TypeError):
                pass

            if val_excel in val_pdf or val_pdf in val_excel:
                resultado["celulas_ok"] += 1
                continue

            resultado["celulas_divergentes"] += 1
            if len(resultado["divergencias"]) < max_divergencias:
                resultado["divergencias"].append({
                    "linha": i + 1,
                    "coluna": col_name,
                    "excel": val_excel[:60],
                    "pdf": val_pdf[:60],
                })

    return resultado


def extrair_meses_original(caminho_xlsx: str) -> dict[str, float | None]:
    """Extrai meses e totais das abas de resumo do arquivo original."""
    sheets = pd.read_excel(caminho_xlsx, sheet_name=None, header=None, engine="openpyxl")
    meses = {}
    _re = re.compile(r"\d{1,2}/\d{4}")

    for _, df in sheets.items():
        if df is None or df.empty:
            continue
        for _, row in df.iterrows():
            vals = row.values
            first = str(vals[0]).strip()

            if first.startswith("Resumo"):
                for c in reversed(range(len(vals))):
                    s = str(vals[c]).strip()
                    if s not in ("", "nan", "NaT", "None", "Resumo:"):
                        try:
                            meses["RESUMO"] = float(s)
                        except ValueError:
                            pass
                        break
                continue

            m = _re.search(first)
            if m:
                mes = m.group()
                strs = [str(v).strip() for v in vals]
                non_empty = [s for s in strs if s not in ("", "nan", "NaT", "None")]
                if len(non_empty) <= 2:
                    for s in reversed(non_empty):
                        if s != mes and not s.startswith("Refer"):
                            try:
                                meses[mes] = float(s)
                            except ValueError:
                                pass
                            break

    return meses


def verificar_meses(pdf_meses: dict, caminho_original: str) -> dict:
    """Compara separadores de mês do PDF com o arquivo original."""
    print(f"  Verificando separadores de mes contra '{caminho_original}'...")
    original = extrair_meses_original(caminho_original)

    res = {
        "total_original": len(original),
        "total_pdf": len(pdf_meses),
        "ok": 0,
        "faltando": [],
        "valor_diferente": [],
    }

    for mes, val in sorted(original.items(),
                           key=lambda x: (0 if x[0] == "RESUMO" else 1,
                                          x[0])):
        if mes not in pdf_meses:
            res["faltando"].append((mes, val))
        elif val is not None and pdf_meses[mes] is not None:
            try:
                if abs(float(val) - float(pdf_meses[mes])) < 0.01:
                    res["ok"] += 1
                else:
                    res["valor_diferente"].append(
                        (mes, val, pdf_meses[mes]))
            except (ValueError, TypeError):
                if str(val) == str(pdf_meses[mes]):
                    res["ok"] += 1
                else:
                    res["valor_diferente"].append(
                        (mes, val, pdf_meses[mes]))
        else:
            res["ok"] += 1

    return res


def relatorio(res: dict) -> None:
    """Imprime relatório da comparação."""
    print()
    print("=" * 60)
    print("  DUPLA VERIFICACAO: PDF vs EXCEL")
    print("=" * 60)
    print()
    print(f"  Linhas no Excel:      {res['linhas_excel']:>8}")
    print(f"  Linhas no PDF:        {res['linhas_pdf']:>8}")
    print(f"  Linhas comparadas:    {res['linhas_comparadas']:>8}")
    print()

    total_celulas = res["celulas_ok"] + res["celulas_divergentes"]
    if total_celulas > 0:
        pct = (res["celulas_ok"] / total_celulas) * 100
    else:
        pct = 0

    print(f"  Celulas verificadas:  {total_celulas:>8}")
    print(f"  Celulas OK:           {res['celulas_ok']:>8}  ({pct:.2f}%)")
    print(f"  Celulas divergentes:  {res['celulas_divergentes']:>8}")
    print()

    diff_linhas = abs(res["linhas_excel"] - res["linhas_pdf"])
    if diff_linhas > 0:
        print(f"  [!] Diferenca de {diff_linhas} linhas entre Excel e PDF")
        print()

    if res["divergencias"]:
        print(f"  Primeiras {len(res['divergencias'])} divergencias encontradas:")
        print(f"  {'Linha':>6} | {'Coluna':<18} | {'Excel':<25} | {'PDF':<25}")
        print(f"  {'-'*6}-+-{'-'*18}-+-{'-'*25}-+-{'-'*25}")
        for d in res["divergencias"]:
            print(f"  {d['linha']:>6} | {d['coluna']:<18} | {d['excel']:<25} | {d['pdf']:<25}")
        print()

    if res["celulas_divergentes"] == 0 and diff_linhas == 0:
        print("  RESULTADO DADOS: PDF e Excel IDENTICOS")
    elif res["celulas_divergentes"] == 0 and diff_linhas > 0:
        print(f"  RESULTADO DADOS: Conteudo OK ({diff_linhas} linha(s) de separador a mais no PDF)")
    else:
        print(f"  RESULTADO DADOS: {res['celulas_divergentes']} divergencia(s) encontrada(s)")

    print("=" * 60)


def relatorio_meses(res_meses: dict) -> None:
    """Imprime relatório da verificação de separadores de mês."""
    print()
    print("=" * 60)
    print("  VERIFICACAO DE SEPARADORES DE MES (PDF vs Original)")
    print("=" * 60)
    print()
    print(f"  Meses no original:    {res_meses['total_original']:>8}")
    print(f"  Meses no PDF:         {res_meses['total_pdf']:>8}")
    print(f"  Meses OK:             {res_meses['ok']:>8}")
    print()

    if res_meses["faltando"]:
        print(f"  [!] {len(res_meses['faltando'])} mes(es) FALTANDO no PDF:")
        for mes, val in res_meses["faltando"]:
            print(f"      {mes}: {val}")
        print()

    if res_meses["valor_diferente"]:
        print(f"  [!] {len(res_meses['valor_diferente'])} mes(es) com VALOR DIFERENTE:")
        for mes, orig, pdf in res_meses["valor_diferente"]:
            print(f"      {mes}: original={orig} | pdf={pdf}")
        print()

    if not res_meses["faltando"] and not res_meses["valor_diferente"]:
        print("  RESULTADO MESES: Todos os separadores CORRETOS")
    else:
        total_erros = len(res_meses["faltando"]) + len(res_meses["valor_diferente"])
        print(f"  RESULTADO MESES: {total_erros} problema(s) encontrado(s)")

    print("=" * 60)


def main():
    if len(sys.argv) > 1:
        nome_base = sys.argv[1]
    else:
        nome_base = "divergencia2"

    caminho_pdf = f"{nome_base}.pdf"
    caminho_xlsx = f"{nome_base}_unificado.xlsx"
    caminho_original = f"{nome_base}.xlsx"

    if not os.path.isfile(caminho_pdf):
        print(f"ERRO: '{caminho_pdf}' nao encontrado.")
        sys.exit(1)
    if not os.path.isfile(caminho_xlsx):
        print(f"ERRO: '{caminho_xlsx}' nao encontrado.")
        sys.exit(1)

    print(f"Dupla Verificacao:")
    print(f"  PDF:      {caminho_pdf}")
    print(f"  Excel:    {caminho_xlsx}")
    print(f"  Original: {caminho_original}")
    print()

    t0 = time.time()

    pdf_rows = extrair_pdf(caminho_pdf)
    df_excel = carregar_excel(caminho_xlsx)

    res = comparar(df_excel, pdf_rows)
    relatorio(res)

    if os.path.isfile(caminho_original):
        res_meses = verificar_meses(res["pdf_meses"], caminho_original)
        relatorio_meses(res_meses)
    else:
        print(f"\n  [!] Arquivo original '{caminho_original}' nao encontrado, pulando verificacao de meses.")

    print(f"\nConcluido em {time.time() - t0:.1f}s")


if __name__ == "__main__":
    main()
