import re
from pathlib import Path

import pandas as pd
from docx import Document
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parent.parent
BASE = RAIZ / "dados_extraidos"
TRATADOS = RAIZ / "dados_tratados"
SAIDAS = RAIZ / "saidas"
TRATADOS.mkdir(exist_ok=True)
SAIDAS.mkdir(exist_ok=True)
CURSOS = {127, 122204, 112891}

docx_p = next(BASE.rglob("*rajet*.docx"))
xlsx_p = next(p for p in BASE.rglob("indicadores_trajetoria*.xlsx") if not p.name.startswith("~$"))

doc = Document(docx_p)
txt = [p.text for p in doc.paragraphs if p.text.strip()]
for t in doc.tables:
    txt.append("\n[TABELA]")
    for r in t.rows:
        txt.append(" | ".join(c.text.strip().replace("\n", " ") for c in r.cells))
(SAIDAS / "dicionario_trajetoria.txt").write_text("\n".join(txt), encoding="utf-8")

rel = [f"XLSX: {xlsx_p}", f"MB: {xlsx_p.stat().st_size/1e6:.1f}"]
wb = load_workbook(xlsx_p, read_only=True)
rel.append(f"Abas: {wb.sheetnames}")
ws = wb[wb.sheetnames[0]]
rel.append(f"Dimensoes (max_row x max_column): {ws.max_row} x {ws.max_column}")
rel.append("Primeiras 12 linhas brutas:")
for i, row in enumerate(ws.iter_rows(values_only=True)):
    rel.append(f"{i}: {list(row)}")
    if i >= 11:
        break
wb.close()

csv_p = TRATADOS / "trajetoria_raw.csv"
if not csv_p.exists():
    df = pd.read_excel(xlsx_p, sheet_name=0, header=None, dtype=str)
    df.to_csv(csv_p, index=False, header=False)
df = pd.read_csv(csv_p, header=None, dtype=str, low_memory=False)
n = df.head(15).notna().sum(axis=1)
h = int(n.idxmax())
df.columns = [str(c).strip() for c in df.iloc[h]]
df = df.iloc[h + 1:].reset_index(drop=True)
df.to_csv(TRATADOS / "trajetoria.csv", index=False)

rel += [f"\nLinha de cabecalho detectada: {h}", f"Shape: {df.shape}", f"Colunas ({len(df.columns)}):"]
rel += [f"  {i}: {c}" for i, c in enumerate(df.columns)]
rel.append("\nPrimeiras 3 linhas:\n" + df.head(3).T.to_string())

c_curso = next((c for c in df.columns if re.search(r"CO_CURSO", c, re.I)), None)
c_ies = next((c for c in df.columns if re.search(r"CO_IES", c, re.I)), None)
rel.append(f"\nColuna curso: {c_curso} | coluna IES: {c_ies}")

if c_curso:
    m = pd.to_numeric(df[c_curso], errors="coerce").isin(CURSOS)
    if c_ies:
        m &= pd.to_numeric(df[c_ies], errors="coerce").eq(2)
    u = df[m]
    u.to_csv(SAIDAS / "trajetoria_unb.csv", index=False)
    rel.append(f"Linhas dos 3 cursos da UnB: {len(u)}")
    for c in df.columns:
        if re.search(r"ANO|COORTE|REFER|INGRESSO", c, re.I):
            rel.append(f"  {c}: {sorted(df[c].dropna().unique())[:30]}")
else:
    nomes = [c for c in df.columns if re.search(r"NO_IES|NOME_IES|IES", c, re.I)]
    rel.append(f"Sem CO_CURSO. Colunas candidatas: {nomes}")

(SAIDAS / "trajetoria_resumo.txt").write_text("\n".join(rel), encoding="utf-8")
print("OK:", list(SAIDAS.glob("*traj*")))