import hashlib
import logging
import re
from collections import defaultdict
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
BASE = RAIZ / "dados_extraidos"
SAIDAS = RAIZ / "saidas"
SAIDAS.mkdir(exist_ok=True)

# Configuração do Logging (Terminal + Arquivo de Log)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(SAIDAS / "execucao.log", encoding="utf-8")
    ]
)

EXT_DADOS = {".csv", ".txt"}
EXT_PLANILHA = {".xlsx", ".xls", ".xlsm", ".ods"}
EXT_DOC = {".pdf", ".docx", ".doc"}
PADRAO_CURSO = r"COMPUTA|SOFTWARE"

def md5(p, n=1 << 20):
    h = hashlib.md5()
    with open(p, "rb") as f:
        while b := f.read(n):
            h.update(b)
    return h.hexdigest()

def formato_real(p):
    with open(p, "rb") as f:
        m = f.read(8)
    if m.startswith(b"PK"):
        return "zip/xlsx"
    if m.startswith(b"\xd0\xcf\x11\xe0"):
        return "ole/xls"
    if m.startswith(b"%PDF"):
        return "pdf"
    if m[:5].lower() in (b"<html", b"<!doc") or m.lstrip().startswith(b"<"):
        return "html"
    return "texto"

def detectar_csv(p):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(p, "r", encoding=enc) as f:
                linha = f.readline()
            sep = max([";", ",", "\t", "|"], key=linha.count)
            return enc, sep
        except UnicodeDecodeError:
            continue
    return "latin-1", ";"

def contar_linhas(p):
    n = 0
    with open(p, "rb") as f:
        for _ in f:
            n += 1
    return max(n - 1, 0)

def ler_planilha(p, fmt):
    engine = "openpyxl" if fmt == "zip/xlsx" else "xlrd" if fmt == "ole/xls" else None
    return pd.read_excel(p, sheet_name=None, header=None, engine=engine)

def main():
    logging.info("=== INICIANDO INSPEÇÃO DE DADOS ===")
    res = []
    dic = []
    P = res.append

    logging.info(f"Mapeando arquivos em: {BASE}")
    arquivos = sorted(p for p in BASE.rglob("*") if p.is_file())
    P(f"BASE: {BASE}")
    P(f"Total de arquivos: {len(arquivos)}")
    
    contagem = defaultdict(int)
    for p in arquivos:
        contagem[p.suffix.lower()] += 1
    P(f"Por extensao: {dict(contagem)}\n")
    logging.info(f"Encontrados {len(arquivos)} arquivos. Distribuição: {dict(contagem)}")

    P("=" * 70 + "\n1. INVENTARIO (pasta | arquivo | MB | formato real)\n" + "=" * 70)
    for p in arquivos:
        rel = p.parent.relative_to(BASE)
        P(f"{rel} | {p.name} | {p.stat().st_size/1e6:.2f} | {formato_real(p)}")

    logging.info("Analisando esquemas de colunas dos arquivos de dados (CSV/TXT)...")
    P("\n" + "=" * 70 + "\n2. ESQUEMAS DE COLUNAS (CSV/TXT agrupados)\n" + "=" * 70)
    esquemas = defaultdict(list)
    info = {}
    for p in arquivos:
        if p.suffix.lower() not in EXT_DADOS:
            continue
        try:
            enc, sep = detectar_csv(p)
            cols = pd.read_csv(p, sep=sep, encoding=enc, nrows=0).columns.tolist()
            info[p] = (enc, sep, cols)
            esquemas[tuple(cols)].append(p)
        except Exception as e:
            logging.error(f"Erro ao ler cabeçalho de {p.name}: {e}")
            P(f"[ERRO] {p.name}: {e}")

    for i, (cols, ps) in enumerate(esquemas.items(), 1):
        enc, sep, _ = info[ps[0]]
        P(f"\n--- Esquema {i}: {len(cols)} colunas | enc={enc} sep={sep!r} | {len(ps)} arquivo(s)")
        for p in ps:
            P(f"    {p.name}  ({p.parent.name})  linhas={contar_linhas(p)}")
        P("    COLUNAS: " + ", ".join(cols))

    base_cols = None
    for cols, ps in esquemas.items():
        if "CO_CURSO" in cols and "QT_ING" in cols:
            base_cols = set(cols)
            break
    if base_cols:
        P("\n--- Diferencas de colunas entre esquemas de CURSOS (vs primeiro)")
        for cols, ps in esquemas.items():
            if "CO_CURSO" in cols and set(cols) != base_cols:
                P(f"  {[p.name for p in ps]}: novas={sorted(set(cols)-base_cols)} faltando={sorted(base_cols-set(cols))}")

    logging.info("Buscando CO_IES da UnB e processando cursos de computação...")
    P("\n" + "=" * 70 + "\n3. UnB E CURSOS DE COMPUTACAO\n" + "=" * 70)
    co_ies = None
    for p, (enc, sep, cols) in info.items():
        if {"CO_IES", "SG_IES", "NO_IES"} <= set(cols):
            d = pd.read_csv(p, sep=sep, encoding=enc, usecols=["CO_IES", "NO_IES", "SG_IES"])
            m = d[d.SG_IES.astype(str).str.upper().eq("UNB")]
            P(f"{p.name}: {m.to_dict('records')}")
            if len(m) and co_ies is None:
                co_ies = int(m.CO_IES.iloc[0])
                logging.info(f"CO_IES da UnB identificado: {co_ies} (via {p.name})")
    
    P(f"CO_IES da UnB usado: {co_ies}")

    linhas = []
    if co_ies is not None:
        for p, (enc, sep, cols) in info.items():
            if not {"CO_IES", "CO_CURSO", "NO_CURSO", "QT_ING"} <= set(cols):
                continue
            
            logging.info(f"Extraindo dados da UnB do arquivo: {p.name}")
            uso = [c for c in ["NU_ANO_CENSO", "CO_IES", "CO_CURSO", "NO_CURSO", "TP_GRAU_ACADEMICO",
                               "TP_MODALIDADE_ENSINO", "NO_MUNICIPIO", "QT_ING", "QT_MAT", "QT_CONC"] if c in cols]
            
            for ch in pd.read_csv(p, sep=sep, encoding=enc, usecols=uso, chunksize=200_000, low_memory=False):
                ch = ch[ch.CO_IES == co_ies]
                ch = ch[ch.NO_CURSO.astype(str).str.contains(PADRAO_CURSO, case=False, regex=True, na=False)]
                if not ch.empty:
                    linhas.append(ch)

    if linhas:
        u = pd.concat(linhas, ignore_index=True)
        out_csv = SAIDAS / "candidatos_unb.csv"
        u.to_csv(out_csv, index=False)
        logging.info(f"Salvos {len(u)} registros de cursos da UnB em {out_csv.name}")
        P(f"\nCandidatos salvos em saidas/candidatos_unb.csv ({len(u)} linhas)")
        P("\nCursos distintos (CO_CURSO | NO_CURSO | anos):")
        g = u.groupby(["CO_CURSO", "NO_CURSO"]).NU_ANO_CENSO.agg(lambda s: sorted(set(s)))
        for (cod, nome), anos in g.items():
            P(f"  {cod} | {nome} | {anos}")
        P("\nUltimo ano disponivel (amostra):")
        ult = u[u.NU_ANO_CENSO == u.NU_ANO_CENSO.max()]
        P(ult.to_string(index=False))
    else:
        logging.warning("Nenhum dado de curso da UnB correspondente ao padrão foi encontrado.")

    logging.info("Verificando planilhas e dicionários de dados...")
    P("\n" + "=" * 70 + "\n4. PLANILHAS/DOCS (dicionarios etc.)\n" + "=" * 70)
    vistos = {}
    for p in arquivos:
        ext = p.suffix.lower()
        if ext in EXT_DOC:
            P(f"[doc] {p.relative_to(BASE)}")
            continue
        if ext not in EXT_PLANILHA:
            continue
        
        h = md5(p)
        if h in vistos:
            logging.info(f"Ignorando planilha duplicada: {p.name}")
            P(f"[duplicado] {p.name} == {vistos[h]}")
            continue
            
        vistos[h] = p.name
        fmt = formato_real(p)
        dic.append("\n" + "#" * 70 + f"\nARQUIVO: {p.name}\nPASTA: {p.parent.relative_to(BASE)}\nFORMATO REAL: {fmt}\n" + "#" * 70)
        
        logging.info(f"Lendo abas da planilha: {p.name}")
        try:
            abas = ler_planilha(p, fmt)
            P(f"[ok] {p.name} | {fmt} | abas={list(abas)}")
            for nome, df in abas.items():
                df = df.dropna(how="all").dropna(axis=1, how="all")
                dic.append(f"\n=== ABA: {nome} | shape={df.shape} ===")
                dic.append(df.to_csv(sep="|", index=False, header=False))
        except Exception as e:
            logging.error(f"Erro ao ler planilha {p.name}: {e}")
            P(f"[ERRO] {p.name} ({fmt}): {e}")
            dic.append(f"ERRO: {e}")

    logging.info("Salvando relatórios em disco...")
    (SAIDAS / "inspecao_resumo.txt").write_text("\n".join(res), encoding="utf-8")
    (SAIDAS / "dicionarios.txt").write_text("\n".join(dic), encoding="utf-8")
    
    logging.info(f"Relatório gerado: {SAIDAS / 'inspecao_resumo.txt'}")
    logging.info(f"Dicionários gerados: {SAIDAS / 'dicionarios.txt'}")
    logging.info("=== INSPEÇÃO FINALIZADA ===")

if __name__ == "__main__":
    main()