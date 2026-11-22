import argparse
import logging
import re
import shutil
import sys
import traceback
import zipfile
from collections import Counter
from pathlib import Path

import pandas as pd
from docx import Document
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parent.parent
BRUTOS, EXTR = RAIZ / "dados_brutos", RAIZ / "dados_extraidos"
TRAT, SAI = RAIZ / "dados_tratados", RAIZ / "saidas"
COD_UNB = 2
CURSOS = {127: ("Ciência da Computação", "Darcy Ribeiro"),
          122204: ("Engenharia de Computação", "Darcy Ribeiro"),
          112891: ("Engenharia de Software", "FCTE")}
PADRAO = r"COMPUTA|SOFTWARE"
H, ANO_FIM = 5, 2024
COORTES_ARTIGO = range(2015, 2020)
REL = []
log = logging.getLogger("insp")


def P(msg=""):
    REL.append(str(msg))
    log.info(msg)


def lixo(p):
    return p.name.startswith(("~$", "md5")) or p.name.lower() == "thumbs.db"


def contar_linhas(p):
    n = 0
    with open(p, "rb") as f:
        while b := f.read(1 << 24):
            n += b.count(b"\n")
    return max(n - 1, 0)


def limpar():
    alvos = [EXTR, TRAT, SAI]
    for a in alvos:
        assert a.resolve().parent == RAIZ.resolve(), f"Caminho inesperado: {a}"
    print("Serão APAGADAS as pastas:")
    for a in alvos:
        print("  ", a)
    print("dados_brutos NÃO será tocada.")
    if input("Digite APAGAR para confirmar: ").strip() != "APAGAR":
        sys.exit("Cancelado.")
    for a in alvos:
        shutil.rmtree(a, ignore_errors=True)
    print("Limpeza concluída.")


def extrair():
    zips = sorted(BRUTOS.glob("*.zip"))
    P(f"== ZIPs em dados_brutos: {len(zips)}")
    bases = Counter(re.sub(r" \(\d+\)$", "", z.stem) for z in zips)
    for z in zips:
        base = re.sub(r" \(\d+\)$", "", z.stem)
        aviso = "  <-- DUPLICATA (mesmo conjunto aparece mais de uma vez)" if bases[base] > 1 else ""
        P(f"  {z.name} ({z.stat().st_size/1e6:.1f} MB){aviso}")
    for z in zips:
        dest = EXTR / z.stem
        if dest.exists() and any(dest.iterdir()):
            continue
        with zipfile.ZipFile(z) as f:
            ruim = f.testzip()
            if ruim:
                P(f"[ERRO] ZIP corrompido: {z.name} ({ruim})")
                continue
            f.extractall(dest)
        log.info("extraído: %s", z.name)


def inventario():
    todos = [p for p in EXTR.rglob("*") if p.is_file()]
    util = [p for p in todos if not lixo(p)]
    P(f"\n== INVENTÁRIO: {len(todos)} arquivos ({len(todos)-len(util)} lixo ignorado: md5, ~$, Thumbs.db)")
    P(f"Por extensão: {dict(Counter(p.suffix.lower() for p in util))}")


def indexar():
    idx, dup = {"CURSOS": {}, "IES": {}}, []
    for p in sorted(EXTR.rglob("*")):
        if not p.is_file() or lixo(p):
            continue
        k = "CURSOS"
        m = re.fullmatch(r"MICRODADOS_CADASTRO_CURSOS_(\d{4})\.csv", p.name, re.I)
        if not m:
            k = "IES"
            m = re.fullmatch(r"MICRODADOS_(?:CADASTRO|ED_SUP)_IES_(\d{4})\.csv", p.name, re.I)
        if m:
            ano = int(m.group(1))
            if ano in idx[k]:
                dup.append(p)
            else:
                idx[k][ano] = p
    return idx, dup


def censo(idx, dup):
    P("\n== CENSO: arquivos e esquemas")
    esperado = set(range(2015, ANO_FIM + 1))
    for k in ("CURSOS", "IES"):
        falta = sorted(esperado - set(idx[k]))
        P(f"{k}: anos {sorted(idx[k])} | faltando: {falta or 'nenhum'}")
    if dup:
        P(f"[AVISO] arquivos duplicados para o mesmo ano: {[str(p.relative_to(EXTR)) for p in dup]}")
    ref = None
    for ano, p in sorted(idx["CURSOS"].items()):
        cols = pd.read_csv(p, sep=";", encoding="latin-1", nrows=0).columns.tolist()
        ref = ref or set(cols)
        novas, aus = sorted(set(cols) - ref), sorted(ref - set(cols))
        P(f"  CURSOS {ano}: {len(cols)} col | {contar_linhas(p)} linhas | novas={len(novas)} {novas[:4]} | ausentes={len(aus)} {aus[:4]}")

    codigos = set()
    for ano, p in sorted(idx["IES"].items()):
        d = pd.read_csv(p, sep=";", encoding="latin-1", usecols=["CO_IES", "SG_IES"])
        codigos |= set(d[d.SG_IES.astype(str).str.upper() == "UNB"].CO_IES)
    P(f"\nCO_IES da UnB (SG_IES=UNB) em todos os anos: {sorted(codigos)} (esperado {{{COD_UNB}}})")

    uso = {"NU_ANO_CENSO", "CO_IES", "CO_CURSO", "NO_CURSO", "TP_GRAU_ACADEMICO", "TP_MODALIDADE_ENSINO",
           "QT_ING", "QT_MAT", "QT_CONC", "QT_SIT_TRANCADA", "QT_SIT_DESVINCULADO", "QT_SIT_TRANSFERIDO"}
    partes = []
    for ano, p in sorted(idx["CURSOS"].items()):
        for ch in pd.read_csv(p, sep=";", encoding="latin-1", usecols=lambda c: c in uso,
                              chunksize=200_000, low_memory=False):
            ch = ch[ch.CO_IES == COD_UNB]
            partes.append(ch[ch.NO_CURSO.astype(str).str.contains(PADRAO, case=False, na=False)])
    cand = pd.concat(partes, ignore_index=True)
    cand.to_csv(SAI / "candidatos_unb.csv", index=False)
    P("\nCandidatos (CO_CURSO | nome | grau | anos):")
    for cod, d in cand.groupby("CO_CURSO"):
        graus = [int(x) for x in sorted(d.TP_GRAU_ACADEMICO.dropna().unique())]
        P(f"  {cod} | {d.NO_CURSO.iloc[-1]} | grau={graus} | {d.NU_ANO_CENSO.min()}-{d.NU_ANO_CENSO.max()} ({d.NU_ANO_CENSO.nunique()} anos)")

    c3 = cand[cand.CO_CURSO.isin(CURSOS)].copy()
    c3["curso"] = c3.CO_CURSO.map(lambda x: CURSOS[x][0])
    c3["campus"] = c3.CO_CURSO.map(lambda x: CURSOS[x][1])
    c3.to_csv(TRAT / "censo_unb_3cursos.csv", index=False)
    for m in ("QT_ING", "QT_MAT", "QT_CONC"):
        P(f"\n{m} por ano e curso:\n" + c3.pivot_table(index="NU_ANO_CENSO", columns="curso", values=m, aggfunc="sum").to_string())


def ler_traj(xlsx):
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    cols, rows, titulo, total, coortes = None, [], "", 0, set()
    i_ies = i_cur = i_ing = None
    for row in ws.iter_rows(values_only=True):
        if cols is None:
            if row and row[0] == "CO_IES":
                cols = list(row)
                i_ies, i_cur, i_ing = cols.index("CO_IES"), cols.index("CO_CURSO"), cols.index("NU_ANO_INGRESSO")
            elif row and isinstance(row[0], str) and "oorte" in row[0]:
                titulo = row[0][:160]
            continue
        try:
            if row[i_ies] is None or row[i_ing] is None:
                continue
            total += 1
            coortes.add(int(row[i_ing]))
            if int(row[i_ies]) == COD_UNB and int(row[i_cur]) in CURSOS:
                rows.append(row)
        except (TypeError, ValueError, IndexError):
            continue
    wb.close()
    if cols is None:
        return pd.DataFrame(), dict(titulo="CABEÇALHO NÃO ENCONTRADO", total=0, coortes=set())
    return pd.DataFrame(rows, columns=cols), dict(titulo=titulo, total=total, coortes=coortes)


def trajetoria():
    P("\n== TRAJETÓRIA")
    arqs = sorted(p for p in EXTR.rglob("*.xlsx") if "traj" in p.name.lower() and not lixo(p))
    P(f"Arquivos .xlsx de trajetória: {len(arqs)}")
    frames = []
    for p in arqs:
        log.info("lendo %s (pode demorar)", p.name)
        df, meta = ler_traj(p)
        P(f"  {p.name} | {p.stat().st_size/1e6:.1f} MB | linhas={meta['total']} | coortes={sorted(meta['coortes'])} | linhas UnB(3 cursos)={len(df)}")
        P(f"     título: {meta['titulo']}")
        frames.append(df)
    frames = [f for f in frames if len(f)]
    if not frames:
        P("[ERRO] nenhuma linha da UnB encontrada nos arquivos de trajetória")
        return
    t = pd.concat(frames, ignore_index=True)
    num = ["CO_CURSO", "NU_ANO_INGRESSO", "NU_ANO_REFERENCIA", "NU_PRAZO_INTEGRALIZACAO", "NU_PRAZO_ACOMPANHAMENTO",
           "NU_ANO_MAXIMO_ACOMPANHAMENTO", "QT_INGRESSANTE", "QT_PERMANENCIA", "QT_CONCLUINTE",
           "QT_DESISTENCIA", "QT_FALECIDO", "TAP", "TCA", "TDA", "TCAN", "TADA"]
    t[num] = t[num].apply(pd.to_numeric, errors="coerce")
    antes = len(t)
    t = t.drop_duplicates(["CO_CURSO", "NU_ANO_INGRESSO", "NU_ANO_REFERENCIA"])
    P(f"\nLinhas UnB: {antes} -> {len(t)} após remover duplicatas")
    coortes = sorted(int(c) for c in t.NU_ANO_INGRESSO.unique())
    fora = [c for c in coortes if c not in COORTES_ARTIGO]
    P(f"Coortes na base: {coortes} | fora do recorte do artigo (2015-2019): {fora or 'nenhuma'}")
    t["curso"] = t.CO_CURSO.map(lambda c: CURSOS[int(c)][0])
    t["campus"] = t.CO_CURSO.map(lambda c: CURSOS[int(c)][1])
    t["anos_apos"] = (t.NU_ANO_REFERENCIA - t.NU_ANO_INGRESSO).astype(int)
    t = t.sort_values(["CO_CURSO", "NU_ANO_INGRESSO", "NU_ANO_REFERENCIA"]).reset_index(drop=True)
    g = t.groupby(["CO_CURSO", "NU_ANO_INGRESSO"])
    t["conc_acum"] = g.QT_CONCLUINTE.cumsum()
    t["desist_acum"] = g.QT_DESISTENCIA.cumsum()
    P(f"Validação (máx |diferença| em p.p.): TDA={(100*t.desist_acum/t.QT_INGRESSANTE-t.TDA).abs().max():.4f} "
      f"TCA={(100*t.conc_acum/t.QT_INGRESSANTE-t.TCA).abs().max():.4f}")
    t.to_csv(TRAT / "base_trajetoria.csv", index=False)

    P("\nCompletude (anos de referência por coorte x curso; esperado = anos de ingresso até %d):" % ANO_FIM)
    comp = t.groupby(["NU_ANO_INGRESSO", "curso"]).NU_ANO_REFERENCIA.nunique().unstack()
    comp["esperado"] = ANO_FIM - comp.index.astype(int) + 1
    P(comp.to_string())
    P("\nIngressantes por coorte x curso:\n" + t.drop_duplicates(["curso", "NU_ANO_INGRESSO"]).pivot(
        index="NU_ANO_INGRESSO", columns="curso", values="QT_INGRESSANTE").to_string())
    P("\nPrazos (integralização / acompanhamento) por curso e coorte:\n" + t.drop_duplicates(
        ["curso", "NU_ANO_INGRESSO"])[["NU_ANO_INGRESSO", "curso", "NU_PRAZO_INTEGRALIZACAO",
                                        "NU_PRAZO_ACOMPANHAMENTO"]].to_string(index=False))
    h = t[t.anos_apos == H][["NU_ANO_INGRESSO", "curso", "QT_INGRESSANTE", "TAP", "TCA", "TDA"]].round(1)
    P(f"\nIndicadores {H} anos após o ingresso:\n" + h.to_string(index=False))


def dicionario():
    docs = [p for p in EXTR.rglob("*.docx") if "traj" in p.name.lower() and not lixo(p)]
    if not docs:
        P("\n[AVISO] dicionário .docx de trajetória não encontrado")
        return
    d = Document(docs[0])
    txt = [p.text for p in d.paragraphs if p.text.strip()]
    for tb in d.tables:
        for r in tb.rows:
            txt.append(" | ".join(c.text.strip().replace("\n", " ") for c in r.cells))
    (SAI / "dicionario_trajetoria.txt").write_text("\n".join(txt), encoding="utf-8")
    P(f"\nDicionário salvo em saidas/dicionario_trajetoria.txt (fonte: {docs[0].name})")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limpar", action="store_true")
    ap.add_argument("--extrair", action="store_true")
    a = ap.parse_args()
    if a.limpar:
        limpar()
    for d in (EXTR, TRAT, SAI):
        d.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s",
                        handlers=[logging.FileHandler(SAI / "execucao.log", encoding="utf-8"), logging.StreamHandler()])
    try:
        P(f"RAIZ: {RAIZ}")
        if a.extrair:
            extrair()
        inventario()
        idx, dup = indexar()
        censo(idx, dup)
        trajetoria()
        dicionario()
    except Exception:
        log.exception("Falha")
        P("[FALHA]\n" + traceback.format_exc())
    finally:
        (SAI / "inspecao_resumo.txt").write_text("\n".join(REL), encoding="utf-8")
        print("Relatório:", SAI / "inspecao_resumo.txt")


if __name__ == "__main__":
    main()