from pathlib import Path

import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("Agg")
import matplotlib.pyplot as plt

RAIZ = Path(__file__).resolve().parent.parent
TRAT, SAI = RAIZ / "dados_tratados", RAIZ / "saidas"
H = 5
COORTES = list(range(2015, 2020))
CURSOS = ["Ciência da Computação", "Engenharia de Computação", "Engenharia de Software"]
CORES = dict(zip(CURSOS, ["tab:blue", "tab:green", "tab:red"]))
MET = (("perm", "QT_PERMANENCIA", "Permanência"), ("conc", "conc_acum", "Conclusão acumulada"),
       ("desist", "desist_acum", "Desistência acumulada"))


def wilson(k, n, z=1.96):
    k, n = np.asarray(k, float), np.asarray(n, float)
    p, d = k / n, 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * np.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return 100 * p, 100 * (c - h), 100 * (c + h)


def newcombe(k1, n1, k2, n2):
    p1, l1, u1 = [float(x) / 100 for x in wilson(k1, n1)]
    p2, l2, u2 = [float(x) / 100 for x in wilson(k2, n2)]
    d = p1 - p2
    lo = d - np.sqrt((p1 - l1) ** 2 + (u2 - p2) ** 2)
    hi = d + np.sqrt((u1 - p1) ** 2 + (p2 - l2) ** 2)
    return 100 * d, 100 * lo, 100 * hi


def fmt(p, lo, hi):
    return f"{p:.1f} [{lo:.1f}; {hi:.1f}]".replace(".", ",")


def com_ic(df, prefixo=""):
    df = df.copy()
    for nome, col, _ in MET:
        p, lo, hi = wilson(df[col], df["den"])
        df[f"{nome}_%"], df[f"{nome}_ic_inf"], df[f"{nome}_ic_sup"] = p.round(1), lo.round(1), hi.round(1)
        df[f"{nome}_fmt"] = [fmt(a, b, c) for a, b, c in zip(p, lo, hi)]
    return df


t = pd.read_csv(TRAT / "base_trajetoria.csv").sort_values(["CO_CURSO", "NU_ANO_INGRESSO", "NU_ANO_REFERENCIA"])
t["falec_acum"] = t.groupby(["CO_CURSO", "NU_ANO_INGRESSO"]).QT_FALECIDO.cumsum()
t["den"] = t.QT_INGRESSANTE - t.falec_acum
for nome, col, taxa in (("TAP", "QT_PERMANENCIA", "TAP"), ("TCA", "conc_acum", "TCA"), ("TDA", "desist_acum", "TDA")):
    print(f"Validação {nome} (máx |dif| p.p., denominador ajustado): {(100*t[col]/t.den - t[taxa]).abs().max():.4f}")

c = pd.read_csv(TRAT / "censo_unb_3cursos.csv")
t1 = c[["NU_ANO_CENSO", "curso", "campus", "QT_ING", "QT_MAT", "QT_CONC"]].sort_values(["NU_ANO_CENSO", "curso"])
t1.to_csv(SAI / "tabela1_censo.csv", index=False)

h = t[(t.anos_apos == H) & t.NU_ANO_INGRESSO.isin(COORTES)]
t2 = com_ic(h)[["NU_ANO_INGRESSO", "curso", "campus", "QT_INGRESSANTE", "den"] + [f"{m[0]}_{s}" for m in MET for s in ("%", "ic_inf", "ic_sup", "fmt")]]
t2 = t2.sort_values(["NU_ANO_INGRESSO", "curso"])
t2.to_csv(SAI / "tabela2_trajetoria_h5.csv", index=False)
print("\nTabela 2 (5 anos):\n", t2[["NU_ANO_INGRESSO", "curso", "den", "perm_fmt", "conc_fmt", "desist_fmt"]].to_string(index=False))

soma = ["den", "QT_PERMANENCIA", "conc_acum", "desist_acum"]
camp = h.groupby(["NU_ANO_INGRESSO", "campus"])[soma].sum().reset_index()
t3 = com_ic(camp)
linhas = []
for co, d in t3.groupby("NU_ANO_INGRESSO"):
    f, dr = d[d.campus == "FCTE"].iloc[0], d[d.campus == "Darcy Ribeiro"].iloc[0]
    row = {"coorte": co}
    for nome, col, _ in MET:
        row[f"dif_{nome}_FCTE_menos_Darcy"] = fmt(*newcombe(f[col], f.den, dr[col], dr.den))
    linhas.append(row)
t3.to_csv(SAI / "tabela3_campus_h5.csv", index=False)
pd.DataFrame(linhas).to_csv(SAI / "tabela3b_diferenca_campus_h5.csv", index=False)
print("\nTabela 3 (campus):\n", t3[["NU_ANO_INGRESSO", "campus", "den", "perm_fmt", "conc_fmt", "desist_fmt"]].to_string(index=False))
print("\nDiferença FCTE - Darcy (p.p.) [IC95%]:\n", pd.DataFrame(linhas).to_string(index=False))

agr = h.groupby("curso")[soma].sum().reset_index()
agr_c = h.groupby("campus")[soma].sum().reset_index().rename(columns={"campus": "curso"})
t4 = com_ic(pd.concat([agr, agr_c], ignore_index=True))
t4.to_csv(SAI / "tabela4_agregado_2015_2019.csv", index=False)
print("\nAgregado coortes 2015-2019 (5 anos):\n", t4[["curso", "den", "perm_fmt", "conc_fmt", "desist_fmt"]].to_string(index=False))

s = t[(t.anos_apos == t.NU_PRAZO_INTEGRALIZACAO) & t.NU_ANO_INGRESSO.isin(COORTES)]
t5 = com_ic(s)[["NU_ANO_INGRESSO", "curso", "NU_PRAZO_INTEGRALIZACAO", "den", "perm_fmt", "conc_fmt", "desist_fmt"]]
t5.sort_values(["NU_ANO_INGRESSO", "curso"]).to_csv(SAI / "tabela5_sensibilidade_prazo.csv", index=False)
print("\nSensibilidade (no ano do prazo de integralização):\n", t5.to_string(index=False))

fig, axs = plt.subplots(1, 3, figsize=(15, 4))
for ax, m in zip(axs, ("QT_ING", "QT_MAT", "QT_CONC")):
    for cu, d in t1.groupby("curso"):
        ax.plot(d.NU_ANO_CENSO, d[m], marker="o", label=cu, color=CORES[cu])
    ax.set_title(m)
axs[0].legend(fontsize=8)
fig.tight_layout()
fig.savefig(SAI / "fig1_censo.png", dpi=200)

coortes_all = sorted(t.NU_ANO_INGRESSO.unique())
fig, axs = plt.subplots(3, len(coortes_all), figsize=(3.6 * len(coortes_all), 9), sharey=True)
for j, co in enumerate(coortes_all):
    for i, (m, titulo) in enumerate((("TAP", "Permanência"), ("TCA", "Conclusão acum."), ("TDA", "Desistência acum."))):
        for cu, d in t[t.NU_ANO_INGRESSO == co].groupby("curso"):
            axs[i, j].plot(d.anos_apos, d[m], marker="o", ms=3, label=cu, color=CORES[cu])
        axs[i, j].axvline(H, color="gray", ls=":")
        axs[i, j].set_title(f"{titulo} - {co}", fontsize=9)
        axs[i, j].set_xlabel("Anos após o ingresso")
axs[0, 0].legend(fontsize=7)
fig.tight_layout()
fig.savefig(SAI / "fig2_trajetorias.png", dpi=200)

th = com_ic(h)
fig, axs = plt.subplots(1, 3, figsize=(15, 4.5), sharex=True)
for ax, (nome, _, titulo) in zip(axs, MET):
    for i, cu in enumerate(CURSOS):
        d = th[th.curso == cu].sort_values("NU_ANO_INGRESSO")
        y = d[f"{nome}_%"]
        ax.errorbar(d.NU_ANO_INGRESSO + (i - 1) * 0.15, y, yerr=[y - d[f"{nome}_ic_inf"], d[f"{nome}_ic_sup"] - y],
                    marker="o", capsize=3, label=cu, color=CORES[cu])
    ax.set_title(f"{titulo} em {H} anos (IC 95%)")
    ax.set_xticks(COORTES)
axs[0].legend(fontsize=8)
fig.tight_layout()
fig.savefig(SAI / "fig3_h5_ic.png", dpi=200)
print("\nPronto. Arquivos em", SAI)