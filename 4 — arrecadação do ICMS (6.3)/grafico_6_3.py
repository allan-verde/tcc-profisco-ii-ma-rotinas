# -*- coding: utf-8 -*-
"""
Gráficos 1 e 2 da seção 6.3, separados em 26/11/2026 por decisão do Allan (opção B), a partir do comentário da
orientadora de 24/11: o gráfico único "continua convidando a leitura que o texto nega", porque a faixa do desembolso
ficava desenhada sobre a linha do ICMS. A versão de dois painéis está em _versões antigas.

  · Gráfico 1: só o ICMS real (R$ bilhões de 2024), 2015–2025, lido de resultado_6_3_ate2025.csv (gerado por
    calcular_6_3.py). Sem marco nenhum da execução.
  · Gráfico 2: em que ano o registrado como alcançado de cada um dos 24 produtos chega ao valor do último informe
    (jan.-jun. 2025), lido de tabela_mestra_pmr.csv (gerada por extrair_pmr.py), 2019–2025. A contagem começa no
    informe de jan.-jun. 2019, o primeiro posterior ao contrato; o traço é lido como ausência de registro. Sob o ano,
    o percentual desembolsado dos três marcos, lido no campo "% Disbursed" da primeira página do informe de jan.-dez.
    daquele ano (MARCOS): o 1º com desembolso (2019, 11,44%), o 1º com mais da metade (2022, 55,97%) e o 1º com 100%
    (2024).

Tons de cinza (o trabalho é impresso) e fonte Arial.
Uso: python grafico_6_3.py   → grava grafico_6_3_icms.png e grafico_6_3_produtos.png nesta pasta
"""

from collections import Counter
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

AQUI = Path(__file__).resolve().parent
MESTRA = next(p for p in (AQUI.parent / "auditoria_pmr" / "tabela_mestra_pmr.csv",
                          AQUI.parent / "1 — informes do BID (PMR)" / "tabela_mestra_pmr.csv") if p.exists())
plt.rcParams.update({"font.family": "Arial", "font.size": 9})
TINTA, CINZA, CLARO = "#1a1a1a", "#555555", "#8c8c8c"


def eixos(eixo):
    eixo.spines[["top", "right"]].set_visible(False)
    eixo.spines[["left", "bottom"]].set_color(CLARO)
    eixo.tick_params(colors=TINTA, length=3)
    eixo.grid(axis="y", color="#dddddd", lw=0.6, zorder=0)


# --- Gráfico 1: a série do ICMS
df = pd.read_csv(AQUI / "resultado_6_3_ate2025.csv")
anos = df["ano"].tolist()
real = (df["icms_B_real_2024_media"] / 1e9).tolist()

fig, ax = plt.subplots(figsize=(15 / 2.54, 7.5 / 2.54), dpi=300)
eixos(ax)
ax.plot(anos, real, color=TINTA, lw=1.8, marker="o", ms=3.5, zorder=3)
for a in (2015, 2025):
    v = real[anos.index(a)]
    ax.annotate(f"{v:.1f}".replace(".", ","), (a, v), textcoords="offset points", xytext=(0, 6), ha="center",
                fontsize=8.5, color=TINTA)
ax.set_ylim(0, 17.5)
ax.set_ylabel("ICMS real\n(R$ bilhões de 2024)", color=TINTA)
ax.set_xticks(anos)
ax.set_xlim(2014.5, 2025.5)
fig.subplots_adjust(left=0.15, right=0.98, top=0.96, bottom=0.12)
saida1 = AQUI / "grafico_6_3_icms.png"
fig.savefig(saida1)
plt.close(fig)

# --- Gráfico 2: em que ano cada produto atinge o valor do último informe
t = pd.read_csv(MESTRA, sep=";", encoding="utf-8-sig")
t["A"] = pd.to_numeric(t["qtd_fim_A"], errors="coerce")
ultimo = sorted(t["informe"].unique())[-1]
final = t[t["informe"] == ultimo].set_index("codigo")["A"]
primeiro = {}
for cod, fim in final.items():
    sub = t[(t["codigo"] == cod) & (t["informe"] >= "2019-S1")].sort_values("informe")
    ok = sub[sub["A"] >= fim]
    primeiro[cod] = int(ok["informe"].iloc[0][:4])
por_ano = Counter(primeiro.values())
assert sum(por_ano.values()) == 24, por_ano
anos2 = list(range(2019, 2026))
assert min(por_ano) >= 2019, por_ano
barras = [por_ano.get(a, 0) for a in anos2]

MARCOS = {2019: "11,44%", 2022: "55,97%", 2024: "100%"}

fig, bx = plt.subplots(figsize=(15 / 2.54, 6.5 / 2.54), dpi=300)
eixos(bx)
bx.bar(anos2, barras, width=0.6, color=CLARO, zorder=3)
for a, n in zip(anos2, barras):
    if n:
        bx.text(a, n + 0.3, str(n), ha="center", va="bottom", fontsize=8, color=TINTA)
bx.set_ylim(0, 13)
bx.set_yticks([0, 5, 10])
bx.set_ylabel("produtos que\natingem o valor\ndo último informe", color=TINTA, fontsize=8.5, linespacing=1.1)
bx.set_xticks(anos2)
bx.set_xticklabels([f"{a}\n({MARCOS[a]})" if a in MARCOS else str(a) for a in anos2])
bx.set_xlim(2018.5, 2025.5)
fig.subplots_adjust(left=0.15, right=0.98, top=0.96, bottom=0.20)
saida2 = AQUI / "grafico_6_3_produtos.png"
fig.savefig(saida2)
plt.close(fig)

print("produtos por ano:", dict(sorted(por_ano.items())))
print("gravados:", saida1, saida2)
