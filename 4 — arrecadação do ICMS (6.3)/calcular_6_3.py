# -*- coding: utf-8 -*-
"""
Secao 6.3 - os numeros da arrecadacao do ICMS do Maranhao, 2015-2024 (ou 2015-2025 com --ate 2025). Bussab, 04/11/2026.

PERGUNTA QUE A CONTA RESPONDE: como a arrecadacao do ICMS do Maranhao variou, em valores reais,
por habitante e como parte da receita corrente, de 2015-2018 (antes da execucao do PROFISCO II)
para 2019-2024 (durante; 2019-2025 com --ate 2025)? So descreve. Nao responde se o programa mudou a arrecadacao.

Le (tudo em Monografia\\Fontes\\, baixado por baixar_insumos_6_3.py ou ja existente):
  serie_icms_maranhao_siconfi.csv            serie A (a do script de 15/09: "ICMS puro")
  dca_maranhao_receitas_icms.csv             contas do DCA-Anexo I-C (SICONFI)
  ipca_numero_indice_sidra_1737.csv          IPCA mensal (IBGE)
  populacao_maranhao_ibge.csv                estimativa, Censo 2022 e projecao rev. 2024 (IBGE)
  pib_maranhao_sidra_5938.csv                PIB do MA (IBGE)
  neef_sefazma_icms_fumacop_2014_2017.csv    ICMS e FUMACOP da SEFAZ-MA (conferencia)
  SEFAZ-MA - Relatorios da Administracao Fazendaria\\_RAT2020.txt e _RAT2021.txt (conferencia)
Grava: resultado_6_3.csv (ou resultado_6_3_ate2025.csv) nesta pasta (entrada do grafico).

Uso:  C:\\Python314\\python.exe calcular_6_3.py [--ate 2025]
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

AQUI = Path(__file__).resolve().parent
FONTES = AQUI.parent.parent / "Fontes"
RAT = FONTES / "SEFAZ-MA — Relatórios da Administração Fazendária"
_ap = argparse.ArgumentParser()
_ap.add_argument("--ate", type=int, default=2024, choices=[2024, 2025])
ATE = _ap.parse_args().ate
BASE = 2024                       # ano-base dos valores reais (R$ de 2024), o mesmo nas duas rodadas
ANOS = list(range(2015, ATE + 1))
ANTES, DURANTE = [2015, 2016, 2017, 2018], list(range(2019, ATE + 1))
BRUTA = "Receitas Brutas Realizadas"
DEDUCOES = ["Deduções - FUNDEB", "Deduções - Transferências Constitucionais", "Outras Deduções da Receita"]


def mi(v):   # R$ milhoes, formato brasileiro
    return f"{v / 1e6:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(v, casas=1):
    return f"{v * 100:+.{casas}f}%".replace(".", ",")


def br(v, casas=2):
    return f"{v:,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ------------------------------------------------------------------ leitura
A = pd.read_csv(FONTES / "serie_icms_maranhao_siconfi.csv").set_index("exercicio")
A = A["icms_receitas_brutas_realizadas_R$"]
dca = pd.read_csv(FONTES / "dca_maranhao_receitas_icms.csv")
ipca = pd.read_csv(FONTES / "ipca_numero_indice_sidra_1737.csv", dtype={"mes_codigo": str})
pop = pd.read_csv(FONTES / "populacao_maranhao_ibge.csv").set_index("ano")
pib = pd.read_csv(FONTES / "pib_maranhao_sidra_5938.csv").set_index("ano")["pib_mil_reais"] * 1000
neef = pd.read_csv(FONTES / "neef_sefazma_icms_fumacop_2014_2017.csv")


def conta(ano, cod, coluna=BRUTA):
    s = dca[(dca.ano == ano) & (dca.cod_conta == cod) & (dca.coluna == coluna)]["valor"]
    if len(s) > 1:
        raise RuntimeError(f"{ano} {cod}: linha repetida")
    return float(s.iloc[0]) if len(s) else 0.0


def liquida(ano, cod):
    return conta(ano, cod) - sum(conta(ano, cod, c) for c in DEDUCOES)


# Serie B = ICMS + Adicional FUMACOP + multas/juros + divida ativa do ICMS, montada para ter o
# mesmo conteudo nos 10 anos (ver conferencia abaixo).
PARTES_B = {
    "ate2017": ["RO1.1.1.3.02.00.00", "RO1.9.1.1.42.00.00", "RO1.9.1.3.15.00.00", "RO1.9.3.1.15.00.00"],
    "2018_21": ["RO1.1.1.8.02.0.0"],
    "2022_": ["RO1.1.1.4.50.0.0"],
}


def partes_b(ano):
    return PARTES_B["ate2017"] if ano <= 2017 else PARTES_B["2018_21"] if ano <= 2021 else PARTES_B["2022_"]


def rc_cod(ano):
    return "RO1.0.0.0.00.00.00" if ano <= 2017 else "RO1.0.0.0.00.0.0"


B = pd.Series({a: sum(conta(a, c) for c in partes_b(a)) for a in ANOS})
# Serie A refeita do DCA (mesmos codigos do CSV de 15/09) e conferida contra ele, centavo a centavo:
A_dca = pd.Series({a: conta(a, "RO1.1.1.3.02.00.00" if a <= 2017 else "RO1.1.1.8.02.1.0" if a <= 2021
                         else "RO1.1.1.4.50.1.0") for a in ANOS})
_dif = (A_dca.loc[[a for a in ANOS if a in A.index]] - A).abs().max()
if _dif > 0.01:
    raise RuntimeError(f"serie A do DCA difere do CSV de 15/09 em {_dif}")
print(f"  serie A refeita do DCA bate com o CSV de 15/09 (maior diferenca: {_dif:.2f})")
A = A_dca
B_liq = pd.Series({a: sum(liquida(a, c) for c in partes_b(a)) for a in ANOS})
RC = pd.Series({a: conta(a, rc_cod(a)) for a in ANOS})
RC_liq = pd.Series({a: liquida(a, rc_cod(a)) for a in ANOS})
FCP = pd.Series({a: (conta(a, "RO1.1.1.8.02.2.0") if 2018 <= a <= 2021 else
                     conta(a, "RO1.1.1.4.50.2.0") if a >= 2022 else float("nan")) for a in ANOS})

print("=" * 96)
print("ENTRADAS")
print("=" * 96)
print(f"  pasta das fontes: {FONTES}")
for nome, df in [("dca", dca), ("ipca", ipca), ("pop", pop), ("neef", neef)]:
    print(f"  {nome}: acesso_em = {sorted(set(df['acesso_em']))}")
print("  serie A (15/09): " + ", ".join(f"{a}={mi(v)}" for a, v in A.items()) + "  (R$ mi)")

# ------------------------------------------------------------------ 1. conferencia das contas
print("\n" + "=" * 96)
print("1. CONFERENCIA DAS TROCAS DE CONTA")
print("=" * 96)
print(f"\n1a. O pai e a soma dos filhos? (2018-{ATE}: ICMS + Adicional FUMACOP)")
for a in range(2018, ATE + 1):
    pai, f1, f2 = (("RO1.1.1.8.02.0.0", "RO1.1.1.8.02.1.0", "RO1.1.1.8.02.2.0") if a <= 2021 else
                   ("RO1.1.1.4.50.0.0", "RO1.1.1.4.50.1.0", "RO1.1.1.4.50.2.0"))
    d = conta(a, pai) - conta(a, f1) - conta(a, f2)
    print(f"   {a}: pai {mi(conta(a, pai))} - ICMS {mi(conta(a, f1))} - FUMACOP {mi(conta(a, f2))} = {d:,.2f}")

print("\n1b. 2015-2017: o codigo velho nao tem filho. O pai (1.1.1.3.00) e igual a ele?")
for a in range(2015, 2018):
    print(f"   {a}: 1.1.1.3.00 = {mi(conta(a, 'RO1.1.1.3.00.00.00'))} | 1.1.1.3.02 = {mi(conta(a, 'RO1.1.1.3.02.00.00'))}"
          f" | multas/juros ICMS = {mi(conta(a, 'RO1.9.1.1.42.00.00') + conta(a, 'RO1.9.1.3.15.00.00'))}"
          f" | divida ativa ICMS = {mi(conta(a, 'RO1.9.3.1.15.00.00'))}")
sep = [(a, c) for a in range(2018, ATE + 1) for c in PARTES_B["ate2017"][1:] if conta(a, c)]
print(f"   multas/divida ativa do ICMS como conta SEPARADA a partir de 2018: {sep or 'nenhuma'}")

print("\n1c. O codigo velho (2015-2016) inclui o FUMACOP? Comparacao com a SEFAZ-MA/NEEF")
print("    (NEEF: 'ICMS' = com acrescimos legais e SEM FUMACOP; 'FUMACOP' a parte)")
for a in (2015, 2016):
    n_icms = float(neef[(neef.ano == a) & (neef.tributo == "ICMS")]["valor_nominal"].iloc[0])
    n_fcp = float(neef[(neef.ano == a) & (neef.tributo == "FUMACOP")]["valor_nominal"].iloc[0])
    print(f"   {a}: DCA (1.1.1.3.02 + multas + DA) = {mi(B[a])}")
    print(f"         NEEF ICMS + FUMACOP          = {mi(n_icms + n_fcp)}  -> diferenca {mi(B[a] - n_icms - n_fcp)}"
          f" ({pct(B[a] / (n_icms + n_fcp) - 1, 2)})")
    print(f"         NEEF ICMS sem FUMACOP        = {mi(n_icms)}  -> diferenca {mi(B[a] - n_icms)}"
          f" ({pct(B[a] / n_icms - 1, 2)})")

print("\n1d. 2019-2021: o FUMACOP do DCA bate com o da SEFAZ-MA (RAT)?")
t20 = (RAT / "_RAT2020.txt").read_text(encoding="utf-8", errors="replace")
t21 = (RAT / "_RAT2021.txt").read_text(encoding="utf-8", errors="replace")
num = lambda s: float(s.replace(".", "").replace(",", "."))
m_i = re.search(r"ICMS\s*\n\s*R\$\s*([\d\.]+)\s+R\$\s*([\d\.]+)", t20)
m_f = re.search(r"FUMACOP\s+R\$\s*([\d\.]+)\s+R\$\s*([\d\.]+)", t20)
m21i = re.search(r"ICMS\s*\n([\d\.,]+)\s*\n[\d,]+%\s*\n([\d\.,]+)\s*\n([\d\.,]+)", t21)
m21f = re.search(r"FUMACOP\s*\n([\d\.,]+)\s*\n[\d,]+%\s*\n([\d\.,]+)\s*\n([\d\.,]+)", t21)
rat = {2019: (num(m_i.group(1)), num(m_f.group(1))), 2020: (num(m_i.group(2)), num(m_f.group(2))),
       2021: (num(m21i.group(3)), num(m21f.group(3)))}
for a, (ri, rf) in rat.items():
    print(f"   {a}: FUMACOP DCA {mi(conta(a, 'RO1.1.1.8.02.2.0'))} x RAT {mi(rf)}"
          f" | ICMS DCA {mi(conta(a, 'RO1.1.1.8.02.1.0'))} x RAT {mi(ri)}"
          f" | ICMS+FUMACOP: DCA - RAT = {mi(B[a] - ri - rf)} ({pct(B[a] / (ri + rf) - 1, 2)})")

print("\n1e. Peso do FUMACOP dentro do total (so existe separado de 2018 em diante)")
print("   " + " | ".join(f"{a}: {br(FCP[a] / B[a] * 100, 1)}%" for a in range(2018, ATE + 1)))

print("\n1f. O que a troca faz na serie A (a do CSV de 15/09) x serie B (conteudo constante)")
for a in (2018, 2022):
    print(f"   {a - 1}->{a}: A {pct(A[a] / A[a - 1] - 1)} | B {pct(B[a] / B[a - 1] - 1)}")

# ------------------------------------------------------------------ 2. IPCA
print("\n" + "=" * 96)
print("2. IPCA (IBGE/SIDRA 1737, numero-indice dez/1993 = 100)")
print("=" * 96)
ipca["ano"] = ipca.mes_codigo.str[:4].astype(int)
ipca["mes"] = ipca.mes_codigo.str[4:].astype(int)
cont = ipca.groupby("ano").size()
if (cont.loc[ANOS] != 12).any():
    raise RuntimeError("ano sem os 12 meses")
I_media = ipca.groupby("ano")["ipca_indice"].mean().loc[ANOS]
I_dez = ipca[ipca.mes == 12].set_index("ano")["ipca_indice"].loc[ANOS]
for a in ANOS:
    print(f"   {a}: media dos 12 meses = {br(I_media[a])} | dezembro = {br(I_dez[a])}")
fator_media = I_media[BASE] / I_media
fator_dez = I_dez[BASE] / I_dez

# ------------------------------------------------------------------ 3. tabela principal
B_real = B * fator_media
A_real = A.loc[ANOS] * fator_media
B_real_dez = B * fator_dez
P = pop["projecao_revisao_2024"].loc[ANOS]
pc = B_real / P
pc_est = B_real / pop["estimativa_sidra_6579"].loc[ANOS]
r_rc = B / RC
r_rc_A = A.loc[ANOS] / RC
r_rc_liq = B_liq / RC_liq
r_pib = B / pib.reindex(ANOS)

print("\n" + "=" * 96)
print("3. SERIE B (ICMS + FUMACOP + acrescimos), R$ de 2024 pela MEDIA anual do IPCA")
print("=" * 96)
print(f"{'ano':>5} {'nominal mi':>11} {'fator':>7} {'real mi':>10} {'var real':>9} {'pop proj':>10}"
      f" {'R$/hab':>8} {'ICMS/RC':>8} {'ICMS/PIB':>8}")
for a in ANOS:
    v = "" if a == 2015 else pct(B_real[a] / B_real[a - 1] - 1)
    rp = "" if pd.isna(r_pib[a]) else br(r_pib[a] * 100, 1) + "%"
    print(f"{a:>5} {mi(B[a]):>11} {br(fator_media[a], 4):>7} {mi(B_real[a]):>10} {v:>9} {br(P[a], 0):>10}"
          f" {br(pc[a], 0):>8} {br(r_rc[a] * 100, 1) + '%':>8} {rp:>8}")

# ------------------------------------------------------------------ 4. antes x durante
print("\n" + "=" * 96)
print(f"4. ANTES (2015-2018, n=4) x DURANTE (2019-{ATE}, n={len(DURANTE)}) - so descritivo")
print("=" * 96)


def taxa(s, a0, a1):
    return (s[a1] / s[a0]) ** (1 / (a1 - a0)) - 1


def bloco(nome, s):
    m0, m1 = s.loc[ANTES].mean(), s.loc[DURANTE].mean()
    vy = s.pct_change()
    print(f"  {nome}")
    print(f"     media anual antes {br(m0 / 1e6 if m0 > 1e5 else m0, 1)} | durante {br(m1 / 1e6 if m1 > 1e5 else m1, 1)}"
          f" | diferenca {pct(m1 / m0 - 1)}")
    print(f"     taxa media anual (geometrica): 2015->2018 {pct(taxa(s, 2015, 2018))} | 2019->{ATE} {pct(taxa(s, 2019, ATE))}"
          f" | 2018->{ATE} {pct(taxa(s, 2018, ATE))}")
    print(f"     variacao ano a ano, media simples: antes (16,17,18) {pct(vy.loc[[2016, 2017, 2018]].mean())}"
          f" | durante (19..{ATE % 100}) {pct(vy.loc[DURANTE].mean())}")


bloco("B real (R$ mi de 2024, IPCA media)", B_real)
bloco("B real (R$ mi de 2024, IPCA dezembro)", B_real_dez)
bloco("A real (serie de 15/09, sem conserto)", A_real)
bloco("B real por habitante (R$ de 2024, projecao rev. 2024)", pc)
m0, m1 = r_rc.loc[ANTES].mean(), r_rc.loc[DURANTE].mean()
print(f"  ICMS/receita corrente (bruta/bruta): media antes {br(m0 * 100, 1)}% | durante {br(m1 * 100, 1)}%"
      f" | diferenca {br((m1 - m0) * 100, 1)} p.p.")
m0, m1 = r_rc_liq.loc[ANTES].mean(), r_rc_liq.loc[DURANTE].mean()
print(f"  ICMS/receita corrente (liquida/liquida): media antes {br(m0 * 100, 1)}% | durante {br(m1 * 100, 1)}%")
d = [a for a in DURANTE if not pd.isna(r_pib[a])]
print(f"  ICMS/PIB: media antes {br(r_pib.loc[ANTES].mean() * 100, 2)}% | durante ({d[0]}-{d[-1]}, n={len(d)})"
      f" {br(r_pib.loc[d].mean() * 100, 2)}%  (PIB de 2024 em diante ainda nao publicado na tabela 5938)")

# ------------------------------------------------------------------ 5. escolhas
print("\n" + "=" * 96)
print("5. O QUE CADA ESCOLHA MUDA")
print("=" * 96)
dif = (B_real_dez / B_real - 1)
print(f"  IPCA dezembro x media: o real muda de {pct(dif.min())} a {pct(dif.max())} conforme o ano")
print("  Populacao: estimativa (6579) x projecao rev. 2024, R$/hab real:")
for a in ANOS:
    e = pop["estimativa_sidra_6579"].get(a)
    if pd.notna(e):
        print(f"     {a}: estimativa {br(pc_est[a], 0)} x projecao {br(pc[a], 0)}")
    else:
        print(f"     {a}: estimativa NAO EXISTE ({'Censo 2022 = ' + br(pop['censo_2022_sidra_9514'][a], 0) if pd.notna(pop['censo_2022_sidra_9514'][a]) else 'sem estimativa nem Censo'}) x projecao {br(pc[a], 0)}")
print("  ICMS/RC com a serie A (sem conserto): " + " | ".join(f"{a}: {br(r_rc_A[a] * 100, 1)}%" for a in (2017, 2018, 2021, 2022)))
print("  ICMS/RC liquida/liquida: " + " | ".join(f"{a}: {br(r_rc_liq[a] * 100, 1)}%" for a in ANOS))

# ------------------------------------------------------------------ 6. caso concreto
print("\n" + "=" * 96)
print("6. CASO CONCRETO - 2020 e a variacao real de 2019 para 2020")
print("=" * 96)
a = 2020
print(f"  ICMS 2020 nominal (B) = {br(B[a])}")
print(f"  IPCA media 2020 = {br(I_media[a])} | IPCA media {BASE} = {br(I_media[BASE])}"
      f" -> fator = {br(I_media[BASE])} / {br(I_media[a])} = {br(fator_media[a], 4)}")
print(f"  ICMS 2020 em R$ de 2024 = {br(B[a])} x {br(fator_media[a], 4)} = {br(B_real[a])}")
print(f"  por habitante = {br(B_real[a])} / {br(P[a], 0)} = R$ {br(pc[a])}")
print(f"  sobre a receita corrente = {br(B[a])} / {br(RC[a])} = {br(r_rc[a] * 100, 2)}%")
print(f"  variacao real 2019->2020 = ({br(B[a])} / {br(B[2019])}) / ({br(I_media[a])} / {br(I_media[2019])}) - 1"
      f" = {br(B[a] / B[2019], 4)} / {br(I_media[a] / I_media[2019], 4)} - 1 = {pct(B_real[a] / B_real[2019] - 1, 2)}")

# ------------------------------------------------------------------ saida
out = pd.DataFrame({
    "ano": ANOS, "icms_A_nominal": A.loc[ANOS].values, "icms_B_nominal": B.values,
    "ipca_media": I_media.values, "ipca_dez": I_dez.values,
    "icms_B_real_2024_media": B_real.values, "icms_B_real_2024_dez": B_real_dez.values,
    "icms_A_real_2024_media": A_real.values, "pop_projecao_2024": P.values,
    "icms_B_real_per_capita": pc.values, "receita_corrente_bruta": RC.values,
    "icms_B_sobre_rc": r_rc.values, "icms_B_sobre_pib": r_pib.values,
})
nome_saida = "resultado_6_3.csv" if ATE == 2024 else f"resultado_6_3_ate{ATE}.csv"
out.to_csv(AQUI / nome_saida, index=False, encoding="utf-8")
print(f"\nGravado: {AQUI / nome_saida}")
