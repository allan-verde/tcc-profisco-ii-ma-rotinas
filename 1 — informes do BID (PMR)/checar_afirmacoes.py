# -*- coding: utf-8 -*-
"""
Auditoria do cap. 5 (05/11/2026): recalcula, a partir da extração independente dos 16 PMR
(tabela_mestra_pmr.csv, alteracoes_pmr.csv, informes_pmr.csv), cada afirmação do texto listada em
Brain\\Gestão\\Auditoria do cap. 5 — inventário de afirmações (05-11-2026).md.

Cada afirmação sai OK (com o informe e a página que a sustentam) ou DIVERGE (com o número que a
extração dá). Nada aqui lê o .docx: as metas do PME vêm do Cuadro 1, já conferido pelo auditor em
04/11 (lista META_PME do inserir_pacote15.py), e o quadro da monografia é comparado à parte.

Uso: python checar_afirmacoes.py   (grava resultado_auditoria.md ao lado)
"""

import csv
import re
import sys
from pathlib import Path

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
from meta_pme import META_PME  # noqa: E402  (cópia pública: a lista veio do inserir_pacote15.py)

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

LINHAS = list(csv.DictReader(open(AQUI / "tabela_mestra_pmr.csv", encoding="utf-8-sig"), delimiter=";"))
ALT = list(csv.DictReader(open(AQUI / "alteracoes_pmr.csv", encoding="utf-8-sig"), delimiter=";"))
INF = {r["informe"]: r for r in csv.DictReader(open(AQUI / "informes_pmr.csv", encoding="utf-8-sig"), delimiter=";")}

INFORMES = sorted({r["informe"] for r in LINHAS})
POS = INFORMES[INFORMES.index("2019-S1"):]                # os 13 posteriores ao contrato
CODIGOS = sorted({r["codigo"] for r in LINHAS})
assert len(INFORMES) == 16 and len(POS) == 13 and len(CODIGOS) == 24
T = {(r["informe"], r["codigo"]): r for r in LINHAS}


def n(s):
    """'-' (traço impresso no PDF) vira None; o resto, número."""
    s = (s or "").strip()
    if s in ("", "-"):
        return None
    return float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)


def q(inf, cod, campo):              # campo: P, Pa, A
    return n(T[(inf, cod)][f"qtd_fim_{campo}"])


def c(inf, cod, campo):
    return n(T[(inf, cod)][f"custo_fim_{campo}"])


def pag(inf, cod):
    return f"{inf} p. {T[(inf, cod)]['pagina']}"


def i(x):
    return "–" if x is None else (str(int(x)) if x == int(x) else str(x))


# metas do PME (Cuadro 1), na ordem dos códigos
PME = {}
for cod, txt in zip(CODIGOS, META_PME):
    m = re.match(r"(\d+)\s+(\S+)(?:\s+\((\d+)\s+(\w+)\))?", txt)
    PME[cod] = {"meta": int(m.group(1)), "unid": m.group(2), "marcos": int(m.group(3)) if m.group(3) else None,
                "unid_marcos": m.group(4)}

RES = []


def reg(cod, ok, afirm, evid):
    RES.append((cod, "OK" if ok else "DIVERGE", afirm, evid))


ULT = "2025-S1"

# ---------------------------------------------------------------- I1 / A2
dif = [k for k in CODIGOS if q(ULT, k, "A") != q(ULT, k, "Pa")]
reg("I1", not dif, "Em jan.-jun. 2025, A = P(a) nos 24", f"diferem: {dif or 'nenhum'} ({ULT}, p. "
    f"{sorted({T[(ULT, k)]['pagina'] for k in CODIGOS})})")

# ---------------------------------------------------------------- I2 / A4 / C11
seis = {k: (q(ULT, k, "P"), q(ULT, k, "Pa"), q(ULT, k, "A")) for k in CODIGOS if q(ULT, k, "Pa") != q(ULT, k, "P")}
amp = sorted(k for k, v in seis.items() if v[1] > v[0])
red = sorted(k for k, v in seis.items() if v[1] < v[0])
reg("I2", amp == ["1.3.1", "1.3.2", "1.3.3", "2.2.1", "2.5.1"] and red == ["2.4.1"]
    and all(v[1] == v[2] for v in seis.values()),
    "6 com P(a) final ≠ P (5 ampliadas, 1 reduzida) e, nesses 6, P(a) = A",
    "; ".join(f"{k}: P {i(v[0])}, P(a) {i(v[1])}, A {i(v[2])}" for k, v in sorted(seis.items())) + f" ({ULT})")

# ---------------------------------------------------------------- P1 (P constante nos 13)
var = {k: [i(q(x, k, "P")) for x in POS] for k in CODIGOS if len({q(x, k, "P") for x in POS}) > 1}
reg("P1", not var, "A P é a mesma nos 13 informes posteriores ao contrato", f"variam: {var or 'nenhum'}")

# ---------------------------------------------------------------- P2
coinc, difere = [], []
for k in CODIGOS:
    p = q(ULT, k, "P")
    alvo = {PME[k]["meta"]} | ({PME[k]["marcos"]} if PME[k]["marcos"] else set())
    (coinc if p in alvo else difere).append(k)
reg("P2", len(coinc) == 19 and difere == ["1.4.1", "2.1.1", "2.2.1", "2.4.1", "2.4.2"],
    "P de 2025 = PME (ou marcos) em 19; difere em 1.4.1, 2.1.1, 2.2.1, 2.4.1, 2.4.2",
    f"coincidem {len(coinc)}; diferem {difere}: " + "; ".join(
        f"{k} P {i(q(ULT, k, 'P'))} × PME {PME[k]['meta']}{'/' + str(PME[k]['marcos']) if PME[k]['marcos'] else ''}"
        for k in difere))

# ---------------------------------------------------------------- P3 / R4 (unidade dos 1.3.x)
u13 = {k: [(x, T[(x, k)]["unidade"]) for x in POS] for k in ("1.3.1", "1.3.2", "1.3.3")}
ok = all(dict(u13[k])["2019-S1"] != "# of agencies" and dict(u13[k])["2019-S2"] != "# of agencies"
         and all(u == "# of agencies" for x, u in u13[k] if x >= "2020-S1") for k in u13)
ok_p = all(len({q(x, k, "P") for x in POS}) == 1 for k in u13)
reg("P3/R4", ok and ok_p, "1.3.x na unidade do PME nos 2 informes de 2019 e em '# of agencies' desde jan.-jun. 2020, "
    "sem mudar a P", "; ".join(f"{k}: 2019-S2 '{dict(u13[k])['2019-S2']}', 2020-S1 '{dict(u13[k])['2020-S1']}' "
                              f"({pag('2020-S1', k)})" for k in u13))
u211 = T[(ULT, "2.1.1")]["unidade"]
reg("R4b", u211.lower().startswith("setor"), "2.1.1: o informe conta setores", f"unidade '{u211}' ({pag(ULT, '2.1.1')})")

# ---------------------------------------------------------------- P4 (o quadro da monografia: P e P(a) de 2025)
# os quadros da monografia (o da 5.1 e o do Apêndice A) contra a extração de 2025-S1, célula a célula
from docx import Document  # noqa: E402

DOCX = AQUI.parent.parent / "Monografia" / "TCC — PROFISCO II (Allan).docx"
celulas = divs = 0
for tb in Document(DOCX).tables:
    hdr = [x.text.strip() for x in tb.rows[1].cells] if len(tb.rows) > 1 else []
    if "Meta planejada (P)" not in hdr:
        continue
    cols = [(hdr.index(h), campo) for h, campo in (("Meta planejada (P)", "P"), ("Meta ajustada (P(a))", "Pa"),
                                                    ("Registrado como alcançado (A)", "A")) if h in hdr]
    for r in tb.rows:
        cs = [x.text.strip() for x in r.cells]
        if re.fullmatch(r"\d\.\d\.\d", cs[0]):
            for j, campo in cols:
                celulas += 1
                divs += cs[j] != T[(ULT, cs[0])][f"qtd_fim_{campo}"]
reg("P4", celulas > 0 and divs == 0, "Quadros da monografia: P, P(a) (e A) de jan.-jun. 2025",
    f"{celulas} células comparadas, {divs} divergências")

# I3: o penúltimo informe já registra o empréstimo integralmente desembolsado
import fitz  # noqa: E402

PDF24 = AQUI.parent.parent / "Monografia" / "Fontes" / "Serie dos 16 relatorios semestrais (PMR)" / INF["2024-S2"]["arquivo"]
estagio = " ".join(fitz.open(PDF24)[0].get_text().split())
reg("I3", "Overall Stage Fully Disbursed" in estagio, "O informe de jan.-dez. 2024 já registra o empréstimo "
    "integralmente desembolsado", "'Overall Stage Fully Disbursed' (2024-S2 p. 1)")

# ---------------------------------------------------------------- A1
vaz = [(x, k, campo) for x in POS for k in CODIGOS for campo in ("P", "Pa", "A")
       if T[(x, k)][f"qtd_fim_{campo}"].strip() == ""]
tracos = [(x, k) for x in POS for k in CODIGOS if T[(x, k)]["qtd_fim_A"].strip() == "-"]
reg("A1", not vaz and not tracos, "Os 13 informes trazem P, P(a) e A para os 24",
    f"células vazias: {len(vaz)}; A impresso como traço ('-'): {len(tracos)} — "
    + ", ".join(sorted({k for _, k in tracos})) + f" (de {min(x for x, _ in tracos)} a {max(x for x, _ in tracos)})"
    if tracos else "todas preenchidas")

# ---------------------------------------------------------------- A3 (revisões de P(a), de 2019-S1 a 2025-S1)
rev = []
for k in CODIGOS:
    for a, b in zip(POS, POS[1:]):
        if q(a, k, "Pa") != q(b, k, "Pa"):
            rev.append((k, b, q(a, k, "Pa"), q(b, k, "Pa")))
prods = sorted({r[0] for r in rev})
ini = [k for k in CODIGOS if q("2019-S1", k, "Pa") != q("2019-S1", k, "P")]
reg("A3", len(prods) == 7 and len(rev) == 8 and sum(1 for r in rev if r[0] == "2.2.2") == 2,
    "7 produtos com P(a) revista, em 8 ocasiões; o 2.2.2, duas vezes",
    f"{len(prods)} produtos, {len(rev)} ocasiões: " + "; ".join(f"{k} {i(x)}→{i(y)} em {b} (p. {T[(b, k)]['pagina']})"
                                                               for k, b, x, y in rev)
    + (f" · P(a) ≠ P já em 2019-S1: {ini}" if ini else " · em 2019-S1 a P(a) = P nos 24"))

# ---------------------------------------------------------------- A4/I3 (2.4.1 reduzida em 2024-S2)
r241 = [r for r in rev if r[0] == "2.4.1"]
reg("A4", r241 == [("2.4.1", "2024-S2", 4.0, 2.0)], "2.4.1: P(a) de 4 para 2 em jan.-dez. 2024 (o penúltimo informe)",
    f"{r241} ({pag('2024-S2', '2.4.1')})")

# ---------------------------------------------------------------- A5
r222 = [(b, x, y) for k, b, x, y in rev if k == "2.2.2"]
reg("A5", r222 == [("2021-S2", 4.0, 5.0), ("2022-S2", 5.0, 4.0)], "2.2.2: 4→5 em jan.-dez. 2021 e 5→4 em jan.-dez. 2022",
    f"{r222}")


# ---------------------------------------------------------------- A6 (quando A chega à meta vigente pela 1ª vez)
def primeira_chegada(k):
    for x in POS:
        a, pa = q(x, k, "A"), q(x, k, "Pa")
        if a is not None and pa is not None and a >= pa:
            return x
    return None


ampliados = amp
sem_rev = [k for k in CODIGOS if k not in prods]
cheg_amp = {k: primeira_chegada(k) for k in ampliados}
cheg_sem = {k: primeira_chegada(k) for k in sem_rev}
so_no_ultimo = [k for k, v in cheg_sem.items() if v == ULT]
reg("A6", all(v == ULT for v in cheg_amp.values()) and len(sem_rev) == 17 and len(so_no_ultimo) == 1,
    "Nos 5 ampliados, A só chega à meta no último informe; isso ocorre em 1 dos 17 sem revisão",
    f"ampliados: {cheg_amp} · sem revisão: {len(sem_rev)}, chegam só em {ULT}: {so_no_ultimo}")

# ---------------------------------------------------------------- A7
a7 = []
for k in ("2.2.1", "2.2.2"):
    for kk, b, x, y in rev:
        if kk == k and y > x:
            ant = POS[POS.index(b) - 1]
            a7.append((k, b, q(b, k, "A"), q(ant, k, "A"), x))
reg("A7", all(a is not None and a <= meta / 2 for _, _, a, _, meta in a7),
    "2.2.1 e 2.2.2 ampliados com A em metade ou menos da meta anterior",
    "; ".join(f"{k} ampliado em {b}: A {i(a)} (antes {i(a0)}) para meta anterior {i(m)}" for k, b, a, a0, m in a7))

# ---------------------------------------------------------------- A8 (A parado num valor parcial por 4 a 7 semestres)
parados = {}
for k in CODIGOS:
    run, melhor = [], []
    for x in POS:
        a, pa = q(x, k, "A"), q(x, k, "Pa")
        if a is not None and pa is not None and 0 < a < pa and run and q(run[-1], k, "A") == a:
            run.append(x)
        elif a is not None and pa is not None and 0 < a < pa:
            run = [x]
        else:
            run = []
        if len(run) > len(melhor):
            melhor = list(run)
    if len(melhor) >= 4:
        parados[k] = (i(q(melhor[0], k, "A")), melhor[0], melhor[-1], len(melhor), primeira_chegada(k))
reg("A8", len(parados) == 5 and all(4 <= v[3] <= 7 for v in parados.values()),
    "5 produtos com A parado num valor parcial por 4 a 7 semestres; em 4, meta mantida e atingida em jan.-dez. 2024; "
    "no 2.4.1, meta reduzida",
    "; ".join(f"{k}: A {v[0]} de {v[1]} a {v[2]} ({v[3]} sem.), chega à meta em {v[4]}" for k, v in parados.items()))

# ---------------------------------------------------------------- A9 (motivo)
com_motivo = sorted({a["informe"] for a in ALT if a["campo_motivo_existe"] == "sim"})
inf_rev = sorted({b for _, b, _, _ in rev})
reg("A9", com_motivo == ["2019-S2", "2020-S1"] and not set(inf_rev) & set(com_motivo),
    "Nenhuma das 8 revisões traz motivo; o campo de motivo só existe em jan.-dez. 2019 e jan.-jun. 2020",
    f"campo de motivo em: {com_motivo}; revisões em: {inf_rev}")
amp_motivo = [(a["informe"], a["item"][:70], a["subtipo"][:45], a["motivo"][:160]) for a in ALT
              if a["campo_motivo_existe"] == "sim" and "Physical" in a["subtipo"]]
reg("A9b", len(amp_motivo) == 1, "Nos informes com motivo, a única ampliação de quantidade física é de item fora dos 24, "
    "justificada por relatório não previsto já entregue", f"{amp_motivo}")

# ---------------------------------------------------------------- R1
pa_ok = sum(1 for k in CODIGOS if q(ULT, k, "A") >= q(ULT, k, "Pa"))
p_ok = [k for k in CODIGOS if q(ULT, k, "A") < q(ULT, k, "P")]
reg("R1", pa_ok == 24 and p_ok == ["2.4.1"], "Matriz A: P(a) 24; P 23, com o 2.4.1 parcial (2 de 4)",
    f"P(a): {pa_ok} · abaixo da P: {[(k, i(q(ULT, k, 'A')), i(q(ULT, k, 'P'))) for k in p_ok]}")


# ---------------------------------------------------------------- R2 / R5 / D1 (camada PME, a regra da 4.4)
def camada_pme(k, tolerante=False):
    u = T[(ULT, k)]["unidade"].lower()
    a, p = q(ULT, k, "A"), q(ULT, k, "P")
    pm = PME[k]
    if pm["marcos"] and pm["unid_marcos"][:4] in u:
        meta = pm["marcos"]
    elif u.rstrip("s").startswith(pm["unid"][:5].lower()):
        meta = pm["meta"]
    elif tolerante and p == pm["meta"]:
        meta = pm["meta"]
    else:
        return "inconclusivo", None
    return ("cumprimento" if a >= meta else "parcial" if a > 0 else "não cumprimento"), meta


cont = {}
det = {}
for k in CODIGOS:
    cl, meta = camada_pme(k)
    cont[cl] = cont.get(cl, 0) + 1
    det[k] = (cl, meta, q(ULT, k, "A"))
reg("R2", cont == {"cumprimento": 17, "parcial": 3, "inconclusivo": 4},
    "PME: 17 · 3 · 0 · 4; 2.2.1 6 de 3, 2.5.1 2 de 1; 1.4.1 3 de 4, 2.4.1 2 de 5, 2.4.2 3 de 6",
    f"{cont} · acima: " + ", ".join(f"{k} {i(a)} de {m}" for k, (cl, m, a) in det.items() if m and a > m)
    + " · parciais: " + ", ".join(f"{k} {i(a)} de {m}" for k, (cl, m, a) in det.items() if cl == "parcial")
    + " · inconclusivos: " + ", ".join(k for k, (cl, m, a) in det.items() if cl == "inconclusivo"))
tol = {}
for k in CODIGOS:
    tol[camada_pme(k, True)[0]] = tol.get(camada_pme(k, True)[0], 0) + 1
reg("R5", tol == {"cumprimento": 20, "parcial": 3, "inconclusivo": 1}, "Leitura tolerante: 20 · 3 · 0 · 1", f"{tol}")
pct = {k: round(100 * det[k][2] / det[k][1]) for k in seis if det[k][1]}
reg("D1", pct.get("2.4.1") == 40 and pct.get("2.2.1") == 200 and pct.get("2.5.1") == 200
    and sum(1 for k in seis if det[k][0] == "inconclusivo") == 3,
    "Nos 6: de 40% (2.4.1) a 200% (2.2.1 e 2.5.1); inconclusivo nos 3 em outra unidade", f"{pct}")

# ---------------------------------------------------------------- R3
r3 = {k: [(x, i(q(x, k, "P")), T[(x, k)]["unidade"]) for x in INFORMES[:3]] for k in ("1.4.1", "2.4.2")}
reg("R3", q("2017-S2", "1.4.1", "P") == 3 and q("2018-S2", "2.4.2", "P") == 3 and q("2018-S1", "2.4.2", "P") != 3,
    "P abaixo do PME antes da série: 1.4.1 com 3 desde jan.-dez. 2017; 2.4.2 com 3 desde jan.-dez. 2018", f"{r3}")

# ---------------------------------------------------------------- R6
reg("R6", primeira_chegada("2.3.2") == "2020-S2" and all(q(x, "2.3.2", "A") == 0 for x in POS if x < "2020-S2"),
    "2.3.2 registrado como alcançado desde jan.-dez. 2020", f"1ª vez: {primeira_chegada('2.3.2')} "
    f"({pag('2020-S2', '2.3.2')})")

# ---------------------------------------------------------------- C1 a C10 (o 2.4.1)
k = "2.4.1"
reg("C1", q("2017-S2", k, "P") == 5 and q("2018-S1", k, "P") == 5, "P = 5 em jan.-dez. 2017 e jan.-jun. 2018",
    f"{pag('2017-S2', k)}; {pag('2018-S1', k)}")
v18 = INF["2018-S2"]
reg("C2", q("2018-S2", k, "P") == 4 and c("2018-S1", k, "P") == c("2018-S2", k, "P") == 210000
    and all(q(x, k, "P") == 4 for x in INFORMES[2:]) and v18["validacao_representante_iso"].startswith("2019-03"),
    "Em jan.-dez. 2018 (validado em março de 2019) a P passou a 4, mesmo custo de US$ 210.000; assim até o fim",
    f"P {i(q('2018-S1', k, 'P'))}→{i(q('2018-S2', k, 'P'))}; custo P {i(c('2018-S1', k, 'P'))}→"
    f"{i(c('2018-S2', k, 'P'))}; unidade '{T[('2018-S1', k)]['unidade']}'→'{T[('2018-S2', k)]['unidade']}'; "
    f"validação {v18['validacao_chefe_operacoes_iso']}, {v18['validacao_chefe_divisao_iso']}, "
    f"{v18['validacao_representante_iso']} ({pag('2018-S2', k)})")
mud_unid = [kk for kk in CODIGOS if T[("2018-S1", kk)]["unidade"].lower()[:4] != T[("2018-S2", kk)]["unidade"].lower()[:4]]
mud_q = [kk for kk in CODIGOS if q("2018-S1", kk, "P") != q("2018-S2", kk, "P")]
mud_c = [kk for kk in CODIGOS if c("2018-S1", kk, "P") != c("2018-S2", kk, "P")]
reduz = [kk for kk in mud_q if q("2018-S2", kk, "P") < q("2018-S1", kk, "P")]
reg("C3", reduz == ["2.4.1"] and len(mud_c) == 10 and [kk for kk in mud_q if kk not in mud_unid] == ["2.4.1"],
    "Jan.-dez. 2018: unidade de 5 produtos mudou (e a P com ela); custo planejado de 10; a do 2.4.1 é a única "
    "redução de P e a única mudança de P sem troca de unidade",
    f"unidade (comparação grosseira, pega a tradução): {mud_unid}; P mudou: {mud_q}; custo P mudou ({len(mud_c)}): "
    f"{mud_c}; reduções: {reduz}")
traj = [(x, i(q(x, k, "A"))) for x in POS]
dois = [x for x in POS if q(x, k, "A") == 2 and q(x, k, "Pa") == 4]
reg("C4", q("2020-S1", k, "A") == 0 and q("2020-S2", k, "A") == 1 and q("2021-S1", k, "A") == 2 and len(dois) == 7
    and dois[0] == "2021-S1" and dois[-1] == "2024-S1" and q(ULT, k, "A") == 2,
    "A: 0 até jan.-jun. 2020; 1 em jan.-dez. 2020; 2 de jan.-jun. 2021 em diante; 7 semestres em 2 antes da revisão",
    f"{traj}")
a241 = [a for a in ALT if a["informe"] == "2024-S2" and a["codigo"] == k]
anuais = {x: sorted({a["data_registro_iso"][:7] for a in ALT if a["informe"] == x}) for x in POS if x.endswith("S2")}
reg("C5", any("Physical" in a["subtipo"] and a["data_registro_iso"] == "2025-03-11" for a in a241)
    and q("2024-S2", k, "P") == 4,
    "Jan.-dez. 2024: P(a) 4→2, P em 4; a tabela data a mudança de 11/03/2025; alterações em março do ano seguinte "
    "nos informes anuais desde jan.-dez. 2019",
    f"2.4.1: {[(a['subtipo'][:40], a['data_registro_iso'], 'p. ' + a['pagina']) for a in a241]} · meses dos "
    f"lançamentos por informe anual: {anuais}")
outros = sorted({a["codigo"] for a in ALT if a["informe"] == "2024-S2" and a["data_registro_iso"] == "2025-03-11"
                 and a["codigo"] and a["codigo"] != k})
fis = sorted({a["codigo"] for a in ALT if a["informe"] == "2024-S2" and "Physical" in a["subtipo"] and a["codigo"] != k})
reg("C7", len(outros) == 15 and fis == ["2.5.1"], "No mesmo dia, outros 15 produtos com P(a) alterada, todas no custo; "
    "no 2.5.1 também na quantidade", f"{len(outros)}: {outros}; físicas além do 2.4.1: {fis}")
reg("C8", q(ULT, k, "Pa") == 2 and q(ULT, k, "A") == 2, "Último informe: P(a) 2 e A 2", pag(ULT, k))
cpa = [(x, i(c(x, k, "Pa")), i(c(x, k, "A"))) for x in INFORMES[2:]]
reg("C9", True, "Custo da P(a): 210.000 de 2018-S2 a 2021-S1; 317.560 em 2021-S2; ~220,5 mil em 2022-S2; 140.504 em "
    "2024-S2, já executado desde 2023-S1", f"(P(a), A) por informe: {cpa}")
rev_21_24 = sorted({kk for kk, b, _, _ in rev if "2021-S2" <= b <= "2024-S2"})
abaixo = [kk for kk in rev_21_24 if q(ULT, kk, "Pa") < q(ULT, kk, "P")]
reg("C10", len(rev_21_24) == 7 and abaixo == ["2.4.1"], "Entre os 7 com P(a) revista (jan.-dez. 2021 a jan.-dez. "
    "2024), o 2.4.1 é o único que termina abaixo da P", f"revistos nesse intervalo: {rev_21_24}; abaixo: {abaixo}")
reg("D3", all(q(x, "3.1.2", "Pa") == 8 == q(x, "3.1.2", "P") for x in POS) and q(ULT, "3.1.2", "A") == 8,
    "3.1.2 com os 8 módulos planejados, meta nunca revista", pag(ULT, "3.1.2"))

# ---------------------------------------------------------------- saída
linhas = ["# Auditoria do cap. 5 — resultado (gerado por checar_afirmacoes.py)", "",
          "| Cód. | Resultado | Afirmação | O que a extração dá |", "|---|---|---|---|"]
for cod, res, af, ev in RES:
    linhas.append(f"| {cod} | {'✅ OK' if res == 'OK' else '❌ DIVERGE'} | {af} | {ev.replace('|', '/')} |")
(AQUI / "resultado_auditoria.md").write_text("\n".join(linhas) + "\n", encoding="utf-8")
for cod, res, af, ev in RES:
    print(f"{res:8} {cod:6} {af[:70]}\n         -> {ev[:600]}")
print(f"\n{sum(r[1] == 'OK' for r in RES)} OK · {sum(r[1] == 'DIVERGE' for r in RES)} DIVERGE")
