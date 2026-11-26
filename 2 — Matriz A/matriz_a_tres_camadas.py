"""Matriz A (execução reportada) nas três camadas — só leitura.

Lê as metas do PME de metas_pme.csv (transcrição do Cuadro 1 do PME, BID 2017a, a mesma do
Apêndice A; desde 26/11/2026 a rotina não lê mais o .docx da monografia) e P, P(a), A e a unidade
de medida direto do PDF do informe final (PMR jan-jun 2025). Classifica cada produto em cada camada
pelas categorias da orientação de 13/10/2026:
cumprimento · parcial · não cumprimento · inconclusivo (unidade incompatível).

Regra de unidade na camada do PME (decisão metodológica, ver 4.4; opção B, 26/11/2026):
- informe na mesma unidade do produto no PME  -> compara com a meta do produto;
- informe em "módulos" e PME com "(n módulos)" -> compara com n;
- outra unidade, com a meta planejada do informe igual à do PME -> compara pelo número;
- outra unidade e outro número                 -> inconclusivo.
Resultado na camada do PME: 20 · 3 · 0 · 1. Com --estrito, toda unidade diferente fica
inconclusiva (a leitura da nota à Tabela 2): 17 · 3 · 0 · 4.

Uso: python matriz_a_tres_camadas.py [--estrito]
"""
import re
import sys
from pathlib import Path

import csv

import fitz

AQUI = Path(__file__).resolve().parent
METAS = AQUI / "metas_pme.csv"
# o informe final do BID (operação BR-L1500, jan-jun 2025); ajuste o caminho se rodar em outra pasta
PMR = Path(r"<pasta dos informes>") / "PMR 2025-S1.pdf"  # ajuste: o informe de jan-jun 2025 da operação BR-L1500
TOLERANTE = "--estrito" not in sys.argv  # padrão = regra da 4.4 (opção B)


def ler_pmr():
    d = fitz.open(PMR)
    t = " | ".join(" | ".join(l.strip() for l in p.get_text().splitlines() if l.strip()) for p in d)
    rx = re.compile(r"(\d\.\d\.\d) ([^|]+?)(?: \| ([^|]+?))? \| ([^|]+?) \| P \| [^|]+ \| (\d+) \|"
                    r".*?P \(a\) \|(?: [^|]+ \|)? (\d+) \|.*?A \|(?: [^|]+ \|)? (\d+) \|")
    out = {}
    for m in rx.finditer(t):
        out.setdefault(m.group(1), (m.group(4).strip(), int(m.group(5)), int(m.group(6)), int(m.group(7))))
    return out


def ler_pme():
    with open(METAS, encoding="utf-8") as f:
        out = {r["codigo"]: (r["unidade"], r["meta_final"]) for r in csv.DictReader(f, delimiter=";")}
    assert len(out) == 24, len(out)
    return out


def classe(a, meta):
    if a >= meta:
        return "cumprimento"
    return "parcial" if a > 0 else "não cumprimento"


def camada_pme(unid_pme, meta_txt, unid_pmr, p):
    meta = int(re.match(r"\d+", meta_txt).group())
    # marcos parciais do PME: "(3 módulos)", "(5 estudos)", "(6 atividades)"
    marcos = re.search(r"\((\d+) (módulo|estudo|atividade)s?\)", meta_txt)
    u_pmr = unid_pmr.lower()
    if marcos and marcos.group(2)[1:5] in u_pmr:   # o informe conta os marcos
        return int(marcos.group(1)), None
    if u_pmr.rstrip("s").startswith(unid_pme.lower()[:5]):
        return meta, None
    if TOLERANTE and p == meta:
        return meta, f"rótulo '{unid_pmr}' no informe; meta planejada igual à do PME"
    return None, f"informe mede em '{unid_pmr}', PME em '{unid_pme}'"


def main():
    pmr, pme = ler_pmr(), ler_pme()
    assert len(pmr) == 24 and len(pme) == 24, (len(pmr), len(pme))
    cont = {k: {} for k in ("PME", "P", "P(a)")}
    for cod in sorted(pme):
        unid_pmr, p, pa, a = pmr[cod]
        unid_pme, meta_txt = pme[cod]
        regua, nota = camada_pme(unid_pme, meta_txt, unid_pmr, p)
        c_pme = "inconclusivo" if regua is None else classe(a, regua)
        linha = {"PME": c_pme, "P": classe(a, p), "P(a)": classe(a, pa)}
        for k, v in linha.items():
            cont[k][v] = cont[k].get(v, 0) + 1
        extra = f"  [{nota}]" if nota else ""
        print(f"{cod}  A={a:<3} PME={meta_txt:<16} -> {c_pme:<14} P={p:<3}-> {linha['P']:<12} "
              f"P(a)={pa:<3}-> {linha['P(a)']}{extra}")
    print()
    for k in ("PME", "P", "P(a)"):
        print(f"camada {k:<5}", cont[k], "total", sum(cont[k].values()))
    assert all(sum(v.values()) == 24 for v in cont.values())


if __name__ == "__main__":
    main()
