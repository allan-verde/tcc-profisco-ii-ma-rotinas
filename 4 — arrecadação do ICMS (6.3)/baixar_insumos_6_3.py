# -*- coding: utf-8 -*-
"""
Baixa os insumos da secao 6.3 (verificacao complementar da arrecadacao do ICMS) e grava CSVs em
Monografia\\Fontes\\. Escrito pelo Bussab em 04/11/2026. So fonte oficial.

O que baixa (cada CSV traz as colunas `fonte` e `acesso_em`):

  1. dca_maranhao_receitas_icms.csv  (2015-2025)
     STN/SICONFI, DCA-Anexo I-C, id_ente=21 (Maranhao), as 4 colunas do anexo, so as linhas que a
     6.3 usa: ICMS (as 3 taxonomias), Adicional FUMACOP, multas/juros e divida ativa do ICMS
     (2015-2017, quando eram contas separadas) e Receitas Correntes.
  2. ipca_numero_indice_sidra_1737.csv
     IBGE/SIDRA, tabela 1737, variavel 2266 (IPCA, numero-indice, dez/1993 = 100), mensal, 2014-2025.
  3. populacao_maranhao_ibge.csv
     IBGE: (a) Estimativas da Populacao, SIDRA 6579 (nao ha 2022 nem 2023);
           (b) Censo 2022, SIDRA 9514;
           (c) Projecoes da Populacao, Revisao 2024 (planilha do FTP do IBGE, tab2), serie anual
               coerente com o Censo 2022.
  4. pib_maranhao_sidra_5938.csv
     IBGE/SIDRA, tabela 5938 (Contas Regionais), PIB a precos correntes, mil reais.
  5. neef_sefazma_icms_fumacop_2014_2017.csv
     SEFAZ-MA/NEEF, planilha "recaudacion del ICMS" anexada pelo BID a operacao BR-L1500
     (EZSHARE-1869036552-10). Usada SO para conferir a troca de conta 2017->2018.

Uso:  C:\\Python314\\python.exe baixar_insumos_6_3.py
"""

import csv
import io
import json
import sys
import time
import urllib.request
from datetime import date
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

AQUI = Path(__file__).resolve().parent
FONTES = AQUI.parent.parent / "Fontes"
HOJE = date.today().isoformat()
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}
ANOS = list(range(2015, 2026))   # 2015-2025 (o DCA 2025 ja estava na API em 04/11/2026)


def baixar(url, tentativas=3):
    for i in range(tentativas):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=120) as r:
                return r.read()
        except Exception as exc:
            print(f"   tentativa {i + 1} falhou: {exc}")
            time.sleep(3)
    raise RuntimeError(f"nao baixou: {url}")


def gravar(nome, cabecalho, linhas):
    caminho = FONTES / nome
    with open(caminho, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cabecalho)
        w.writerows(linhas)
    print(f"   gravado: {caminho} ({len(linhas)} linhas)")


# ---------------------------------------------------------------- 1. SICONFI / DCA
# Contas que a 6.3 usa. A regra: o codigo comeca com um destes prefixos.
CONTAS_DCA = {
    "RO1.0.0.0.00.00.00", "RO1.0.0.0.00.0.0",          # Receitas Correntes (antes / depois de 2018)
    "RO1.1.1.3.00.00.00", "RO1.1.1.3.02.00.00",        # ICMS, taxonomia ate 2017
    "RO1.9.1.1.42.00.00", "RO1.9.1.3.15.00.00",        # multas e juros do ICMS e da DA do ICMS (ate 2017)
    "RO1.9.3.1.15.00.00",                              # divida ativa do ICMS (ate 2017)
    "RO1.1.1.8.02.0.0", "RO1.1.1.8.02.1.0", "RO1.1.1.8.02.2.0",   # 2018-2021
    "RO1.1.1.4.00.0.0",                                            # grupo pai, 2022 em diante
    "RO1.1.1.4.50.0.0", "RO1.1.1.4.50.1.0", "RO1.1.1.4.50.2.0",   # 2022 em diante
}


def baixar_dca():
    print("1. SICONFI / DCA-Anexo I-C")
    linhas = []
    for ano in ANOS:
        url = ("https://apidatalake.tesouro.gov.br/ords/siconfi/tt/dca"
               f"?an_exercicio={ano}&id_ente=21&no_anexo=DCA-Anexo+I-C")
        dados = json.loads(baixar(url).decode("utf-8"))
        if dados.get("hasMore"):
            raise RuntimeError(f"{ano}: a API paginou (hasMore) e este script nao pagina")
        itens = dados.get("items", [])
        # Prova de que nao ha linha de ICMS FORA da lista (conta com 'ICMS' ou 'Circula' no nome
        # que nao esteja em CONTAS_DCA, exceto transferencias da Uniao, grupo 1.7):
        fora = sorted({i["cod_conta"] for i in itens
                       if ("ICMS" in i["conta"] or "Circula" in i["conta"])
                       and i["cod_conta"] not in CONTAS_DCA
                       and not i["cod_conta"].startswith("RO1.7")})
        for i in itens:
            if i["cod_conta"] in CONTAS_DCA:
                linhas.append([ano, i["cod_conta"], i["conta"].strip(), i["coluna"], i["valor"],
                               "STN/SICONFI DCA-Anexo I-C id_ente=21", HOJE])
        print(f"   {ano}: {len(itens)} itens no anexo; linhas de ICMS fora da lista: {fora or 'nenhuma'}")
        time.sleep(0.5)
    gravar("dca_maranhao_receitas_icms.csv",
           ["ano", "cod_conta", "conta", "coluna", "valor", "fonte", "acesso_em"], linhas)


# ---------------------------------------------------------------- 2. IPCA
def baixar_ipca():
    print("2. IBGE/SIDRA 1737 - IPCA numero-indice")
    url = "https://apisidra.ibge.gov.br/values/t/1737/n1/all/v/2266/p/201401-202512"
    dados = json.loads(baixar(url).decode("utf-8"))[1:]
    linhas = [[d["D3C"], d["D3N"], d["V"], "IBGE/SIDRA tabela 1737 v2266 (dez/1993=100)", HOJE]
              for d in dados]
    if len(linhas) != 144:
        raise RuntimeError(f"esperava 144 meses (2014-2025), vieram {len(linhas)}")
    gravar("ipca_numero_indice_sidra_1737.csv",
           ["mes_codigo", "mes", "ipca_indice", "fonte", "acesso_em"], linhas)


# ---------------------------------------------------------------- 3. POPULACAO
def baixar_populacao():
    print("3. IBGE - populacao do Maranhao")
    est = json.loads(baixar("https://apisidra.ibge.gov.br/values/t/6579/n3/21/v/9324/p/all").decode("utf-8"))[1:]
    estimativa = {int(d["D3C"]): int(d["V"]) for d in est}

    cen = json.loads(baixar("https://apisidra.ibge.gov.br/values/t/9514/n3/21/v/93/p/2022"
                            "/c2/6794/c287/100362/c286/113635").decode("utf-8"))[1:]
    censo = {int(d["D3C"]): int(d["V"]) for d in cen}

    url_proj = ("https://ftp.ibge.gov.br/Projecao_da_Populacao/Projecao_da_Populacao_2024/"
                "projecoes_2024_tab2_grupo_quinquenal.xlsx")
    bruto = baixar(url_proj)
    df = pd.read_excel(io.BytesIO(bruto), engine="openpyxl", header=None)
    titulo = str(df.iloc[0, 0])
    if "Revis" not in titulo or "2024" not in titulo:
        raise RuntimeError(f"planilha inesperada: {titulo}")
    cab = df.iloc[5].tolist()
    col_ano = {int(float(v)): j for j, v in enumerate(cab) if str(v).replace(".0", "").isdigit()}
    ma = df[(df[3] == "MA") & (df[1].astype(str).str.strip() == "Ambos")]
    if len(ma) != 19:
        raise RuntimeError(f"esperava 19 grupos etarios do MA, vieram {len(ma)}")
    projecao = {a: int(ma[col_ano[a]].sum()) for a in ANOS}

    linhas = []
    for a in ANOS:
        linhas.append([a, estimativa.get(a, ""), censo.get(a, ""), projecao[a],
                       "IBGE: SIDRA 6579 (estimativa) | SIDRA 9514 (Censo 2022) | "
                       "Projecoes da Populacao Revisao 2024, tab2 (FTP)", HOJE])
    gravar("populacao_maranhao_ibge.csv",
           ["ano", "estimativa_sidra_6579", "censo_2022_sidra_9514", "projecao_revisao_2024",
            "fonte", "acesso_em"], linhas)


# ---------------------------------------------------------------- 4. PIB
def baixar_pib():
    print("4. IBGE/SIDRA 5938 - PIB do Maranhao")
    dados = json.loads(baixar("https://apisidra.ibge.gov.br/values/t/5938/n3/21/v/37/p/all").decode("utf-8"))[1:]
    linhas = [[int(d["D3C"]), d["V"], "IBGE/SIDRA tabela 5938 v37 (PIB precos correntes, mil R$)", HOJE]
              for d in dados if int(d["D3C"]) in ANOS]
    gravar("pib_maranhao_sidra_5938.csv", ["ano", "pib_mil_reais", "fonte", "acesso_em"], linhas)


# ---------------------------------------------------------------- 5. NEEF / SEFAZ-MA
def baixar_neef():
    print("5. SEFAZ-MA/NEEF (planilha anexada pelo BID) - so para conferencia")
    bruto = baixar("https://idbdocs.iadb.org/wsdocs/getdocument.aspx?docnum=EZSHARE-1869036552-10")
    if bruto[:2] != b"PK":
        raise RuntimeError("o BID nao devolveu uma planilha (.xlsx)")
    x = pd.ExcelFile(io.BytesIO(bruto), engine="openpyxl")
    linhas = []
    for aba in ("ICMS", "FUMACOP"):
        df = x.parse(aba, header=None)
        cab = df.iloc[5].tolist()           # "Arrecadacao 2014" ... "Arrecadacao 2017"
        tot = df[df[0].astype(str).str.startswith("Total")].iloc[0].tolist()
        meses = df.iloc[7:19, :]
        nota = " ".join(str(v) for v in df[0].tolist() if "Nota" in str(v))
        for j, v in enumerate(cab):
            if str(v).startswith("Arrecada"):
                ano = int(str(v).split()[-1])
                n_meses = int(meses[j].notna().sum())   # coluna de valores nominais
                linhas.append([ano, aba, tot[j], n_meses, nota.strip()[:120],
                               "SEFAZ-MA/NEEF, Arrecadacao Online (consulta 06.10.17), via BID EZSHARE-1869036552-10",
                               HOJE])
    gravar("neef_sefazma_icms_fumacop_2014_2017.csv",
           ["ano", "tributo", "valor_nominal", "meses_com_dado", "nota_da_planilha", "fonte", "acesso_em"], linhas)


if __name__ == "__main__":
    print(f"Destino: {FONTES}\n")
    baixar_dca()
    baixar_ipca()
    baixar_populacao()
    baixar_pib()
    baixar_neef()
    print("\nPronto.")
