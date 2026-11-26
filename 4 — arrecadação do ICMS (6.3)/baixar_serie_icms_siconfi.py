# -*- coding: utf-8 -*-
"""
Baixa a série anual de arrecadação de ICMS do Maranhão na API do SICONFI (Tesouro Nacional) —
fonte para a seção 6.3 (verificação complementar da arrecadação).

Achado em 15/09/2026: o DCA-Anexo I-C tem a conta "1.1.1.4.50.1.0 - Imposto sobre Operações
Relativas à Circulação de Mercadorias..." — é o ICMS "puro" (sem o Adicional-Fundo de Combate à
Pobreza, que é a conta 1.1.1.4.50.2.0, separada). Confere com reportagem pública: R$ 10,7 bi em
2023 aqui contra R$ 10,9 bi noticiado (a diferença é esse adicional, não itemizado nesta consulta).

Uso:
    python baixar_serie_icms_siconfi.py            # 2015 a 2024 (o último exercício já declarado)

Fonte: https://apidatalake.tesouro.gov.br/ords/siconfi/tt/dca  (dados abertos, sem chave de API)
"""

import csv
import json
import sys
import time
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

ID_ENTE_MA = 21  # código IBGE do Estado do Maranhão
# O código da conta do ICMS "puro" (sem o Adicional-Fundo de Combate à Pobreza) MUDOU de taxonomia
# em algum ponto entre 2019 e 2022 (plano de contas do MCASP revisado) — achado em 15/09/2026 ao ver
# que 2015-2021 vinham vazios com o código novo. As duas formas, tentadas nesta ordem:
CONTAS_ICMS = ["RO1.1.1.4.50.1.0", "RO1.1.1.8.02.1.0", "RO1.1.1.3.02.00.00"]
COLUNA_ALVO = "Receitas Brutas Realizadas"
ANOS = range(2015, 2025)  # 2015-2024; 2025 ainda não tem DCA (declaração fecha após o exercício)

BASE = "https://apidatalake.tesouro.gov.br/ords/siconfi/tt/dca"
SAIDA = Path(__file__).resolve().parent.parent / "Fontes" / "serie_icms_maranhao_siconfi.csv"


def buscar_ano(ano):
    url = f"{BASE}?an_exercicio={ano}&id_ente={ID_ENTE_MA}&no_anexo=DCA-Anexo+I-C"
    with urllib.request.urlopen(url, timeout=30) as resp:
        dados = json.loads(resp.read().decode("utf-8"))
    for conta_icms in CONTAS_ICMS:
        for item in dados.get("items", []):
            if item.get("cod_conta") == conta_icms and item.get("coluna") == COLUNA_ALVO:
                return item["valor"], conta_icms
    return None, None


def main():
    linhas = []
    for ano in ANOS:
        try:
            valor, conta = buscar_ano(ano)
        except Exception as exc:
            print(f"{ano}: ERRO — {exc}")
            valor, conta = None, None
        print(f"{ano}: {valor} (conta {conta})")
        linhas.append((ano, valor, conta))
        time.sleep(0.5)  # cortesia com a API pública

    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    with open(SAIDA, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["exercicio", "icms_receitas_brutas_realizadas_R$", "cod_conta_usado"])
        w.writerows(linhas)
    print(f"\nSalvo em {SAIDA}")


if __name__ == "__main__":
    main()
