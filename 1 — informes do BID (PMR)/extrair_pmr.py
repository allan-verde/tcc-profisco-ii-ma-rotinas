# -*- coding: utf-8 -*-
r"""
extrair_pmr.py
==============

Extracao INDEPENDENTE dos 16 informes semestrais de monitoramento (PMR, Project
Monitoring Report) do BID para a operacao BR-L1500 (PROFISCO II Maranhao,
contrato de emprestimo 4458/OC-BR).

O script foi escrito para servir de prova: qualquer pessoa pode roda-lo de novo,
do zero, e obter os mesmos arquivos. Ele NAO interpreta nada (nao diz se o
programa cumpriu ou nao); so le o que esta impresso e prova a leitura.

O que ele faz, em ordem
-----------------------
1. Le cada PDF da pasta de fontes. Os PDFs sao so leitura: nada e gravado la.
   Guarda o SHA-256 de cada PDF, para provar que a entrada e a mesma.
2. Da pagina 1, tira o ciclo do informe (ex.: "Second period Jan-Dec 2017"),
   o estagio de validacao e as datas de validacao impressas.
3. CAMINHO 1 (o principal): le a matriz de produtos ("OUTPUTS: ANNUAL PHYSICAL
   AND FINANCIAL PROGRESS") pelas LINHAS DE GRADE desenhadas no PDF
   (PyMuPDF `find_tables`, estrategia "lines"). Cada celula sai da grade.
4. CAMINHO 2 (conferencia completa): le a mesma matriz de novo, agora pelas
   COORDENADAS (x, y) de cada palavra, sem usar a grade. As colunas sao achadas
   pela posicao do cabecalho ("2019", "EOP 2024"...). As duas leituras sao
   comparadas celula a celula, para TODOS os produtos de TODOS os informes.
5. CAMINHO 3 (conferencia amostral): sorteia 3 produtos por informe (semente
   fixa, sempre os mesmos) e confere contra o TEXTO BRUTO da pagina
   (`get_text("text")`, a ordem em que o PDF guarda o texto). Gera tambem um
   PNG recortado da pagina com esses 3 produtos, para conferencia visual.
6. Le a tabela "CHANGES TO THE MATRIX" (alteracoes da matriz de resultados).
7. Grava, na pasta de saida:
     tabela_mestra_pmr.csv   uma linha por informe x produto
     alteracoes_pmr.csv      uma linha por alteracao registrada
     informes_pmr.csv        uma linha por informe (datas, layout, contagens)
     conferencia_amostral.md o relatorio das conferencias
     conferencia_png\        os recortes de pagina usados na conferencia visual
   CSV em UTF-8 com BOM, separador ";" (abre direto no Excel em portugues).

Regras de leitura (nunca adivinhar)
-----------------------------------
* Celula sem nada impresso fica VAZIA no CSV e o motivo vai para a coluna obs.
* Um traco impresso ("-") e gravado como "-": foi lido com seguranca e nao e a
  mesma coisa que zero, nem que celula vazia.
* Numero e gravado sem separador de milhar (390,000 no PDF -> 390000 no CSV).
* Produto ausente num informe e registrado como achado, nunca preenchido.
* O codigo contratual de 3 niveis (1.1.1 ... 3.2.2) so e preenchido quando esta
  impresso no nome do produto. Nos informes que imprimem so o codigo de 2
  niveis ("1.1"), ele e atribuido pela correspondencia codigo-interno-do-PMR <->
  codigo-do-nome observada nos informes que imprimem os dois, e a coluna obs
  diz isso.

Uso
---
    python extrair_pmr.py
    python extrair_pmr.py --fontes "<pasta dos PDFs>" --saida "<pasta de saida>"

Dependencia: PyMuPDF (fitz) >= 1.23, por causa do find_tables. Testado com 1.28.0.
"""

import argparse
import csv
import glob
import hashlib
import os
import random
import re
import sys

import fitz  # PyMuPDF

PASTA_FONTES = r"H:\Meu Drive\30-Estudy\UFMA\Monografia\Monografia\Fontes\Serie dos 16 relatorios semestrais (PMR)"
PASTA_SAIDA = os.path.dirname(os.path.abspath(__file__))

# Os 24 codigos do contrato (componente.subcomponente.produto). O script NAO
# confia nesta lista: ele compara com o conjunto que efetivamente le nos
# informes e avisa se diferir.
CODIGOS_CONTRATO = [
    "1.1.1", "1.1.2", "1.2.1", "1.2.2", "1.3.1", "1.3.2", "1.3.3", "1.4.1",
    "2.1.1", "2.2.1", "2.2.2", "2.3.1", "2.3.2", "2.3.3", "2.4.1", "2.4.2",
    "2.5.1", "2.5.2", "2.6.1", "2.6.2", "3.1.1", "3.1.2", "3.2.1", "3.2.2",
]

SEMENTE_AMOSTRA = "auditoria-pmr-2026"   # semente fixa do sorteio dos 3 produtos

# Registro da inspecao VISUAL dos PNGs (caminho 3, segunda metade). Ela foi
# feita a olho por quem rodou o script, e por isso fica escrita aqui, como
# constante: rodar o script de novo reproduz o mesmo relatorio. Chave = informe.
_OK_VISUAL = ("os 3 produtos conferem: unidade e as 6 celulas de fim de projeto (qtd e custo de P, P(a) e A) "
              "da imagem sao iguais as do CSV; colunas do ano tambem conferidas")
CONFERENCIA_VISUAL = {
    "2017-S2": "1.3.3 (impresso 1.7), 2.2.2 (2.3), 2.3.1 (2.4): " + _OK_VISUAL.replace("; colunas do ano tambem conferidas", "; este informe nao tem coluna do ano"),
    "2018-S1": "2.3.3 (2.6), 2.4.1 (2.7), 3.1.2 (3.2): " + _OK_VISUAL + "; as celulas do ano de P(a) e A estao em branco na imagem, como no CSV",
    "2018-S2": "1.2.2 (1.4), 2.4.1 (2.7), 3.1.2 (3.2): " + _OK_VISUAL.replace("; colunas do ano tambem conferidas", "; este informe nao tem coluna do ano"),
    "2019-S1": "2.4.1, 2.6.2, 3.1.1: " + _OK_VISUAL,
    "2019-S2": "1.2.2, 2.2.1, 2.2.2: " + _OK_VISUAL,
    "2020-S1": "1.3.3, 1.4.1, 2.3.2: " + _OK_VISUAL + "; unidade de 1.3.3 impressa como '# of agencies'",
    "2020-S2": "1.3.3, 2.2.2, 2.6.2: " + _OK_VISUAL,
    "2021-S1": "1.4.1, 2.2.1, 2.3.3: " + _OK_VISUAL + "; o A de custo de 2.2.1 esta impresso com centavos (1,579,861.38)",
    "2021-S2": "1.3.2, 2.3.2, 2.6.1: " + _OK_VISUAL + "; os tracos ('-') da imagem estao como '-' no CSV",
    "2022-S1": "1.2.1, 2.3.3, 2.6.2: " + _OK_VISUAL + "; tracos conferidos",
    "2022-S2": "1.1.1, 2.3.1, 2.5.2: " + _OK_VISUAL + "; tracos conferidos",
    "2023-S1": "1.1.2, 2.2.1, 2.4.2: " + _OK_VISUAL + "; tracos conferidos",
    "2023-S2": "1.3.3, 2.4.2, 3.2.1: " + _OK_VISUAL + "; tracos conferidos",
    "2024-S1": "2.2.1, 2.6.2, 3.2.2: " + _OK_VISUAL + "; tracos conferidos",
    "2024-S2": "1.3.3, 2.3.1, 2.5.2: " + _OK_VISUAL + "; o custo de 2.5.2 esta impresso com centavos (2,826.16)",
    "2025-S1": "1.2.2, 1.3.2, 3.1.1: " + _OK_VISUAL + "; tracos conferidos",
}
CONFERENCIA_VISUAL_DATA = "05/11/2026"

# ---------------------------------------------------------------------------
# Expressoes regulares
# ---------------------------------------------------------------------------
# codigo impresso no inicio do nome do produto: "1.1.1 Sistemas..." ou "1.1 Sistemas..."
RE_COD_NOME = re.compile(r"^(\d)\.(\d{1,2})(?:\.(\d{1,2}))?\s+(.+)$", re.S)
# codigo interno do PMR, na coluna da esquerda: "1.1", "2.10", "1.01"
RE_COD_PMR = re.compile(r"^(\d)\.(\d{1,2})$")
# numero no padrao americano do BID: 390,000 / 1,820,582 / 0 / 0.59 / (20.51)
RE_NUM = re.compile(r"^\(?-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\)?$")
RE_DATA = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})$")

ROTULOS = {"P": "P", "P(a)": "Pa", "P (a)": "Pa", "A": "A"}
ORDEM_LINHAS = ["P", "Pa", "A"]

# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def limpar(s):
    """Junta quebras de linha e espacos repetidos. None continua None."""
    if s is None:
        return None
    return re.sub(r"\s+", " ", s.replace("\n", " ")).strip()


def ler_valor(bruto):
    """Converte o texto de uma celula de valor.

    Devolve (valor_para_o_csv, problema). problema e None quando a leitura e
    segura; senao, e o motivo pelo qual a celula ficou vazia.
    """
    if bruto is None:
        return "", "celula ausente na grade do PDF"
    s = limpar(bruto)
    if s == "":
        return "", "celula vazia (nada impresso)"
    if s in ("-", "\u2013"):
        return "-", None
    if RE_NUM.match(s):
        negativo = s.startswith("(") and s.endswith(")")
        t = s.strip("()").replace(",", "")
        if negativo or t.startswith("-"):
            t = "-" + t.lstrip("-")
        if "." in t:
            # decimal: grava com virgula, como o Excel em portugues espera
            return t.replace(".", ","), None
        return str(int(t)), None
    return "", "conteudo nao reconhecido como numero: %r" % s


def data_iso(texto):
    """Datas do PMR sao mes/dia/ano (padrao americano). Devolve AAAA-MM-DD."""
    if not texto:
        return ""
    m = RE_DATA.match(texto)
    if not m:
        return ""
    mes, dia, ano = int(m.group(1)), int(m.group(2)), m.group(3)
    if mes > 12:
        raise ValueError("data %r nao esta no formato mes/dia/ano" % texto)
    ano = int(ano) + (2000 if len(ano) == 2 else 0)
    return "%04d-%02d-%02d" % (ano, mes, dia)


def chave_codigo(cod):
    """'1.01' -> (1, 1); '2.10' -> (2, 10); '1.1.1' -> (1, 1, 1)."""
    return tuple(int(p) for p in cod.split("."))


def sha256(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def palavras(page):
    """Lista de palavras (x0, y0, x1, y1, texto) da pagina."""
    return [(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words")]


def achar_sequencia(ws, seq):
    """Acha as palavras `seq` em sequencia numa mesma linha.

    Devolve (x0, x1, y0) do rotulo inteiro, ou None.
    """
    for w in sorted(ws, key=lambda v: (v[1], v[0])):
        if w[4] != seq[0]:
            continue
        linha = sorted([v for v in ws if abs(v[1] - w[1]) < 2.0 and v[0] >= w[0]],
                       key=lambda v: v[0])
        if [v[4] for v in linha[:len(seq)]] == seq:
            return (w[0], linha[len(seq) - 1][2], w[1])
    return None


def a_direita(ws, rotulo, limite_x, so_datas=False):
    """Palavras na mesma linha do rotulo, a direita dele e antes de limite_x."""
    if rotulo is None:
        return []
    x0, x1, y = rotulo
    saida = sorted([v for v in ws if abs(v[1] - y) < 2.5 and v[0] > x1 and v[0] < limite_x],
                   key=lambda v: v[0])
    if so_datas:
        saida = [v for v in saida if RE_DATA.match(v[4])]
    return [v[4] for v in saida]


# ---------------------------------------------------------------------------
# Pagina 1: identificacao do informe e datas de validacao
# ---------------------------------------------------------------------------


def ler_cabecalho(page):
    ws = palavras(page)
    # titulo = a primeira linha da pagina ("PMR Operational Report" / "PMR Public Report")
    y_topo = min(w[1] for w in ws)
    titulo = " ".join(w[4] for w in sorted([w for w in ws if w[1] < y_topo + 3], key=lambda w: w[0]))
    r_chefe_op = achar_sequencia(ws, ["Chief", "of", "Operations"])
    # tudo que esta a esquerda do bloco "Chief of Operations" e o painel da esquerda
    limite_esq = r_chefe_op[0] if r_chefe_op else page.rect.width
    fim = page.rect.width + 1

    ciclo = " ".join(a_direita(ws, achar_sequencia(ws, ["Year-", "PMR", "Cycle"]), limite_esq))
    operacao = " ".join(a_direita(ws, achar_sequencia(ws, ["Operation", "Number"]), limite_esq))
    estagio = " ".join(a_direita(ws, achar_sequencia(ws, ["PMR", "Validation", "Stage"]), limite_esq))
    ultima = a_direita(ws, achar_sequencia(ws, ["Last", "Update"]), limite_esq, so_datas=True)
    chefe_op = a_direita(ws, r_chefe_op, fim, so_datas=True)
    chefe_div = a_direita(ws, achar_sequencia(ws, ["Division", "Chief"]), fim, so_datas=True)
    repres = a_direita(ws, achar_sequencia(ws, ["Country", "Representative"]), fim, so_datas=True)

    m = re.match(r"^(First|Second) period (Jan-Jun|Jan-Dec) (\d{4})$", ciclo)
    if not m:
        raise ValueError("ciclo do PMR nao reconhecido: %r" % ciclo)
    sem = "S1" if m.group(1) == "First" else "S2"
    periodo = ("jan-jun " if m.group(2) == "Jan-Jun" else "jan-dez ") + m.group(3)

    def um(lista):
        return lista[0] if lista else ""

    return {
        "titulo": titulo,
        "operacao": operacao,
        "ciclo_impresso": ciclo,
        "informe": "%s-%s" % (m.group(3), sem),
        "periodo": periodo,
        "estagio_validacao": estagio,
        "ultima_atualizacao": um(ultima),
        "validacao_chefe_operacoes": um(chefe_op),
        "validacao_chefe_divisao": um(chefe_div),
        "validacao_representante": um(repres),
    }


# ---------------------------------------------------------------------------
# CAMINHO 1: matriz de produtos lida pela grade do PDF
# ---------------------------------------------------------------------------


def paginas_da_matriz(doc):
    return [i for i, p in enumerate(doc)
            if "Component Nbr." in p.get_text() and "Unit of Measure" in p.get_text()]


def ler_matriz_grade(page):
    """Devolve (produtos, itens_que_nao_sao_produto, avisos, colunas)."""
    produtos, outros, avisos = [], [], []
    colunas = None
    tabs = page.find_tables(strategy="lines")
    for t in tabs.tables:
        linhas = t.extract()
        h = None
        for i, r in enumerate(linhas):
            cs = [limpar(c) or "" for c in r]
            if "Output" in cs and any(c.startswith("Unit of Measure") for c in cs):
                h = i
                break
        if h is None:
            continue
        cab = [limpar(c) or "" for c in linhas[h]]
        sup = [limpar(c) or "" for c in linhas[h - 1]] if h > 0 else [""] * len(cab)
        col_nome = cab.index("Output")
        col_unid = next(j for j, c in enumerate(cab) if c.startswith("Unit of Measure"))
        col_num = [j for j, c in enumerate(cab) if re.match(r"^(EOP\s+\S+|\d{4})$", c)]
        col_rot = col_unid + 1
        if cab[col_rot] != "" or not col_num or col_rot >= col_num[0]:
            raise ValueError("cabecalho da matriz fora do esperado: %r" % cab)
        j_fis = next(j for j, c in enumerate(sup) if c.startswith("PHYSICAL"))
        j_fin = next(j for j, c in enumerate(sup) if c.startswith("FINANCIAL"))
        sem = {}
        for j in col_num:
            if j_fis <= j < j_fin:
                grandeza = "qtd"
            elif j >= j_fin:
                grandeza = "custo"
            else:
                raise ValueError("coluna %d fora de PHYSICAL/FINANCIAL" % j)
            per = "fim" if cab[j].startswith("EOP") else "ano"
            sem[j] = (grandeza, per)
        if len(set(sem.values())) != len(sem):
            raise ValueError("colunas repetidas no cabecalho: %r" % cab)
        colunas = [(sem[j], cab[j]) for j in col_num]

        atual = None
        for k in range(h + 1, len(linhas)):
            r = linhas[k]
            bbox = tuple(t.rows[k].bbox)

            def cel(j):
                return limpar(r[j]) if j < len(r) else None

            nome = cel(col_nome)
            rot = cel(col_rot)
            esquerda = [cel(j) for j in range(0, col_unid + 1)]
            if nome:
                m = RE_COD_NOME.match(nome)
                if m:
                    atual = {
                        "codigo_pmr": cel(0) or "",
                        "nome_impresso": nome,
                        "codigo_no_nome": nome.split(" ", 1)[0],
                        "produto": m.group(4).strip(),
                        "unidade": cel(col_unid) or "",
                        "linhas": {}, "rotulo_impresso": {}, "bboxes": [],
                        "avisos": [],
                    }
                    produtos.append(atual)
                else:
                    atual = None
                    outros.append(nome)
            elif any(esquerda):
                # texto na parte esquerda sem nome de produto: e outra secao
                # (ex.: "Other Cost"). Para de atribuir linhas ao produto anterior.
                if not (atual is not None and not atual["linhas"]):
                    outros.append(" / ".join(x for x in esquerda if x))
                    atual = None
            if rot in ROTULOS and atual is not None:
                chave = ROTULOS[rot]
                if chave in atual["linhas"]:
                    atual["avisos"].append("linha %s repetida no produto" % rot)
                atual["linhas"][chave] = {sem[j]: (r[j] if j < len(r) else None) for j in col_num}
                atual["rotulo_impresso"][chave] = rot
                atual["bboxes"].append(bbox)
            elif rot and rot not in ROTULOS and atual is not None:
                atual["avisos"].append("rotulo de linha desconhecido: %r" % rot)
        # cabecalho (duas linhas) para o recorte do PNG
        cab_bbox = fitz.Rect(t.rows[h].bbox)
        if h > 0:
            cab_bbox |= fitz.Rect(t.rows[h - 1].bbox)
        for p in produtos:
            p.setdefault("cab_bbox", tuple(cab_bbox))
    return produtos, outros, avisos, colunas


# ---------------------------------------------------------------------------
# CAMINHO 2: a mesma matriz lida pelas coordenadas das palavras (sem grade)
# ---------------------------------------------------------------------------


def ler_matriz_coordenadas(page):
    """Devolve {codigo_no_nome: {"unidade": str, "linhas": {rot: {(grandeza, per): texto}}}}.

    Independente do caminho 1: nao usa as linhas de grade. Usa so a posicao
    (x, y) de cada palavra. Os valores sao alinhados a direita no PDF; por isso
    cada numero vai para a ultima coluna cujo cabecalho comeca (x0) antes da
    borda direita (x1) do numero.
    """
    ws = palavras(page)
    unit = [w for w in ws if w[4] == "Unit"]
    if not unit:
        return {}, ["cabecalho 'Unit' nao achado"]
    yh = unit[0][1]
    lin_cab = sorted([w for w in ws if abs(w[1] - yh) < 2], key=lambda w: w[0])
    x_output = next(w for w in lin_cab if w[4] == "Output")[0]
    meas = next(w for w in lin_cab if w[4] == "Measure")
    toks = [w for w in lin_cab if w[0] > meas[2]]
    grupos = []
    i = 0
    while i < len(toks):
        if toks[i][4] == "EOP":
            grupos.append(("fim", toks[i][0]))
            i += 2
        elif re.match(r"^\d{4}$", toks[i][4]):
            grupos.append(("ano", toks[i][0]))
            i += 1
        else:
            i += 1
    n = len(grupos)
    if n not in (2, 4):
        return {}, ["cabecalho numerico com %d colunas" % n]
    cols = [("qtd" if k < n // 2 else "custo", per, x0) for k, (per, x0) in enumerate(grupos)]
    zona0, zona1 = meas[2], grupos[0][1]
    candidatos = sorted([w for w in ws if w[1] > yh + 3 and zona0 < w[0] < zona1
                         and w[4] in ("P", "P(a)", "A", "(a)")], key=lambda w: (w[1], w[0]))
    linhas = []
    for w in candidatos:
        if w[4] == "(a)":
            continue
        if w[4] == "P":
            tem_a = any(v[4] == "(a)" and abs(v[1] - w[1]) < 1.5 and 0 <= v[0] - w[2] < 4
                        for v in candidatos)
            rot = "Pa" if tem_a else "P"
        elif w[4] == "P(a)":
            rot = "Pa"
        else:
            rot = "A"
        linhas.append((w, rot))
    avisos = []
    blocos = []
    # borda direita do que esta a esquerda dos rotulos (nome e unidade)
    x_rotulos = min((w[0] for w, rot in linhas), default=zona1) - 1
    for w, rot in linhas:
        yc = (w[1] + w[3]) / 2
        valores = {}
        for v in ws:
            if abs((v[1] + v[3]) / 2 - yc) >= 3 or v[0] <= w[2] + 1:
                continue
            if not (v[4] in ("-", "\u2013") or RE_NUM.match(v[4])):
                continue
            k = max([i for i, c in enumerate(cols) if c[2] <= v[2]], default=-1)
            if k < 0:
                avisos.append("numero %r sem coluna (y=%.1f)" % (v[4], v[1]))
                continue
            chave = (cols[k][0], cols[k][1])
            if chave in valores:
                avisos.append("dois numeros na mesma coluna (y=%.1f)" % v[1])
            valores[chave] = v[4]
        if rot == "P":
            # inicio de bloco: nome e unidade estao na mesma linha do rotulo P
            esquerda = sorted([v for v in ws if abs((v[1] + v[3]) / 2 - yc) < 3
                               and v[0] >= x_output - 2 and v[2] < x_rotulos], key=lambda v: v[0])
            segmentos = []
            for v in esquerda:
                if segmentos and v[0] - segmentos[-1][-1][2] <= 15:
                    segmentos[-1].append(v)
                else:
                    segmentos.append([v])
            cod = ""
            unidade = ""
            if segmentos and abs(segmentos[0][0][0] - x_output) < 3:
                primeiro = segmentos[0][0][4]
                if re.match(r"^\d\.\d{1,2}(\.\d{1,2})?$", primeiro):
                    cod = primeiro
            unid_words, x_unid = [], None
            if len(segmentos) >= 2:
                unid_words = list(segmentos[-1])
                x_unid = unid_words[0][0]
            blocos.append({"codigo_no_nome": cod, "unid_words": unid_words, "x_unid": x_unid,
                           "yc": yc, "y_fim": yc, "linhas": {"P": valores}})
        elif blocos:
            if rot in blocos[-1]["linhas"]:
                avisos.append("linha %s repetida (y=%.1f)" % (rot, w[1]))
            blocos[-1]["linhas"][rot] = valores
            blocos[-1]["y_fim"] = yc
    saida = {}
    for b in blocos:
        # a unidade pode quebrar em mais de uma linha: junta as palavras da mesma
        # coluna que estao abaixo dela, ate a linha A do mesmo produto
        cont = []
        if b["x_unid"] is not None:
            cont = sorted([v for v in ws if b["yc"] + 3 <= (v[1] + v[3]) / 2 <= b["y_fim"] + 3
                           and v[0] >= b["x_unid"] - 2 and v[2] < x_rotulos], key=lambda v: (v[1], v[0]))
        b["unidade"] = " ".join(v[4] for v in b["unid_words"] + cont)
        if b["codigo_no_nome"]:
            saida[b["codigo_no_nome"]] = b
    return saida, avisos


# ---------------------------------------------------------------------------
# CAMINHO 3: texto bruto da pagina, para a amostra
# ---------------------------------------------------------------------------


def ler_texto_bruto(page, cod_nome):
    """Le, na ordem do texto bruto do PDF, a unidade e os valores de um produto.

    Devolve (nome, unidade, [(rotulo, [valores...]), ...]) ou None.
    """
    linhas = [l.strip() for l in page.get_text("text").split("\n")]
    idx = next((i for i, l in enumerate(linhas) if l.startswith(cod_nome + " ")), None)
    if idx is None:
        return None
    j = next((i for i in range(idx + 1, len(linhas)) if linhas[i] == "P"), None)
    if j is None:
        return None
    nome = " ".join(linhas[idx:j - 1])
    unidade = linhas[j - 1]
    pagina_txt = str(page.number + 1)
    seq = []
    for i in range(j, len(linhas)):
        l = linhas[i]
        if l in ROTULOS:
            if len(seq) == 3:
                break
            seq.append((ROTULOS[l], []))
            continue
        if not seq:
            break
        prox = linhas[i + 1] if i + 1 < len(linhas) else ""
        # o codigo interno do proximo produto ("1.2", "1.02") parece numero:
        # para quando ele vem seguido de uma linha de nome de produto
        if RE_COD_PMR.match(l) and RE_COD_NOME.match(prox):
            break
        # numero da pagina no rodape (layout 2021-S2 em diante)
        if l == pagina_txt and not any(x for x in linhas[i + 1:]):
            break
        if l in ("-", "\u2013") or RE_NUM.match(l):
            seq[-1][1].append(l)
            continue
        break
    return nome, unidade, seq


def comparar_caminhos(prods, coord):
    """Compara caminho 1 (grade) x caminho 2 (coordenadas). Devolve (comparacoes, divergencias)."""
    n, div = 0, []
    for p in prods:
        c = coord.get(p["codigo_no_nome"])
        if c is None:
            div.append((p["codigo_no_nome"], "produto", "nao achado pelas coordenadas", ""))
            continue
        n += 1
        if c["unidade"] != p["unidade"]:
            div.append((p["codigo_no_nome"], "unidade", p["unidade"], c["unidade"]))
        for rot in ORDEM_LINHAS:
            g = p["linhas"].get(rot, {})
            cc = c["linhas"].get(rot, {})
            for chave in set(g) | set(cc):
                n += 1
                vg = limpar(g.get(chave)) or ""
                vc = cc.get(chave, "")
                if vg != vc:
                    div.append((p["codigo_no_nome"], "%s %s_%s" % (rot, chave[0], chave[1]), vg, vc))
    return n, div


def conferir_texto_bruto(page, p, colunas):
    """Caminho 3 para um produto. Devolve (resultado, detalhe, sequencia_lida_no_texto)."""
    ordem = [c for c, _ in colunas]
    esperado = []
    for rot in ORDEM_LINHAS:
        vals = [limpar(v) for _, v in sorted(p["linhas"].get(rot, {}).items(),
                                             key=lambda kv: ordem.index(kv[0]))]
        esperado.append((rot, [v for v in vals if v]))
    bruto = ler_texto_bruto(page, p["codigo_no_nome"])
    if bruto is None:
        return "NAO CONFERIDO", "produto nao localizado no texto bruto", ""
    nome_b, unid_b, seq = bruto
    problemas = []
    if seq != esperado:
        problemas.append("valores diferem")
    if unid_b != p["unidade"]:
        problemas.append("unidade difere (%r x %r)" % (unid_b, p["unidade"]))
    if limpar(nome_b) != p["nome_impresso"]:
        problemas.append("nome difere")
    seq_txt = " / ".join("%s: %s" % ({"Pa": "P(a)"}.get(r, r), " ".join(v) or "(nada)") for r, v in seq)
    return ("OK" if not problemas else "DIVERGENTE"), "; ".join(problemas), seq_txt


def autoteste(doc, pno, colunas):
    """Prova que as conferencias MORDEM: estraga de proposito a leitura do caminho 1
    e verifica se os caminhos 2 e 3 acusam. Uma conferencia que nao reprova um erro
    plantado nao prova nada. Devolve lista de (mutacao, acusou?)."""
    import copy
    page = doc[pno]
    prods, _, _, _ = ler_matriz_grade(page)
    coord, _ = ler_matriz_coordenadas(page)
    saida = []
    _, div0 = comparar_caminhos(prods, coord)
    saida.append(("sem mutacao (controle): caminho 2 nao deve acusar nada", len(div0) == 0))

    def primeiro(cond):
        return next((q for q in prods if cond(q)), None)

    # 1) troca as linhas P(a) e A de um produto em que elas diferem
    m = copy.deepcopy(prods)
    alvo = next((q for q in m if q["linhas"].get("Pa") != q["linhas"].get("A")), None)
    if alvo:
        alvo["linhas"]["Pa"], alvo["linhas"]["A"] = alvo["linhas"]["A"], alvo["linhas"]["Pa"]
        _, div = comparar_caminhos(m, coord)
        saida.append(("troca das linhas P(a) e A no produto %s" % alvo["codigo_no_nome"], len(div) > 0))
        r, _, _ = conferir_texto_bruto(page, alvo, colunas)
        saida.append(("idem, conferido pelo texto bruto (caminho 3)", r == "DIVERGENTE"))
    # 2) troca as colunas de quantidade e custo de fim de projeto na linha P
    m = copy.deepcopy(prods)
    alvo = next((q for q in m if q["linhas"]["P"].get(("qtd", "fim")) != q["linhas"]["P"].get(("custo", "fim"))), None)
    if alvo:
        d = alvo["linhas"]["P"]
        d[("qtd", "fim")], d[("custo", "fim")] = d[("custo", "fim")], d[("qtd", "fim")]
        _, div = comparar_caminhos(m, coord)
        saida.append(("troca qtd x custo de fim de projeto (linha P) no produto %s" % alvo["codigo_no_nome"],
                      len(div) > 0))
        r, _, _ = conferir_texto_bruto(page, alvo, colunas)
        saida.append(("idem, conferido pelo texto bruto (caminho 3)", r == "DIVERGENTE"))
    # 3) altera um digito do custo P(a) de fim de projeto
    m = copy.deepcopy(prods)
    alvo = next((q for q in m if RE_NUM.match(limpar(q["linhas"]["Pa"].get(("custo", "fim"))) or "")), None)
    if alvo:
        v = limpar(alvo["linhas"]["Pa"][("custo", "fim")])
        alvo["linhas"]["Pa"][("custo", "fim")] = v[:-1] + ("1" if v[-1] != "1" else "2")
        _, div = comparar_caminhos(m, coord)
        saida.append(("um digito alterado no custo P(a) de fim de projeto do produto %s" % alvo["codigo_no_nome"],
                      len(div) > 0))
        r, _, _ = conferir_texto_bruto(page, alvo, colunas)
        saida.append(("idem, conferido pelo texto bruto (caminho 3)", r == "DIVERGENTE"))
    # 4) troca a unidade
    m = copy.deepcopy(prods)
    m[0]["unidade"] = m[0]["unidade"] + "x"
    _, div = comparar_caminhos(m, coord)
    saida.append(("unidade alterada no produto %s" % m[0]["codigo_no_nome"], len(div) > 0))
    r, _, _ = conferir_texto_bruto(page, m[0], colunas)
    saida.append(("idem, conferido pelo texto bruto (caminho 3)", r == "DIVERGENTE"))
    return saida


# ---------------------------------------------------------------------------
# Tabela de alteracoes ("CHANGES TO THE MATRIX")
# ---------------------------------------------------------------------------

MAPA_COLUNAS_ALT = {
    "Section": "secao", "Name": "item", "Reasons": "motivo",
    "Type of Change": "tipo", "Subtype": "subtipo", "Sub type": "subtipo",
    "Modified By": "modificado_por", "Entered in the System": "data",
    "Entered in System": "data",
}

# traducao literal do subtipo impresso (nao e interpretacao: e o que o rotulo diz)
O_QUE_MUDOU = [
    ("Physical EOP P(a)", "quantidade fisica de fim de projeto (P(a))"),
    ("Financial EOP P(a)", "custo de fim de projeto (P(a))"),
    ("Financial Historical Actual", "custo realizado (A) de periodos anteriores"),
    ("Modify EOP P(a)", "meta de fim de projeto do indicador (P(a))"),
    ("Update Baseline", "linha de base e/ou ano da linha de base"),
    ("Modify Indicator name", "nome do indicador"),
    ("measure progress using Output", "forma de medir o progresso (passa a medir pelo produto)"),
    ("CRF indicator", "vinculo com indicador corporativo do BID (CRF)"),
    ("IF indicator", "vinculo com indicador corporativo do BID (IF)"),
]


def ler_alteracoes(doc):
    """Devolve (linhas_de_alteracao, info_do_informe)."""
    info = {"pagina": "", "sem_informacao": False, "colunas": [], "paginas_tabela": []}
    saida = []
    for pno, page in enumerate(doc):
        txt = page.get_text()
        if "CHANGES TO THE MATRIX" in txt:
            info["pagina"] = pno + 1
            if "No information available for this section" in txt:
                info["sem_informacao"] = True
        tabs = page.find_tables(strategy="lines")
        for t in tabs.tables:
            linhas = t.extract()
            if not linhas:
                continue
            cab = [limpar(c) or "" for c in linhas[0]]
            if not ("Section" in cab and "Name" in cab):
                continue
            nomes = []
            for c in cab:
                if c not in MAPA_COLUNAS_ALT:
                    raise ValueError("coluna desconhecida na tabela de alteracoes: %r" % c)
                nomes.append(MAPA_COLUNAS_ALT[c])
            info["colunas"] = cab
            info["paginas_tabela"].append(pno + 1)
            # conferencia independente: cada linha de alteracao tem uma data; conta as
            # datas no texto bruto da pagina para comparar com as linhas lidas pela grade
            info["datas_texto_bruto"] = info.get("datas_texto_bruto", 0) + sum(
                1 for l in txt.split("\n") if RE_DATA.match(l.strip()))
            herdado = {}
            for k, r in enumerate(linhas[1:], start=1):
                if all((c is None or limpar(c) == "") for c in r):
                    continue
                reg = {"pagina": pno + 1, "obs": []}
                for nome_col, c in zip(nomes, r):
                    if c is None and nome_col in ("secao", "item", "motivo", "tipo"):
                        # celula mesclada com a de cima: vale o texto de cima
                        reg[nome_col] = herdado.get(nome_col, "")
                        reg.setdefault("herdados", []).append(nome_col)
                    elif c is None:
                        reg[nome_col] = ""
                        reg["obs"].append("%s: celula ausente na grade" % nome_col)
                    else:
                        reg[nome_col] = limpar(c)
                        herdado[nome_col] = limpar(c)
                reg["tem_motivo"] = "motivo" in nomes
                saida.append(reg)
    return saida, info


# ---------------------------------------------------------------------------
# PNG de conferencia: cabecalho da matriz + linhas dos produtos sorteados
# ---------------------------------------------------------------------------


def gerar_png(doc, pedacos, destino, zoom=2.0):
    """pedacos = [(pno, Rect), ...] empilhados de cima para baixo."""
    largura = max(r.width for _, r in pedacos)
    altura = sum(r.height for _, r in pedacos) + 6 * (len(pedacos) - 1)
    novo = fitz.open()
    pg = novo.new_page(width=largura, height=altura)
    pg.draw_rect(pg.rect, color=None, fill=(1, 1, 1))
    y = 0
    for pno, r in pedacos:
        alvo = fitz.Rect(0, y, r.width, y + r.height)
        pg.show_pdf_page(alvo, doc, pno, clip=r)
        y += r.height + 6
        if y < altura:
            pg.draw_line(fitz.Point(0, y - 3), fitz.Point(largura, y - 3), color=(0.8, 0, 0), width=0.6)
    pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    pix.save(destino)
    novo.close()


# ---------------------------------------------------------------------------
# Leitura de um informe inteiro
# ---------------------------------------------------------------------------


def ler_informe(caminho):
    doc = fitz.open(caminho)
    cab = ler_cabecalho(doc[0])
    cab["arquivo"] = os.path.basename(caminho)
    cab["paginas"] = doc.page_count
    cab["sha256"] = sha256(caminho)
    cab["tamanho_pagina"] = "%dx%d" % (round(doc[0].rect.width), round(doc[0].rect.height))

    pags = paginas_da_matriz(doc)
    produtos, outros, colunas_vistas = [], [], []
    conf2 = {"celulas": 0, "divergencias": [], "avisos": []}
    for pno in pags:
        page = doc[pno]
        prods, out, _, colunas = ler_matriz_grade(page)
        for p in prods:
            p["pagina"] = pno + 1
            p["pno"] = pno
        produtos.extend(prods)
        outros.extend(out)
        if colunas:
            colunas_vistas.append(tuple(colunas))
        # caminho 2 na mesma pagina
        coord, avisos = ler_matriz_coordenadas(page)
        conf2["avisos"].extend("p.%d: %s" % (pno + 1, a) for a in avisos)
        n, div = comparar_caminhos(prods, coord)
        conf2["celulas"] += n
        conf2["divergencias"].extend(div)
    if len(set(colunas_vistas)) > 1:
        raise ValueError("%s: colunas diferentes entre as paginas da matriz" % caminho)
    cab["colunas_matriz"] = colunas_vistas[0] if colunas_vistas else ()
    cab["texto_matriz"] = " ".join(doc[p].get_text() for p in pags)

    alteracoes, info_alt = ler_alteracoes(doc)
    return doc, cab, produtos, outros, conf2, alteracoes, info_alt


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

COLUNAS_MESTRA = [
    "informe", "periodo", "arquivo", "pagina", "codigo", "codigo_pmr", "codigo_no_nome",
    "produto", "unidade",
    "rotulo_coluna_fim", "qtd_fim_P", "qtd_fim_Pa", "qtd_fim_A",
    "custo_fim_P", "custo_fim_Pa", "custo_fim_A",
    "rotulo_coluna_ano", "qtd_ano_P", "qtd_ano_Pa", "qtd_ano_A",
    "custo_ano_P", "custo_ano_Pa", "custo_ano_A",
    "rotulo_ajustado_impresso", "obs",
]

COLUNAS_ALT = [
    "informe", "periodo", "pagina", "ordem", "secao", "item", "codigo",
    "tipo_alteracao", "subtipo", "o_que_mudou", "modificado_por",
    "data_registro_impressa", "data_registro_iso",
    "campo_motivo_existe", "motivo", "motivo_vazio", "valores_citados_no_motivo",
    "valor_anterior_impresso", "valor_novo_impresso",
    "informe_comparado", "valor_anterior_comparado", "valor_novo_comparado",
    "obs",
]

COLUNAS_INFORMES = [
    "informe", "periodo", "arquivo", "sha256", "paginas", "tamanho_pagina", "titulo",
    "idioma_produtos", "variante_layout", "ciclo_impresso", "estagio_validacao",
    "ultima_atualizacao", "ultima_atualizacao_iso",
    "validacao_chefe_operacoes", "validacao_chefe_operacoes_iso",
    "validacao_chefe_divisao", "validacao_chefe_divisao_iso",
    "validacao_representante", "validacao_representante_iso",
    "colunas_da_matriz", "produtos_lidos", "produtos_faltantes",
    "tabela_alteracoes", "linhas_de_alteracao", "colunas_tabela_alteracoes",
    "campo_motivo_existe", "motivos_vazios",
]


def idioma(texto):
    es = sum(texto.count(p) for p in ("Gestión", "Administración", "hacendaria", "planeación"))
    pt = sum(texto.count(p) for p in ("Gestão", "Administração", "fazendária"))
    if es and not pt:
        return "espanhol"
    if pt and not es:
        return "portugues"
    return "indefinido (es=%d, pt=%d)" % (es, pt)


def gravar_csv(caminho, colunas, linhas):
    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=colunas, delimiter=";", quoting=csv.QUOTE_MINIMAL,
                           lineterminator="\r\n")
        w.writeheader()
        for l in linhas:
            w.writerow({c: l.get(c, "") for c in colunas})


def md_tabela(cabecalho, linhas):
    def esc(x):
        return str(x).replace("|", "\\|").replace("\n", " ")
    out = ["| " + " | ".join(cabecalho) + " |", "|" + "---|" * len(cabecalho)]
    for l in linhas:
        out.append("| " + " | ".join(esc(x) for x in l) + " |")
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--fontes", default=PASTA_FONTES)
    ap.add_argument("--saida", default=PASTA_SAIDA)
    args = ap.parse_args()

    arquivos = sorted(glob.glob(os.path.join(args.fontes, "PMR *.pdf")))
    if len(arquivos) != 16:
        print("AVISO: esperava 16 PDFs, achei %d" % len(arquivos))
    os.makedirs(os.path.join(args.saida, "conferencia_png"), exist_ok=True)

    informes = []
    for caminho in arquivos:
        informes.append(ler_informe(caminho))
    # ordem cronologica pelo ciclo impresso (nao pelo nome do arquivo)
    informes.sort(key=lambda x: x[1]["informe"])

    # --- correspondencia codigo interno do PMR <-> codigo contratual (3 niveis)
    corresp = {}
    fontes_corresp = []
    for doc, cab, produtos, *_ in informes:
        usou = False
        for p in produtos:
            if len(p["codigo_no_nome"].split(".")) == 3 and RE_COD_PMR.match(p["codigo_pmr"]):
                k = chave_codigo(p["codigo_pmr"])
                corresp.setdefault(k, set()).add(p["codigo_no_nome"])
                usou = True
        if usou:
            fontes_corresp.append(cab["informe"])
    conflitos = {k: v for k, v in corresp.items() if len(v) != 1}
    if conflitos:
        raise ValueError("correspondencia de codigos inconsistente: %r" % conflitos)
    corresp = {k: next(iter(v)) for k, v in corresp.items()}

    # --- monta a tabela mestra
    mestra, por_chave = [], {}
    resumo_inf = []
    variantes = {}
    nomes_por_codigo = {}
    for doc, cab, produtos, outros, conf2, alteracoes, info_alt in informes:
        cab["idioma"] = idioma(cab["texto_matriz"])
        colunas = cab["colunas_matriz"]
        rot_fim = next((r for (g, per), r in colunas if per == "fim"), "")
        rot_ano = next((r for (g, per), r in colunas if per == "ano"), "")
        lidos = set()
        for p in produtos:
            obs = list(p["avisos"])
            cnome = p["codigo_no_nome"]
            if len(cnome.split(".")) == 3:
                codigo = cnome
            else:
                k = chave_codigo(p["codigo_pmr"]) if RE_COD_PMR.match(p["codigo_pmr"]) else None
                codigo = corresp.get(k, "")
                if codigo:
                    obs.append("codigo de 3 niveis nao impresso neste informe (o nome traz '%s'); "
                               "atribuido pela correspondencia codigo interno do PMR '%s' -> '%s', "
                               "observada nos informes %s a %s"
                               % (cnome, p["codigo_pmr"], codigo, fontes_corresp[0], fontes_corresp[-1]))
                else:
                    obs.append("codigo de 3 niveis nao impresso e sem correspondencia")
            if p["codigo_pmr"] and RE_COD_PMR.match(p["codigo_pmr"]) and codigo \
                    and corresp.get(chave_codigo(p["codigo_pmr"])) not in (None, codigo):
                obs.append("codigo interno do PMR '%s' nao corresponde ao codigo do nome" % p["codigo_pmr"])
            if codigo in lidos:
                obs.append("produto repetido neste informe")
            lidos.add(codigo)
            linha = {
                "informe": cab["informe"], "periodo": cab["periodo"], "arquivo": cab["arquivo"],
                "pagina": p["pagina"], "codigo": codigo, "codigo_pmr": p["codigo_pmr"],
                "codigo_no_nome": cnome, "produto": p["produto"], "unidade": p["unidade"],
                "rotulo_coluna_fim": rot_fim, "rotulo_coluna_ano": rot_ano,
                "rotulo_ajustado_impresso": p["rotulo_impresso"].get("Pa", ""),
            }
            for rot in ORDEM_LINHAS:
                if rot not in p["linhas"]:
                    obs.append("linha %s nao encontrada" % rot)
                    continue
                for (g, per), bruto in p["linhas"][rot].items():
                    valor, problema = ler_valor(bruto)
                    col = "%s_%s_%s" % (g, per, rot)
                    linha[col] = valor
                    if problema and per == "fim":
                        obs.append("%s: %s" % (col, problema))
                    elif problema and "nao reconhecido" in problema:
                        obs.append("%s: %s" % (col, problema))
            if not p["unidade"]:
                obs.append("unidade: celula vazia")
            linha["obs"] = " | ".join(obs)
            linha["_p"] = p
            mestra.append(linha)
            por_chave[(cab["informe"], codigo)] = linha
            nomes_por_codigo.setdefault(codigo, []).append((cab["informe"], p["produto"], p["unidade"]))
        faltam = [c for c in CODIGOS_CONTRATO if c not in lidos]
        sobram = sorted(c for c in lidos if c not in CODIGOS_CONTRATO)
        cab["produtos_lidos"] = len(lidos)
        cab["faltantes"] = faltam
        cab["sobram"] = sobram
        cab["outros"] = outros
        # assinatura do layout
        cods_pmr = {("com zero a esquerda (1.01 ... 2.12)" if any(re.match(r"^\d\.0\d$", q["codigo_pmr"]) for q in produtos)
                     else "sem zero (1.1 ... 2.12)") for p in produtos[:1]}
        niveis = {len(p["codigo_no_nome"].split(".")) for p in produtos}
        rot_aj = {p["rotulo_impresso"].get("Pa", "") for p in produtos}
        sig = (
            cab["titulo"], cab["idioma"], cab["tamanho_pagina"],
            " + ".join("%s/%s" % (g, per) for (g, per), _ in colunas),
            "/".join(sorted(cods_pmr)), "/".join(str(x) for x in sorted(niveis)),
            "/".join(sorted(rot_aj)),
        )
        if sig not in variantes:
            variantes[sig] = "V%d" % (len(variantes) + 1)
        cab["variante"] = variantes[sig]
        resumo_inf.append(cab)

    conjunto_lido = {l["codigo"] for l in mestra}
    if conjunto_lido != set(CODIGOS_CONTRATO):
        print("AVISO: conjunto de codigos lido difere da lista de 24: %r"
              % sorted(conjunto_lido ^ set(CODIGOS_CONTRATO)))

    # --- tabela de alteracoes
    ordem_inf = [c["informe"] for c in resumo_inf]
    alt_saida = []
    for doc, cab, produtos, outros, conf2, alteracoes, info_alt in informes:
        i_inf = ordem_inf.index(cab["informe"])
        anterior = ordem_inf[i_inf - 1] if i_inf > 0 else ""
        n_vazios = 0
        for n, a in enumerate(alteracoes, start=1):
            obs = list(a["obs"])
            if a.get("herdados"):
                obs.append("celula(s) %s mesclada(s) com a linha de cima no PDF: texto herdado"
                           % ", ".join(a["herdados"]))
            item = a.get("item", "")
            m = RE_COD_NOME.match(item or "")
            codigo = ""
            if m and m.group(3) is not None:
                codigo = "%s.%s.%s" % (m.group(1), m.group(2), m.group(3))
            elif "Milestone" in (a.get("tipo") or "") or item.startswith("Etapa"):
                obs.append("marco (milestone); a tabela nao diz a que produto ele pertence")
            sub = a.get("subtipo", "")
            o_que = next((t for chave, t in O_QUE_MUDOU if chave in sub), "")
            if not o_que:
                obs.append("subtipo sem traducao na lista do script")
            motivo = a.get("motivo", "") if a["tem_motivo"] else ""
            vazio = ("sim" if not motivo else "nao") if a["tem_motivo"] else "n.a. (coluna inexistente)"
            if a["tem_motivo"] and not motivo:
                n_vazios += 1
            citados = "; ".join("%s -> %s" % par for par in re.findall(
                r"(?:\bde\s+)?(\d+(?:[.,]\d+)?)\s+para\s+(\d+(?:[.,]\d+)?)", motivo))
            reg = {
                "informe": cab["informe"], "periodo": cab["periodo"], "pagina": a["pagina"],
                "ordem": n, "secao": a.get("secao", ""), "item": item, "codigo": codigo,
                "tipo_alteracao": a.get("tipo", ""), "subtipo": sub, "o_que_mudou": o_que,
                "modificado_por": a.get("modificado_por", ""),
                "data_registro_impressa": a.get("data", ""),
                "data_registro_iso": data_iso(a.get("data", "")),
                "campo_motivo_existe": "sim" if a["tem_motivo"] else "nao",
                "motivo": motivo, "motivo_vazio": vazio, "valores_citados_no_motivo": citados,
            }
            obs.append("a tabela de alteracoes nao tem coluna de valor anterior nem de valor novo")
            if not a.get("modificado_por"):
                obs.append("modificado_por: celula vazia")
            # comparacao com o informe anterior (derivada, nao impressa na tabela)
            campo = None
            if codigo and "Physical EOP P(a)" in sub:
                campo = "qtd_fim_Pa"
            elif codigo and "Financial EOP P(a)" in sub:
                campo = "custo_fim_Pa"
            if campo and anterior:
                la = por_chave.get((anterior, codigo))
                lc = por_chave.get((cab["informe"], codigo))
                if la and lc:
                    reg["informe_comparado"] = anterior
                    reg["valor_anterior_comparado"] = la.get(campo, "")
                    reg["valor_novo_comparado"] = lc.get(campo, "")
                    obs.append("valores comparados = %s no informe anterior (%s) e neste; calculo do script, "
                               "nao impresso na tabela" % (campo, anterior))
            reg["obs"] = " | ".join(obs)
            alt_saida.append(reg)
        cab["alt_info"] = info_alt
        cab["alt_n"] = len(alteracoes)
        cab["alt_vazios"] = n_vazios
        cab["alt_tem_motivo"] = any(a["tem_motivo"] for a in alteracoes)

    # --- gravacao dos CSVs
    gravar_csv(os.path.join(args.saida, "tabela_mestra_pmr.csv"), COLUNAS_MESTRA, mestra)
    gravar_csv(os.path.join(args.saida, "alteracoes_pmr.csv"), COLUNAS_ALT, alt_saida)
    linhas_inf = []
    for cab in resumo_inf:
        ia = cab["alt_info"]
        if ia["sem_informacao"]:
            tab = "sem informacao ('No information available for this section')"
        elif ia["paginas_tabela"]:
            tab = "tabela na(s) pagina(s) %s" % ",".join(str(x) for x in ia["paginas_tabela"])
        else:
            tab = "secao nao encontrada"
        linhas_inf.append({
            "informe": cab["informe"], "periodo": cab["periodo"], "arquivo": cab["arquivo"],
            "sha256": cab["sha256"], "paginas": cab["paginas"], "tamanho_pagina": cab["tamanho_pagina"],
            "titulo": cab["titulo"], "idioma_produtos": cab["idioma"], "variante_layout": cab["variante"],
            "ciclo_impresso": cab["ciclo_impresso"], "estagio_validacao": cab["estagio_validacao"],
            "ultima_atualizacao": cab["ultima_atualizacao"],
            "ultima_atualizacao_iso": data_iso(cab["ultima_atualizacao"]),
            "validacao_chefe_operacoes": cab["validacao_chefe_operacoes"],
            "validacao_chefe_operacoes_iso": data_iso(cab["validacao_chefe_operacoes"]),
            "validacao_chefe_divisao": cab["validacao_chefe_divisao"],
            "validacao_chefe_divisao_iso": data_iso(cab["validacao_chefe_divisao"]),
            "validacao_representante": cab["validacao_representante"],
            "validacao_representante_iso": data_iso(cab["validacao_representante"]),
            "colunas_da_matriz": " + ".join("%s (%s/%s)" % (r, g, per) for (g, per), r in cab["colunas_matriz"]),
            "produtos_lidos": cab["produtos_lidos"], "produtos_faltantes": ", ".join(cab["faltantes"]),
            "tabela_alteracoes": tab, "linhas_de_alteracao": cab["alt_n"],
            "colunas_tabela_alteracoes": " | ".join(ia["colunas"]),
            "campo_motivo_existe": ("sim" if cab["alt_tem_motivo"] else "nao") if cab["alt_n"] else "n.a.",
            "motivos_vazios": cab["alt_vazios"] if cab["alt_tem_motivo"] else "",
        })
    gravar_csv(os.path.join(args.saida, "informes_pmr.csv"), COLUNAS_INFORMES, linhas_inf)

    # --- caminho 3: amostra de 3 produtos por informe + PNG
    amostra = []
    for doc, cab, produtos, outros, conf2, alteracoes, info_alt in informes:
        linhas = [l for l in mestra if l["informe"] == cab["informe"]]
        rnd = random.Random("%s|%s" % (SEMENTE_AMOSTRA, cab["informe"]))
        sorteados = sorted(rnd.sample(linhas, 3), key=lambda l: chave_codigo(l["codigo"]))
        pedacos = []
        for l in sorteados:
            p = l["_p"]
            res, det, seq_txt = conferir_texto_bruto(doc[p["pno"]], p, cab["colunas_matriz"])
            amostra.append({"informe": cab["informe"], "codigo": l["codigo"], "pagina": p["pagina"],
                            "unidade": p["unidade"], "resultado": res, "detalhe": det,
                            "texto_bruto": seq_txt,
                            "csv": "qtd fim P/P(a)/A = %s/%s/%s; custo fim P/P(a)/A = %s/%s/%s" % (
                                l.get("qtd_fim_P", ""), l.get("qtd_fim_Pa", ""), l.get("qtd_fim_A", ""),
                                l.get("custo_fim_P", ""), l.get("custo_fim_Pa", ""), l.get("custo_fim_A", ""))})
            r = fitz.Rect(p["bboxes"][0])
            for b in p["bboxes"][1:]:
                r |= fitz.Rect(b)
            pedacos.append((p["pno"], fitz.Rect(p["cab_bbox"]), r))
        blocos = []
        for pno, cabr, r in pedacos:
            if not blocos or blocos[-1][0] != pno:
                blocos.append((pno, cabr))
            blocos.append((pno, r))
        nome_png = "%s.png" % cab["informe"]
        gerar_png(doc, blocos, os.path.join(args.saida, "conferencia_png", nome_png))
        cab["png"] = nome_png

    # --- autoteste por mutacao: uma pagina de matriz de cada variante de layout
    testes, vistos = [], set()
    for doc, cab, produtos, *_ in informes:
        if cab["variante"] in vistos:
            continue
        vistos.add(cab["variante"])
        pno = produtos[0]["pno"]
        for desc, ok in autoteste(doc, pno, cab["colunas_matriz"]):
            testes.append((cab["variante"], cab["informe"], pno + 1, desc, ok))

    escrever_relatorio(args, resumo_inf, informes, mestra, alt_saida, amostra, variantes,
                       corresp, fontes_corresp, nomes_por_codigo, testes)
    print("autoteste por mutacao: %d de %d testes acusaram o erro plantado (ou o controle passou)"
          % (sum(1 for t in testes if t[4]), len(testes)))

    # resumo no console
    n_div = sum(len(x[4]["divergencias"]) for x in informes)
    n_cel = sum(x[4]["celulas"] for x in informes)
    print("informes lidos: %d" % len(informes))
    print("linhas informe x produto: %d" % len(mestra))
    print("caminho 2 (coordenadas): %d celulas comparadas, %d divergencias" % (n_cel, n_div))
    print("caminho 3 (texto bruto): %d produtos sorteados, %d OK"
          % (len(amostra), sum(1 for a in amostra if a["resultado"] == "OK")))
    print("alteracoes: %d linhas" % len(alt_saida))
    for doc, *_ in informes:
        doc.close()


# ---------------------------------------------------------------------------
# Relatorio de conferencia (markdown)
# ---------------------------------------------------------------------------


def escrever_relatorio(args, resumo_inf, informes, mestra, alt_saida, amostra, variantes,
                       corresp, fontes_corresp, nomes_por_codigo, testes):
    L = []
    w = L.append
    w("# Conferencia da extracao dos 16 PMR (BR-L1500, PROFISCO II-MA)")
    w("")
    w("> Gerado por `extrair_pmr.py`. Rodar o script de novo, com os mesmos PDFs (mesmo SHA-256, ver "
      "`informes_pmr.csv`), reproduz este arquivo e os CSVs byte a byte. Nenhuma interpretacao: so o que "
      "esta impresso e a prova da leitura.")
    w("")
    # ---- resumo
    n_inf = len(resumo_inf)
    n_lin = len(mestra)
    w("## 1. Resumo")
    w("")
    w("- Informes lidos: **%d** (de %s a %s)." % (n_inf, resumo_inf[0]["informe"], resumo_inf[-1]["informe"]))
    w("- Linhas informe x produto: **%d** (esperado 16 x 24 = 384)." % n_lin)
    falt = [(c["informe"], c["faltantes"]) for c in resumo_inf if c["faltantes"]]
    w("- Produtos faltantes: %s." % ("nenhum; os 24 foram achados em todos os informes" if not falt else
                                     "; ".join("%s: %s" % (i, ", ".join(f)) for i, f in falt)))
    cols_fim = ["qtd_fim_P", "qtd_fim_Pa", "qtd_fim_A", "custo_fim_P", "custo_fim_Pa", "custo_fim_A"]
    cols_ano = ["qtd_ano_P", "qtd_ano_Pa", "qtd_ano_A", "custo_ano_P", "custo_ano_Pa", "custo_ano_A"]
    vazias_fim = sum(1 for l in mestra for c in cols_fim if l.get(c, "") == "")
    tracos_fim = sum(1 for l in mestra for c in cols_fim if l.get(c, "") == "-")
    w("- Celulas de fim de projeto (6 por linha = %d): **%d vazias**, %d com traco impresso (\"-\")."
      % (6 * n_lin, vazias_fim, tracos_fim))
    n_div = sum(len(x[4]["divergencias"]) for x in informes)
    n_cel = sum(x[4]["celulas"] for x in informes)
    w("- Caminho 2 (coordenadas x grade): %d comparacoes (unidade + celulas de valor), **%d divergencias**."
      % (n_cel, n_div))
    ok3 = sum(1 for a in amostra if a["resultado"] == "OK")
    w("- Caminho 3 (texto bruto, 3 produtos sorteados por informe): **%d de %d OK**." % (ok3, len(amostra)))
    w("- Tabela de alteracoes: %d linhas em %d informes." % (
        len(alt_saida), sum(1 for c in resumo_inf if c["alt_n"])))
    w("")
    w("Convencoes do CSV: numero sem separador de milhar; `-` = traco impresso no PDF; celula vazia = nada "
      "impresso (o motivo vai na coluna `obs` quando e coluna de fim de projeto). `Pa` = linha P(a), "
      "planejado ajustado. `fim` = coluna EOP (fim de projeto); `ano` = coluna do ano do ciclo.")
    w("")
    # ---- por informe
    w("## 2. Os informes")
    w("")
    linhas = []
    for c in resumo_inf:
        ia = c["alt_info"]
        alt = ("sem informacao" if ia["sem_informacao"] else
               ("%d linhas (p. %s)" % (c["alt_n"], ",".join(str(x) for x in ia["paginas_tabela"]))
                if ia["paginas_tabela"] else "nao achada"))
        linhas.append([c["informe"], c["periodo"], c["paginas"], c["variante"], c["estagio_validacao"],
                       data_iso(c["validacao_chefe_operacoes"]) or "(vazio)",
                       data_iso(c["validacao_chefe_divisao"]) or "(vazio)",
                       data_iso(c["validacao_representante"]) or "(vazio)",
                       "%d/24" % c["produtos_lidos"], alt])
    w(md_tabela(["Informe", "Periodo", "Pags", "Layout", "Estagio de validacao impresso",
                 "Valid. Chefe de Operacoes", "Valid. Chefe de Divisao", "Valid. Representante",
                 "Produtos", "Tabela de alteracoes"], linhas))
    w("")
    div_nome = [c["informe"] for c in resumo_inf if c["informe"] not in c["arquivo"]]
    w("O informe (ex.: 2017-S2) e o periodo foram tirados do campo \"Year- PMR Cycle\" impresso na pagina 1 "
      "(ex.: \"Second period Jan-Dec 2017\"), nao do nome do arquivo. Conferencia com o nome do arquivo: %s."
      % ("os 16 batem" if not div_nome else "NAO batem em " + ", ".join(div_nome)))
    w("")
    w("Datas impressas no padrao americano (mes/dia/ano), convertidas para AAAA-MM-DD. O padrao foi "
      "confirmado no proprio material: ha datas com dia maior que 12 na segunda posicao (ex.: 03/28/18), "
      "e o script para com erro se alguma data tiver mes > 12. \"(vazio)\" = o campo existe na pagina 1 "
      "e nao tem data impressa.")
    w("")
    # ---- variantes
    w("## 3. Variantes de layout encontradas")
    w("")
    inv = {}
    for c in resumo_inf:
        inv.setdefault(c["variante"], []).append(c["informe"])
    linhas = []
    for sig, v in sorted(variantes.items(), key=lambda kv: int(kv[1][1:])):
        titulo, idi, tam, cols, cpmr, niveis, rotaj = sig
        linhas.append([v, ", ".join(inv[v]), titulo, idi, tam + " pt", cols, cpmr, niveis, rotaj])
    w(md_tabela(["Variante", "Informes", "Titulo", "Idioma dos produtos", "Pagina", "Colunas numericas",
                 "Codigo interno", "Niveis do codigo no nome", "Rotulo do ajustado"], linhas))
    w("")
    rot_fim = {}
    for c in resumo_inf:
        r = next((r for (g, per), r in c["colunas_matriz"] if per == "fim"), "")
        rot_fim.setdefault(r, []).append(c["informe"])
    w("Rotulo impresso da coluna de fim de projeto: " + "; ".join(
        "`%s` em %s" % (r, ", ".join(i)) for r, i in rot_fim.items()) + ".")
    w("")
    w("Colunas da tabela de alteracoes, por informe que tem tabela:")
    w("")
    for c in resumo_inf:
        if c["alt_info"]["colunas"]:
            w("- %s: %s" % (c["informe"], " | ".join(c["alt_info"]["colunas"])))
    w("")
    w("Como cada variante foi tratada: o caminho 1 nao depende de posicao fixa. Ele acha o cabecalho da "
      "matriz pela grade, identifica cada coluna numerica pelo texto do cabecalho (`2019`, `EOP 2024`, "
      "`EOP 0`) e pela faixa PHYSICAL/FINANCIAL acima dela, aceita os rotulos `P(a)` e `P (a)`, e le o "
      "codigo do produto tanto no formato de 2 niveis (`1.1`) quanto no de 3 (`1.1.1`). O codigo interno "
      "do PMR (`1.1` ou `1.01`) e normalizado para numeros (`(1, 1)`) antes de comparar.")
    w("")
    # ---- correspondencia de codigos
    w("## 4. Codigo contratual nos informes que so imprimem 2 niveis")
    w("")
    so2 = [c["informe"] for c in resumo_inf if c["variante"] in
           {v for sig, v in variantes.items() if sig[5] == "2"}]
    w("Nos informes %s, o nome do produto comeca pelo codigo de 2 niveis (`1.1`, ..., `3.4`), igual ao "
      "codigo interno do PMR. O codigo de 3 niveis foi atribuido pela correspondencia codigo interno -> "
      "codigo do nome, que os informes %s a %s imprimem lado a lado. A correspondencia foi verificada: "
      "cada codigo interno levou sempre ao mesmo codigo de 3 niveis (%d de %d, nenhum conflito)."
      % (", ".join(so2), fontes_corresp[0], fontes_corresp[-1], len(corresp), len(CODIGOS_CONTRATO)))
    w("")
    linhas = []
    for k in sorted(corresp):
        cod = corresp[k]
        hist = nomes_por_codigo.get(cod, [])
        primeiro = next((n for i, n, u in hist if i == resumo_inf[0]["informe"]), "")
        ref = next((n for i, n, u in hist if i == fontes_corresp[0]), "")
        linhas.append([".".join(str(x) for x in k), cod, primeiro, ref])
    w(md_tabela(["Codigo interno PMR", "Codigo atribuido", "Nome em %s" % resumo_inf[0]["informe"],
                 "Nome em %s" % fontes_corresp[0]], linhas))
    w("")
    # ---- mudancas de unidade e de nome
    w("## 5. Mudancas de unidade de medida entre informes")
    w("")
    w("Comparacao do texto impresso na coluna \"Unit of Measure\" de um informe para o seguinte. Inclui a "
      "troca de idioma (espanhol -> portugues), que tambem muda o texto.")
    w("")
    linhas = []
    for cod in CODIGOS_CONTRATO:
        hist = nomes_por_codigo.get(cod, [])
        for (i0, n0, u0), (i1, n1, u1) in zip(hist, hist[1:]):
            if u0 != u1:
                linhas.append([cod, "%s -> %s" % (i0, i1), u0, u1])
    w(md_tabela(["Produto", "Entre", "Unidade antes", "Unidade depois"], linhas) if linhas else "Nenhuma.")
    w("")
    w("Sequencia completa das unidades por produto (informe: unidade), so para os que mudaram:")
    w("")
    for cod in CODIGOS_CONTRATO:
        hist = nomes_por_codigo.get(cod, [])
        if len({u for _, _, u in hist}) > 1:
            blocos = []
            for i, n, u in hist:
                if not blocos or blocos[-1][1] != u:
                    blocos.append([i, u, i])
                else:
                    blocos[-1][2] = i
            w("- %s: %s" % (cod, "; ".join("%s%s = `%s`" % (a, (" a " + b) if b != a else "", u)
                                          for a, u, b in blocos)))
    w("")
    w("## 6. Mudancas no nome impresso do produto")
    w("")
    linhas = []
    for cod in CODIGOS_CONTRATO:
        hist = nomes_por_codigo.get(cod, [])
        for (i0, n0, u0), (i1, n1, u1) in zip(hist, hist[1:]):
            if n0 != n1:
                linhas.append([cod, "%s -> %s" % (i0, i1), n0, n1])
    w(md_tabela(["Produto", "Entre", "Nome antes", "Nome depois"], linhas) if linhas else "Nenhuma.")
    w("")
    # ---- caminho 2
    w("## 7. Conferencia completa pelo caminho 2 (coordenadas das palavras)")
    w("")
    w("O caminho 2 le cada pagina da matriz sem usar a grade: acha a linha do cabecalho (`Unit of Measure`), "
      "as posicoes x das colunas numericas, os rotulos P / P(a) / A pela posicao, e atribui cada numero a "
      "coluna pela borda direita (os numeros sao alinhados a direita). Depois compara, para cada produto, "
      "a unidade e cada celula de valor com o caminho 1. Os dois caminhos leem a mesma camada de texto do "
      "PDF, mas inferem a estrutura da tabela de modos diferentes (linhas de grade x posicao).")
    w("")
    linhas = []
    for doc, cab, produtos, outros, conf2, *_ in informes:
        linhas.append([cab["informe"], conf2["celulas"], len(conf2["divergencias"]), len(conf2["avisos"])])
    w(md_tabela(["Informe", "Comparacoes", "Divergencias", "Avisos do caminho 2"], linhas))
    w("")
    todas = [(x[1]["informe"],) + d for x in informes for d in x[4]["divergencias"]]
    if todas:
        w("Divergencias:")
        w("")
        w(md_tabela(["Informe", "Produto", "Campo", "Caminho 1 (grade)", "Caminho 2 (coordenadas)"],
                    [list(t) for t in todas]))
        w("")
    avs = [(x[1]["informe"], a) for x in informes for a in x[4]["avisos"]]
    if avs:
        w("Avisos do caminho 2:")
        w("")
        for i, a in avs:
            w("- %s: %s" % (i, a))
        w("")
    w("### As conferencias mordem? (autoteste por mutacao)")
    w("")
    w("Zero divergencia so prova alguma coisa se a conferencia for capaz de reprovar. Para cada variante de "
      "layout, o script pega a primeira pagina da matriz, estraga de proposito a leitura do caminho 1 (sem "
      "tocar no PDF) e verifica se os caminhos 2 e 3 acusam. \"sim\" = o erro plantado foi acusado (ou, na "
      "linha de controle, nada foi acusado sem mutacao).")
    w("")
    w(md_tabela(["Variante", "Informe", "Pag.", "Mutacao", "Acusou?"],
                [[v, i, p, d, "sim" if ok else "**NAO**"] for v, i, p, d, ok in testes]))
    w("")
    # ---- caminho 3
    w("## 8. Conferencia amostral pelo caminho 3 (texto bruto + imagem)")
    w("")
    w("Sorteio com semente fixa (`%s` + informe), 3 produtos por informe. Para cada um, o script le o "
      "texto bruto da pagina na ordem em que o PDF o guarda (`get_text(\"text\")`): acha a linha do nome, "
      "toma a linha anterior ao primeiro `P` como unidade, e junta os numeros que seguem cada rotulo. "
      "Compara com a leitura pela grade (so as celulas nao vazias, na ordem das colunas). O PNG de cada "
      "informe, em `conferencia_png\\`, mostra o cabecalho da matriz e as linhas dos 3 produtos sorteados, "
      "recortados da pagina original." % SEMENTE_AMOSTRA)
    w("")
    w(md_tabela(["Informe", "Produto", "Pag.", "Resultado", "Texto bruto (P / P(a) / A, na ordem das colunas)",
                 "No CSV (fim de projeto)"],
                [[a["informe"], a["codigo"], a["pagina"], a["resultado"] + (" (%s)" % a["detalhe"] if a["detalhe"] else ""),
                  a["texto_bruto"], a["csv"]] for a in amostra]))
    w("")
    w("### Inspecao visual dos PNGs")
    w("")
    if CONFERENCIA_VISUAL:
        w("Feita a olho por quem rodou o script%s: cada PNG foi aberto e os numeros da imagem comparados "
          "com a coluna \"No CSV\" acima." % ((" em " + CONFERENCIA_VISUAL_DATA) if CONFERENCIA_VISUAL_DATA else ""))
        w("")
        for c in resumo_inf:
            w("- %s (`%s`): %s" % (c["informe"], c.get("png", ""), CONFERENCIA_VISUAL.get(c["informe"], "nao inspecionado")))
    else:
        w("Ainda nao registrada.")
    w("")
    # ---- vazios e duvidas
    w("## 9. Celulas vazias e observacoes da tabela mestra")
    w("")
    sem_ano = sorted({l["informe"] for l in mestra if not l["rotulo_coluna_ano"]})
    w("Informes cuja matriz NAO tem a coluna do ano (so a de fim de projeto): %s. Nesses, as colunas "
      "`*_ano_*` ficam vazias porque a coluna nao existe no PDF, e nao entram na contagem abaixo."
      % (", ".join(sem_ano) or "nenhum"))
    w("")
    cont = {}
    for l in mestra:
        for c in cols_fim + cols_ano:
            if c in cols_ano and not l["rotulo_coluna_ano"]:
                continue
            v = l.get(c, "")
            k = "vazia" if v == "" else ("traco" if v == "-" else None)
            if k:
                cont.setdefault((c, k), []).append(l["informe"])
    linhas = []
    for c in cols_fim + cols_ano:
        for k in ("vazia", "traco"):
            inf = cont.get((c, k), [])
            if inf:
                por = {}
                for i in inf:
                    por[i] = por.get(i, 0) + 1
                linhas.append([c, k, len(inf), "; ".join("%s: %d" % kv for kv in por.items())])
    w(md_tabela(["Coluna", "Situacao", "Celulas", "Por informe"], linhas) if linhas else "Nenhuma.")
    w("")
    w("Observacoes distintas registradas na coluna `obs` (com quantas linhas tem cada uma):")
    w("")
    tipos = {}
    for l in mestra:
        for o in [x for x in l["obs"].split(" | ") if x]:
            chave = re.sub(r"'[^']*'", "'...'", o)
            tipos[chave] = tipos.get(chave, 0) + 1
    for o, n in sorted(tipos.items(), key=lambda kv: -kv[1]):
        w("- (%d) %s" % (n, o))
    w("")
    outros = [(x[1]["informe"], o) for x in informes for o in x[3]]
    w("Linhas da matriz que nao sao produto, achadas dentro da tabela dos produtos (ignoradas na tabela "
      "mestra): " + ("; ".join(sorted({o for _, o in outros})) if outros else
                     "nenhuma (\"Other Cost\" e \"Total Cost\" ficam em tabelas proprias, que o script nao le)") + ".")
    w("")
    dec = [(l["informe"], l["codigo"], c, l[c]) for l in mestra for c in cols_fim + cols_ano
           if "," in l.get(c, "")]
    w("Valores impressos com casas decimais (gravados no CSV com virgula decimal): " + (
        "; ".join("%s %s %s = %s" % t for t in dec) if dec else "nenhum") + ".")
    w("")
    # ---- mudancas do P (planejado original) de um informe para o seguinte
    w("## 9b. Mudancas no valor P (planejado) de fim de projeto entre informes")
    w("")
    w("Registro factual, sem interpretacao: casos em que a linha P de fim de projeto muda de um informe "
      "para o seguinte. (P(a) e A mudam por natureza e estao na tabela mestra.)")
    w("")
    linhas = []
    por_cod = {}
    for l in mestra:
        por_cod.setdefault(l["codigo"], []).append(l)
    for cod in CODIGOS_CONTRATO:
        hist = por_cod.get(cod, [])
        for a, b in zip(hist, hist[1:]):
            for c in ("qtd_fim_P", "custo_fim_P"):
                if a.get(c, "") != b.get(c, ""):
                    linhas.append([cod, c, "%s -> %s" % (a["informe"], b["informe"]), a.get(c, ""), b.get(c, "")])
    w(md_tabela(["Produto", "Coluna", "Entre", "Antes", "Depois"], linhas) if linhas else "Nenhuma.")
    w("")
    # ---- alteracoes
    w("## 10. A tabela de alteracoes (\"CHANGES TO THE MATRIX\")")
    w("")
    linhas = []
    for c in resumo_inf:
        ia = c["alt_info"]
        if not c["alt_n"]:
            linhas.append([c["informe"], "sem informacao" if ia["sem_informacao"] else "nao achada", "", "", "", ""])
            continue
        sub = [a for a in alt_saida if a["informe"] == c["informe"]]
        prod = sum(1 for a in sub if a["codigo"])
        linhas.append([c["informe"], c["alt_n"], ia.get("datas_texto_bruto", ""), prod,
                       "sim" if c["alt_tem_motivo"] else "nao",
                       c["alt_vazios"] if c["alt_tem_motivo"] else "n.a."])
    w(md_tabela(["Informe", "Linhas lidas (grade)", "Datas no texto bruto da pagina", "Linhas de produto",
                 "Coluna de motivo existe", "Motivos vazios"], linhas))
    w("")
    w("Conferencia da contagem: cada linha de alteracao tem exatamente uma data em \"Entered in System\"; "
      "o numero de datas no texto bruto da pagina tem de ser igual ao numero de linhas lidas pela grade.")
    w("")
    w("- Nenhuma das tabelas tem coluna de valor anterior ou de valor novo: as colunas `valor_anterior_impresso` "
      "e `valor_novo_impresso` ficam vazias em todas as linhas, e a `obs` diz isso.")
    w("- Quando o texto do motivo cita numeros (\"de 4 para 5\"), eles vao para `valores_citados_no_motivo`.")
    w("- `valor_anterior_comparado` / `valor_novo_comparado` sao CALCULO do script, nao impressao: para "
      "alteracoes de produto cujo subtipo e \"Physical EOP P(a)\" ou \"Financial EOP P(a)\", trazem o "
      "P(a) de fim de projeto do informe anterior e o deste informe, lidos na tabela mestra.")
    w("- `o_que_mudou` e a traducao literal do subtipo impresso.")
    w("- Celulas mescladas no PDF (o mesmo item, ou o mesmo tipo, para varias linhas) tem o texto herdado da "
      "linha de cima; a `obs` registra quais.")
    w("")
    sem_mod = [a for a in alt_saida if not a["modificado_por"]]
    if sem_mod:
        w("Linhas sem \"Modified By\" impresso: " + "; ".join(
            "%s, ordem %s (%s)" % (a["informe"], a["ordem"], a["item"]) for a in sem_mod) + ".")
        w("")
    nao_prod = [a for a in alt_saida if not a["codigo"]]
    w("Linhas que nao sao de produto (indicadores, marcos e outros): %d." % len(nao_prod))
    for a in nao_prod:
        w("- %s, ordem %s: secao `%s`, item `%s`, %s" % (a["informe"], a["ordem"], a["secao"], a["item"], a["subtipo"]))
    w("")
    # ---- cruzamento: mudancas de P(a) de fim de projeto x registros na tabela de alteracoes
    w("## 11. Cruzamento: mudanca de P(a) de fim de projeto x registro na tabela de alteracoes")
    w("")
    w("Para cada informe, o script compara o P(a) de fim de projeto (quantidade e custo) de cada produto com "
      "o do informe imediatamente anterior, e procura, na tabela de alteracoes DESTE informe, uma linha do "
      "mesmo produto com subtipo \"Physical EOP P(a)\" (quantidade) ou \"Financial EOP P(a)\" (custo). "
      "Registro factual dos dois desencontros possiveis; nao diz por que ocorreram.")
    w("")
    ordem_inf = [c["informe"] for c in resumo_inf]
    idx = {(l["informe"], l["codigo"]): l for l in mestra}
    reg = set()
    for a in alt_saida:
        if a["codigo"] and "Physical EOP P(a)" in a["subtipo"]:
            reg.add((a["informe"], a["codigo"], "qtd_fim_Pa"))
        if a["codigo"] and "Financial EOP P(a)" in a["subtipo"]:
            reg.add((a["informe"], a["codigo"], "custo_fim_Pa"))
    sem_registro, sem_diferenca = [], []
    for i in range(1, len(ordem_inf)):
        ant, atu = ordem_inf[i - 1], ordem_inf[i]
        for cod in CODIGOS_CONTRATO:
            la, lc = idx.get((ant, cod)), idx.get((atu, cod))
            if not la or not lc:
                continue
            for campo in ("qtd_fim_Pa", "custo_fim_Pa"):
                mudou = la.get(campo, "") != lc.get(campo, "")
                registrado = (atu, cod, campo) in reg
                if mudou and not registrado:
                    sem_registro.append([atu, ant, cod, campo, la.get(campo, ""), lc.get(campo, "")])
                if registrado and not mudou:
                    sem_diferenca.append([atu, ant, cod, campo, la.get(campo, ""), lc.get(campo, "")])
    prim = [x for x in sem_registro if x[4] == "0" and x[0] == "2018-S2"]
    resto = [x for x in sem_registro if x not in prim]
    w("**(a) P(a) mudou e a tabela de alteracoes do informe nao tem linha correspondente: %d casos.**"
      % len(sem_registro))
    w("")
    if prim:
        w("- %d deles sao a passagem 2018-S1 -> 2018-S2, em que o P(a) de fim de projeto impresso deixa de ser "
          "`0` em todos os produtos e passa a ter valor (a tabela de alteracoes de 2018-S2 imprime \"No "
          "information available for this section\"). Lista completa no bloco recolhivel ao fim desta parte (a)." % len(prim))
    if resto:
        w("- Os outros %d:" % len(resto))
        w("")
        w(md_tabela(["Informe", "Comparado com", "Produto", "Campo", "Antes", "Depois"], resto))
    w("")
    if prim:
        w("<details><summary>Os %d casos de 2018-S1 -> 2018-S2</summary>" % len(prim))
        w("")
        w(md_tabela(["Informe", "Comparado com", "Produto", "Campo", "Antes", "Depois"], prim))
        w("")
        w("</details>")
        w("")
    w("**(b) A tabela de alteracoes registra mudanca de P(a) de fim de projeto e o valor e igual ao do informe "
      "anterior: %d casos.**" % len(sem_diferenca))
    w("")
    if sem_diferenca:
        w(md_tabela(["Informe", "Comparado com", "Produto", "Campo", "Antes", "Depois"], sem_diferenca))
        w("")
    # ---- limites
    w("## 12. Limites desta conferencia")
    w("")
    w("- Os caminhos 1, 2 e 3 leem a mesma camada de texto do PDF (PyMuPDF). Se o texto embutido no PDF "
      "diferisse do que aparece na imagem da pagina, so a inspecao visual dos PNGs (secao 8) pegaria; ela "
      "cobre 48 produtos, nao os 384.")
    w("- O caminho 3 compara a sequencia de valores NAO vazios. Ele nao distingue, sozinho, em que coluna um "
      "valor caiu quando ha celulas em branco na mesma linha (caso de 2018-S1, linhas P(a) e A). Quem "
      "garante a coluna e o caminho 2 (posicao x de cada numero), conferido em todos os 384 produtos, e a "
      "inspecao visual.")
    w("- O script nao le as tabelas \"Other Cost\" e \"Total Cost\", nem os indicadores de impacto e de "
      "resultado (paginas 2 a 4): estao fora do pedido.")
    w("- O codigo de 3 niveis de 2017-S2, 2018-S1 e 2018-S2 e atribuido por correspondencia (secao 4), nao "
      "lido; a coluna `obs` de cada uma dessas 72 linhas diz isso.")
    w("")
    with open(os.path.join(args.saida, "conferencia_amostral.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(L) + "\n")


if __name__ == "__main__":
    main()
