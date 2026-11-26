import json, re, time, html, urllib.request, urllib.parse
BASE = "https://diariooficial.ma.gov.br/ajax.busca.php?"
TERMOS = {
 "2.3.2": "REDESIM", "2.3.1": "obrigações acessórias", "2.1.1": "grandes contribuintes",
 "2.6.1": "SEFAZNET", "1.1.1": "sistemas corporativos", "3.1.2": "Planejamento e Gestão Fiscal",
 "3.2.2": "custo setorial", "1.2.1": "gestão estratégica de pessoas", "1.2.2": "competências fazendárias",
 "2.4.2": "contencioso judicial", "1.3.1": "tecnologia da informação PROFISCO", "2.2.1": "ações fiscais PROFISCO",
}
PROG = re.compile(r"PROFISCO|Banco Interamericano|\bBID\b|4458", re.I)
out = {}
for cod, termo in TERMOS.items():
    frags, scroll = [], ""
    for volta in range(3):
        q = {"termo": termo, "datai": "2019-01-01", "dataf": "2025-12-31"}
        if scroll: q["scrollId"] = scroll
        try:
            o = json.loads(urllib.request.urlopen(urllib.request.Request(BASE + urllib.parse.urlencode(q), headers={"User-Agent": "Mozilla/5.0"}), timeout=60).read().decode("utf-8"))["busca"]
        except Exception as e:
            print(cod, termo, "ERRO", e); break
        cards = re.split(r'<div class="card flex-md-row', o.get("es_html", ""))[1:]
        for c in cards:
            ed = re.search(r"setModal\('[^']*','[^']*','([A-Z]+\d{8})'", c)
            pag = re.search(r'P&aacute;gina: </strong>(\d+)', c)
            txt = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", t)) for t in re.findall(r'<p class="card-text mb-auto">(.*?)</p>', c, re.S) if "gina:" not in t)
            frags.append({"edicao": ed and ed.group(1), "pagina": pag and pag.group(1), "texto": re.sub(r"\s+", " ", txt).strip()})
        total = o.get("total"); scroll = o.get("scrollId")
        time.sleep(20)
        if not cards or not o.get("btnLoad") or len(frags) >= (total or 0): break
    prog = [f for f in frags if PROG.search(f["texto"])]
    out[cod] = {"termo": termo, "total_declarado": total, "lidos": len(frags), "com_programa": prog}
    print(f"{cod} '{termo}': total {total}, lidos {len(frags)}, com PROFISCO/BID no trecho: {len(prog)}", flush=True)
    for f in prog[:3]: print("    ", f["edicao"], "p."+str(f["pagina"]), f["texto"][:260])
json.dump({"consultado_em": time.strftime("%Y-%m-%d %H:%M"), "periodo": "2019-01-01 a 2025-12-31", "termos": out},
          open("doe/doe_termos_por_produto.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
