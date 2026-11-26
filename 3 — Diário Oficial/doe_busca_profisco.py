import json, re, time, html, sys, urllib.request, urllib.parse
BASE = "https://diariooficial.ma.gov.br/ajax.busca.php?"
termo, saida = sys.argv[1], sys.argv[2]
params = {"termo": termo, "datai": "2019-01-01", "dataf": "2025-12-31"}
frags, scroll, total = [], "", None
for volta in range(20):
    q = dict(params)
    if scroll: q["scrollId"] = scroll
    req = urllib.request.Request(BASE + urllib.parse.urlencode(q), headers={"User-Agent": "Mozilla/5.0"})
    o = json.loads(urllib.request.urlopen(req, timeout=60).read().decode("utf-8"))["busca"]
    total = o.get("total")
    cards = re.split(r'<div class="card flex-md-row', o.get("es_html", ""))[1:]
    for c in cards:
        cad = re.search(r'<strong class="d-inline-block mb-2 text-primary">([^<]+)', c)
        data = re.search(r'Publicado em ([0-9/]+)', c)
        ed = re.search(r"setModal\('[^']*','[^']*','([A-Z]+\d{8})'", c)
        pag = re.search(r'P&aacute;gina: </strong>(\d+)', c)
        txt = " ".join(html.unescape(re.sub(r"<[^>]+>", " ", t)) for t in re.findall(r'<p class="card-text mb-auto">(.*?)</p>', c, re.S) if "gina:" not in t)
        frags.append({"caderno": cad and cad.group(1), "data": data and data.group(1), "edicao": ed and ed.group(1),
                      "pagina": pag and pag.group(1), "texto": re.sub(r"\s+", " ", txt).strip()})
    print(f"volta {volta}: +{len(cards)} (acumulado {len(frags)} de {total})", flush=True)
    scroll = o.get("scrollId")
    if not cards or not o.get("btnLoad") or len(frags) >= (total or 0): break
    time.sleep(20)
json.dump({"termo": termo, "periodo": "2019-01-01 a 2025-12-31", "consultado_em": time.strftime("%Y-%m-%d %H:%M"),
           "url": BASE + urllib.parse.urlencode(params), "total_declarado": total, "fragmentos": frags},
          open(saida, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("gravado", saida, len(frags))
