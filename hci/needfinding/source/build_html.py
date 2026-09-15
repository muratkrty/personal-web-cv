"""Inject the shared questionnaire (and, for vis.html, the sample dataset) into the HTML templates."""
import json, csv
Q = json.load(open("questions.json"))
qjson = json.dumps(Q, ensure_ascii=False, separators=(",", ":"))
u = open("templates/user.html").read().replace("/*__QUESTIONS__*/", qjson)
open("out/user.html", "w").write(u)
import os
if os.path.exists("templates/vis.html"):
    rows = list(csv.DictReader(open("out/responses_final.csv")))
    for r in rows:
        for k, v in list(r.items()):
            if k in ("id", "submitted") or k.startswith(("A1_", "A2_", "A3_", "A4_", "A5_", "A6_", "A7_", "B1_", "B3_", "B4_", "B7_", "B8_", "E", "Segment")): continue
            if k in ("PC1", "PC2"): r.pop(k); continue
            if v == "": continue
            try: r[k] = int(v)
            except ValueError:
                try: r[k] = float(v)
                except ValueError: pass
        r.pop("Segment", None)
    tpl = open("templates/vis.html").read().replace("/*__QUESTIONS__*/", qjson).replace("/*__DATA__*/", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
    open("out/vis.html", "w").write(tpl.replace("/*__TITLE__*/", "iCanvas Needfinding Analysis"))
print("built")
