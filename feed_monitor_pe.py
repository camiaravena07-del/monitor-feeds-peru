#!/usr/bin/env python3
"""
feed_monitor_pe.py — Monitor de feeds PERÚ (Grupo Axo)

Versión solo-disponibilidad del monitor de Chile. Perú no tiene sitios multimarca,
así que no se calculan cruces de precio ni cobertura entre feeds.

Genera 2 pestañas en la Sheet de Perú:
  1) historico (APPEND diario): fecha, marca, pmax, pmax2, total_productos,
                                disponibles, unidades_disponibles
  2) meta      (OVERWRITE):     ultima_actualizacion

Salida: si existe GOOGLE_SHEET_ID -> Google Sheet; si no -> CSV en ./output (pruebas).
OJO: en el workflow de Perú, GOOGLE_SHEET_ID se llena con el secret GOOGLE_SHEET_ID_PE.
"""

import csv
import os
import re
from datetime import date, datetime
from collections import defaultdict
import xml.etree.ElementTree as ET

FEEDS = {
    "Banana Republic": "https://resources.komaxchile.cl/google-feeds/banana-republic-pe-google-feed.xml",
    "GAP":             "https://resources.komaxchile.cl/google-feeds/gap-pe-google-feed.xml",
    "Guess":           "https://resources.komaxchile.cl/google-feeds/guess-pe-google-feed.xml",
    "Kipling":         "https://resources.komaxchile.cl/google-feeds/kipling-pe-google-feed.xml",
    "The North Face":  "https://resources.komaxchile.cl/google-feeds/the-north-face-pe-google-feed.xml",
    "Coach":           "https://resources.komaxchile.cl/google-feeds/coach-pe-google-feed.xml",
}

PMAX_FIELD  = "pmax"
PMAX2_FIELD = "pmax2"
AVAILABLE_VALUES = {"in stock", "in_stock"}
OUT_DIR = "output"
HIST_HEADER = ["fecha", "marca", "pmax", "pmax2", "total_productos", "disponibles", "unidades_disponibles"]


def _local(tag): return tag.split("}")[-1]


def parse_int(raw):
    """'290' -> 290. Devuelve 0 si no hay numero."""
    if not raw:
        return 0
    m = re.search(r"-?\d+", raw.replace(".", "").replace(",", ""))
    return int(m.group()) if m else 0


def iter_items(source):
    if source.startswith(("http://", "https://")):
        import urllib.request
        req = urllib.request.Request(source, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/xml,text/xml,application/rss+xml,*/*",
        })
        stream = urllib.request.urlopen(req, timeout=120)
    else:
        stream = open(source, "rb")
    try:
        for _ev, elem in ET.iterparse(stream, events=("end",)):
            if _local(elem.tag) != "item":
                continue
            yield {_local(c.tag): (c.text or "").strip() for c in elem}
            elem.clear()
    finally:
        stream.close()


def _ahora_peru():
    """Fecha y hora actual en horario de Perú (UTC-5, sin horario de verano)."""
    try:
        from zoneinfo import ZoneInfo
        dt = datetime.now(ZoneInfo("America/Lima"))
    except Exception:
        from datetime import timezone, timedelta
        dt = datetime.now(timezone(timedelta(hours=-5)))
    return dt.strftime("%d-%m-%Y %H:%M") + " hrs (hora Perú)"


def analyze(feeds):
    today = date.today().isoformat()
    avail = defaultdict(lambda: [0, 0, 0])   # [total_productos, disponibles, unidades_disponibles]

    for brand, source in feeds.items():
        try:
            for rec in iter_items(source):
                pmax  = rec.get(PMAX_FIELD)  or "(sin categoria)"
                pmax2 = rec.get(PMAX2_FIELD) or "(sin subcategoria)"
                cell = avail[(brand, pmax, pmax2)]
                cell[0] += 1
                if rec.get("availability", "").lower() in AVAILABLE_VALUES:
                    cell[1] += 1
                    cell[2] += parse_int(rec.get("qty"))
        except Exception as e:
            # un feed caido no debe botar la corrida completa
            print(f"  (aviso) no se pudo leer {brand}: {e}")

    historico = [[today, b, p, p2, tot, disp, uni]
                 for (b, p, p2), (tot, disp, uni) in sorted(avail.items())]
    return {
        "historico": {"header": HIST_HEADER, "rows": historico, "mode": "append"},
        "meta": {"header": ["ultima_actualizacion"], "rows": [[_ahora_peru()]], "mode": "overwrite"},
    }


def write_csv(results, out_dir=OUT_DIR):
    os.makedirs(out_dir, exist_ok=True)
    for name, t in results.items():
        path = os.path.join(out_dir, name + ".csv")
        mode = "a" if (t["mode"] == "append" and os.path.exists(path)) else "w"
        with open(path, mode, newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if mode == "w":
                w.writerow(t["header"])
            w.writerows(t["rows"])
    _summary(results)


def write_sheets(results, sheet_id):
    import gspread, json
    gc = gspread.service_account_from_dict(json.loads(os.environ["GCP_SA_KEY"]))
    sh = gc.open_by_key(sheet_id)
    existing = {ws.title for ws in sh.worksheets()}
    for name, t in results.items():
        if name not in existing:
            sh.add_worksheet(title=name, rows=1000, cols=max(20, len(t["header"])))
            existing.add(name)
        ws = sh.worksheet(name)
        if t["mode"] == "append":
            today = date.today().isoformat()
            vals = ws.get_all_values()
            if not vals:
                ws.update("A1", [t["header"]], value_input_option="RAW")
            else:
                if vals[0] != t["header"]:
                    vals[0] = t["header"]
                # se conservan las filas de dias anteriores; las de hoy se reemplazan (idempotente)
                kept = [vals[0]] + [r for r in vals[1:] if not r or r[0] != today]
                kept = [[_numify(c) for c in row] for row in kept]
                ws.clear()
                ws.update("A1", kept, value_input_option="RAW")
            if t["rows"]:
                ws.append_rows(t["rows"], value_input_option="RAW")
        else:
            ws.clear()
            ws.update("A1", [t["header"]] + t["rows"], value_input_option="RAW")
        _reset_number_format(ws)
    _summary(results)


def _numify(v):
    """'3481' -> 3481, para que Sheets lo guarde como numero y no como texto."""
    if isinstance(v, str) and re.fullmatch(r"-?\d+", v):
        return int(v)
    return v


def _reset_number_format(ws):
    """Resetea el formato de celdas a AUTOMATICO (evita el apostrofo de texto)."""
    try:
        sid = ws._properties["sheetId"]
        ws.spreadsheet.batch_update({"requests": [{
            "repeatCell": {
                "range": {"sheetId": sid},
                "cell": {"userEnteredFormat": {}},
                "fields": "userEnteredFormat.numberFormat"
            }
        }]})
    except Exception as e:
        print("  (aviso) no se pudo resetear formato de", ws.title, ":", e)


def _summary(results):
    rows = results["historico"]["rows"]
    print("[" + date.today().isoformat() + "] Peru")
    print("  historico: " + str(len(rows)) + " filas (marca x cat x subcat)")
    por_marca = defaultdict(lambda: [0, 0])
    for _f, b, _p, _p2, tot, disp, _u in rows:
        por_marca[b][0] += tot
        por_marca[b][1] += disp
    for b, (tot, disp) in sorted(por_marca.items()):
        print(f"    {b}: {disp} disponibles de {tot}")


def main(feeds=FEEDS):
    results = analyze(feeds)
    sid = os.environ.get("GOOGLE_SHEET_ID")
    if sid:
        write_sheets(results, sid)
    else:
        write_csv(results)


if __name__ == "__main__":
    main()
