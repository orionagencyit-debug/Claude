"""Scarica i dataset Apify (Google Maps), tiene solo cellulari, deduplica e crea la base xlsx.
Uso: python3 process_apify.py runs.txt out.xlsx
runs.txt: righe «Provincia runId datasetId stato» (il token Apify è iniettato dal proxy)."""
import json, re, sys, urllib.request
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

NICHE_ORDER = ["imbianchino", "cartongesso", "ristrutturazioni", "ristrutturazione bagni", "muratore",
               "posa pavimenti", "cappotto termico", "impermeabilizzazioni"]
FIELDS = "title,phone,phoneUnformatted,city,state,address,categoryName,totalScore,reviewsCount,url,placeId,searchString,permanentlyClosed,temporarilyClosed,website"
BLOCK = 400
OWNERS = ["Егор", "Егор", "Setter 1", "Setter 2", "Setter 3", "Setter 4"]  # блоки 1–6, дальше — резерв


def fetch(ds):
    url = f"https://api.apify.com/v2/datasets/{ds}/items?format=json&clean=true&fields={FIELDS}"
    with urllib.request.urlopen(url, timeout=120) as r:
        return json.load(r)


def norm_phone(p):
    d = re.sub(r"[^0-9+]", "", str(p or ""))
    if d.startswith("0039"):
        d = "+39" + d[4:]
    if not d.startswith("+"):
        d = "+39" + d
    return d


def is_mobile(p):
    return p.startswith("+393")


def pretty(p):
    n = p[3:]
    return f"{n[:3]} {n[3:6]} {n[6:]}".strip()


def main(runs_file, out):
    rows, seen_phone, seen_place, stats = [], set(), set(), {}
    for line in open(runs_file):
        parts = line.split()
        if len(parts) < 3:
            continue
        prov, ds = parts[0], parts[2]
        items = fetch(ds)
        st = stats.setdefault(prov, {"scraped": len(items), "mobile": 0, "kept": 0})
        for it in items:
            if it.get("permanentlyClosed") or it.get("temporarilyClosed") or it.get("website"):
                continue
            ph = norm_phone(it.get("phoneUnformatted") or it.get("phone"))
            if not is_mobile(ph):
                continue
            st["mobile"] += 1
            pid = it.get("placeId")
            if ph in seen_phone or (pid and pid in seen_place):
                continue
            seen_phone.add(ph)
            if pid:
                seen_place.add(pid)
            st["kept"] += 1
            rows.append({
                "name": it.get("title"), "phone": pretty(ph), "city": it.get("city") or "",
                "prov": prov, "niche": it.get("searchString") or "", "cat": it.get("categoryName") or "",
                "reviews": it.get("reviewsCount") or 0, "rating": it.get("totalScore") or 0,
                "maps": it.get("url") or "",
            })
    # активность: больше отзывов и выше рейтинг → выше в списке
    rows.sort(key=lambda r: (-min(r["reviews"], 60), -r["rating"],
                             NICHE_ORDER.index(r["niche"]) if r["niche"] in NICHE_ORDER else 99))
    for i, r in enumerate(rows):
        b = i // BLOCK + 1
        r["block"] = b
        r["owner"] = OWNERS[b - 1] if b <= len(OWNERS) else "Резерв"

    wb = Workbook()
    ws = wb.active
    ws.title = "База"
    cols = ["Название", "Телефон", "Статус", "Дата звонка", "Заметки", "Город", "Провинция", "Ниша",
            "Категория Google", "Отзывов", "Рейтинг", "Google Maps", "Блок", "Кто звонит"]
    dark = PatternFill("solid", fgColor="12355B")
    inp = PatternFill("solid", fgColor="FFF4C2")
    thin = Side(style="thin", color="B9C4D0")
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.fill, c.font = dark, Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(bottom=thin)
    for k, r in enumerate(rows, 2):
        vals = [r["name"], r["phone"], "", None, "", r["city"], r["prov"], r["niche"], r["cat"],
                r["reviews"], r["rating"], r["maps"], r["block"], r["owner"]]
        for j, v in enumerate(vals, 1):
            ws.cell(row=k, column=j, value=v)
        for j in (3, 4, 5):
            ws.cell(row=k, column=j).fill = inp
        ws.cell(row=k, column=4).number_format = "dd.mm"
    for col, w in zip("ABCDEFGHIJKLMN", [34, 15, 15, 11, 30, 18, 11, 20, 24, 9, 8, 14, 7, 11]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"

    sm = wb.create_sheet("Сводка")
    sm.append(["Провинция", "Найдено без сайта", "Из них мобильных", "В базе (после дублей)"])
    for p, s in stats.items():
        sm.append([p, s["scraped"], s["mobile"], s["kept"]])
    sm.append(["ИТОГО", sum(s["scraped"] for s in stats.values()), sum(s["mobile"] for s in stats.values()), len(rows)])
    sm.append([])
    sm.append(["Ниша", "Мобильных в базе"])
    for n in NICHE_ORDER:
        sm.append([n, sum(1 for r in rows if r["niche"] == n)])
    sm.append([])
    sm.append(["Блок", "Кто звонит", "Номеров"])
    for b in sorted({r["block"] for r in rows}):
        sm.append([b, OWNERS[b - 1] if b <= len(OWNERS) else "Резерв", sum(1 for r in rows if r["block"] == b)])
    for col, w in zip("ABCD", [22, 20, 20, 22]):
        sm.column_dimensions[col].width = w
    wb.save(out)
    print(json.dumps({"rows": len(rows), "stats": stats}, ensure_ascii=False))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
