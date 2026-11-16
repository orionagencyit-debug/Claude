"""Calcola le metriche giornaliere dall'export della tabella lead (CSV o XLSX).

Uso:  python3 import_day.py <export.csv|xlsx> <YYYY-MM-DD>

- Chiamate / numeri inesistenti / conversazioni: righe con "Дата звонка" = giorno.
- Demo fissate, demo svolte, vendite, euro: cambi di stato rispetto all'ultimo
  snapshot salvato in snapshots/ (al primo avvio contano solo le righe del giorno).
Stampa un JSON con i valori per il foglio «Дневник» e salva lo snapshot del giorno.
"""
import csv, datetime as dt, hashlib, json, re, sys
from pathlib import Path

SNAP = Path(__file__).with_name("snapshots")

DEAD = {"inesistente"}
NO_ANSWER = {"non risponde"}
CALLBACK = {"da chiamare"}
TALK = {"no", "c'è sito", "gia in lavoro", "demo", "demo fatta", "attesa acconto", "pagato"}
DEMO_BOOKED = {"demo", "demo fatta", "attesa acconto", "pagato"}
DEMO_DONE = {"demo fatta", "attesa acconto", "pagato"}
PAID = {"pagato"}


def norm_status(s):
    s = (s or "").strip().lower().replace("’", "'")
    s = re.sub(r"\s+", " ", s)
    if s.startswith(("ineist", "inesist")):
        return "inesistente"
    if s.startswith("non ris"):
        return "non risponde"
    if s.startswith(("da chi", "da chiam", "da ric")):
        return "da chiamare"
    if s.startswith(("c'è sit", "ce sito", "c'e sit")):
        return "c'è sito"
    if s.startswith(("gia in", "già in")):
        return "gia in lavoro"
    if s in ("sold", "attesa", "attesa acconto"):
        return "attesa acconto"
    if s.startswith("demo fat"):
        return "demo fatta"
    if s.startswith("demo"):
        return "demo"
    if s.startswith(("pagat", "paid")):
        return "pagato"
    return s


def parse_date(v, year):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.search(r"(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?", str(v or ""))
    if not m:
        return None
    d, mo, y = int(m[1]), int(m[2]), m[3]
    y = int(y) + (2000 if y and len(y) == 2 else 0) if y else year
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def money(v):
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def read_rows(path):
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        ws = load_workbook(path, data_only=True).worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    header = [str(h or "").strip() for h in rows[0]]
    return [dict(zip(header, r)) for r in rows[1:] if any(c not in (None, "") for c in r)], header


def find_col(header, *names):
    low = {h.lower(): h for h in header}
    for n in names:
        for k, h in low.items():
            if n in k:
                return h
    return None


def main(path, day):
    day = dt.date.fromisoformat(day)
    rows, header = read_rows(Path(path))
    c_name = find_col(header, "назван", "nome") or header[0]
    c_phone = find_col(header, "телефон", "telefono")
    c_status = find_col(header, "call", "статус", "stato")
    c_date = find_col(header, "дата звонка", "data chiamata", "дата")
    c_pay = find_col(header, "оплата", "pagamento", "€")
    if not (c_status and c_date):
        sys.exit(f"Colonne non trovate: status={c_status} data={c_date}. Header: {header}")

    cur = {}
    for r in rows:
        # chiave anonima: nello snapshot non finiscono nomi né telefoni
        key = hashlib.sha1(f"{r.get(c_name)}|{r.get(c_phone)}".encode()).hexdigest()[:16]
        cur[key] = {"status": norm_status(r.get(c_status)),
                    "date": parse_date(r.get(c_date), day.year),
                    "pay": money(r.get(c_pay)) if c_pay else 0.0}

    prev_files = sorted(p for p in SNAP.glob("*.json") if p.stem < day.isoformat())
    prev = json.loads(prev_files[-1].read_text()) if prev_files else None

    today = [v for v in cur.values() if v["date"] == day]
    out = {"date": day.isoformat(), "calls": len(today),
           "dead": sum(v["status"] in DEAD for v in today),
           "talks": sum(v["status"] in TALK for v in today),
           "no_answer": sum(v["status"] in NO_ANSWER for v in today),
           "callbacks": sum(v["status"] in CALLBACK for v in today)}

    def events(cond_now, cond_prev):
        n = 0
        for k, v in cur.items():
            if prev is None:
                p_ok = v["date"] != day  # primo avvio: lo stato precedente è sconosciuto
                if cond_now(v) and not p_ok:
                    n += 1
            elif cond_now(v) and not cond_prev(prev.get(k, {"status": "", "pay": 0})):
                n += 1
        return n

    out["demo_booked"] = events(lambda v: v["status"] in DEMO_BOOKED, lambda v: v["status"] in DEMO_BOOKED)
    out["demo_done"] = events(lambda v: v["status"] in DEMO_DONE, lambda v: v["status"] in DEMO_DONE)
    out["sales"] = events(lambda v: v["status"] in PAID, lambda v: v["status"] in PAID)
    if prev is None:
        out["euro"] = sum(v["pay"] for v in today)
    else:
        out["euro"] = sum(max(0.0, v["pay"] - prev.get(k, {}).get("pay", 0.0)) for k, v in cur.items())
    out["leads_total"] = len(cur)
    out["leads_never_called"] = sum(1 for v in cur.values() if not v["status"])
    out["warnings"] = [] if c_pay else ["Нет колонки «Оплата €» — деньги не считаются"]

    SNAP.mkdir(exist_ok=True)
    (SNAP / f"{day.isoformat()}.json").write_text(json.dumps(
        {k: {"status": v["status"], "pay": v["pay"]} for k, v in cur.items()}, ensure_ascii=False))
    print(json.dumps(out, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main(*sys.argv[1:3])
