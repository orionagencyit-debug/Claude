"""Genera la tabella per i setter (xlsx) da caricare su Google Sheets.
Validazioni (menu a tendina, caselle di spunta) vengono aggiunte via API dopo l'upload."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

DARK = PatternFill("solid", fgColor="12355B")
SOFT = PatternFill("solid", fgColor="E8EEF5")
INPUT = PatternFill("solid", fgColor="FFF4C2")
WB = Font(bold=True, color="FFFFFF")
B = Font(bold=True)
H = Font(bold=True, size=14, color="12355B")
thin = Side(style="thin", color="B9C4D0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
N_SETTERS = 4
LEAD_ROWS = 450
DEMO_ROWS = 200


def hdr(ws, row, cols):
    for i, h in enumerate(cols, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.fill, c.font, c.alignment, c.border = DARK, WB, CENTER, BOX


wb = Workbook()

# ---------- Istruzioni ----------
ins = wb.active
ins.title = "Istruzioni"
lines = [
    ("ISTRUZIONI PER I SETTER — progetto di 2 settimane", "h"),
    ("Obiettivo: fissare appuntamenti telefonici di 15 minuti tra Egor e il titolare. Non devi vendere nulla.", ""),
    ("", ""),
    ("1. Apri il TUO foglio (Setter 1, Setter 2…) e chiama i numeri dall'alto verso il basso.", ""),
    ("2. Dopo ogni chiamata compila: Stato (menu), Data chiamata, Note.", ""),
    ("3. Se il titolare accetta: scegli un orario LIBERO nel foglio «Slot» (al massimo il giorno dopo) e scrivi «Tuo nome — Attività».", ""),
    ("4. Aggiungi una riga nel foglio «Demo» e manda subito a Egor su WhatsApp il messaggio «DEMO ✅» (modello sotto).", ""),
    ("5. Egor scrive al cliente per confermare. Quando il cliente risponde «ok», Egor spunta «Confermato».", ""),
    ("6. A fine giornata scrivi nel foglio «Ore»: data, nome, ore lavorate.", ""),
    ("", ""),
    ("Pagamento: 5 €/ora + 10 € per appuntamento confermato dal cliente + 15 € se l'appuntamento si svolge. Pagamento ogni venerdì.", "b"),
    ("Riservatezza: i contatti sono di Egor. Non copiarli e non usarli per altro.", "b"),
    ("", ""),
    ("Modello messaggio a Egor:", "b"),
    ("DEMO ✅ · Setter: [nome] · Attività: [nome] — [città], [settore] · Titolare: [nome], tel. [numero] · Quando: [giorno] [ora] · Decide: da solo / con socio · Note: […]", ""),
    ("", ""),
    ("Se qualcuno ti richiama: «La chiamavo per conto di Egor, che realizza siti web per imprese della zona. Le lascio il suo numero: [numero].»", ""),
    ("", ""),
    ("— RU для Егора —", "b"),
    ("Сеттеры работают каждый в своём листе. Ты: подтверждаешь демо (галочка «Confermato»), отмечаешь «Svolto» и «Vendita», смотришь лист «Pagamenti» в пятницу.", ""),
    ("Имена сеттеров и ставки — в листе «Impostazioni». Базу номеров вставляешь блоками по ~400 в листы Setter 1–4 (колонки A–G).", ""),
]
for i, (t, kind) in enumerate(lines, 1):
    c = ins.cell(row=i, column=1, value=t)
    c.alignment = WRAP
    if kind == "h":
        c.font = H
    elif kind == "b":
        c.font = B
ins.column_dimensions["A"].width = 130

# ---------- Impostazioni ----------
st = wb.create_sheet("Impostazioni")
st["A1"] = "IMPOSTAZIONI / НАСТРОЙКИ"
st["A1"].font = H
rows = [
    ("Paga oraria, €", 5),
    ("Bonus appuntamento confermato, €", 10),
    ("Bonus appuntamento svolto, €", 15),
    ("Bonus vendita (facoltativo), €", 0),
]
for i, (k, v) in enumerate(rows):
    st.cell(row=3 + i, column=1, value=k).border = BOX
    c = st.cell(row=3 + i, column=2, value=v)
    c.fill, c.border, c.font, c.number_format = INPUT, BOX, B, '#,##0 "€"'
st["A9"] = "Nomi dei setter (modifica qui → si aggiornano ovunque)"
st["A9"].font = B
for i in range(N_SETTERS):
    c = st.cell(row=10 + i, column=1, value=f"Setter {i+1}")
    c.fill, c.border = INPUT, BOX
st.column_dimensions["A"].width = 46
st.column_dimensions["B"].width = 14

# ---------- Setter N ----------
lead_cols = ["Attività", "Telefono", "Città", "Settore", "Recensioni", "Valutazione", "Google Maps",
             "Stato", "Data chiamata", "Note", "Demo: giorno e ora"]
for s in range(1, N_SETTERS + 1):
    ws = wb.create_sheet(f"Setter {s}")
    hdr(ws, 1, lead_cols)
    ws.row_dimensions[1].height = 30
    for r in range(2, LEAD_ROWS + 2):
        ws.cell(row=r, column=9).number_format = "dd.mm"
        for c in (8, 9, 10, 11):
            ws.cell(row=r, column=c).fill = INPUT
    for col, w in zip("ABCDEFGHIJK", [32, 15, 16, 20, 10, 10, 14, 16, 12, 34, 18]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    ws.conditional_formatting.add(f"A2:K{LEAD_ROWS+1}", FormulaRule(formula=['$H2="Demo fissata"'], fill=PatternFill("solid", fgColor="C8E6C9")))
    ws.conditional_formatting.add(f"A2:K{LEAD_ROWS+1}", FormulaRule(formula=['OR($H2="No",$H2="Inesistente",$H2="C\'è sito",$H2="Già in lavoro")'], fill=PatternFill("solid", fgColor="EEEEEE")))

# ---------- Slot ----------
sl = wb.create_sheet("Slot")
sl["A1"] = "SLOT PER LE DEMO — scrivi «Tuo nome — Attività» in una cella libera (max 6 demo al giorno)"
sl["A1"].font = H
times = ["11:00", "11:30", "12:00", "12:30", "14:30", "15:00", "15:30", "16:00", "16:30"]
days = []
d = dt.date(2026, 11, 17)
while d <= dt.date(2026, 12, 5):
    if d.weekday() < 5 and d != dt.date(2026, 12, 8):
        days.append(d)
    d += dt.timedelta(days=1)
dn = ["Lun", "Mar", "Mer", "Gio", "Ven", "Sab", "Dom"]
sl.cell(row=3, column=1, value="Ora")
for j, day in enumerate(days, 2):
    sl.cell(row=3, column=j, value=f"{dn[day.weekday()]} {day:%d.%m}")
hdr(sl, 3, range(1, len(days) + 2))
for i, t in enumerate(times):
    r = 4 + i
    c = sl.cell(row=r, column=1, value=t)
    c.font, c.fill, c.border, c.alignment = B, SOFT, BOX, CENTER
    for j in range(2, len(days) + 2):
        cell = sl.cell(row=r, column=j)
        cell.border, cell.alignment = BOX, CENTER
    if t == "14:30":
        pass
sl.cell(row=4 + len(times), column=1, value="Demo nel giorno").font = B
for j in range(2, len(days) + 2):
    col = sl.cell(row=3, column=j).column_letter
    c = sl.cell(row=4 + len(times), column=j, value=f"=COUNTA({col}4:{col}{3+len(times)})")
    c.alignment, c.border = CENTER, BOX
last_col = sl.cell(row=3, column=len(days) + 1).column_letter
sl.conditional_formatting.add(f"B4:{last_col}{3+len(times)}", FormulaRule(formula=['LEN(B4)>0'], fill=PatternFill("solid", fgColor="FFE0B2")))
sl.conditional_formatting.add(f"B{4+len(times)}:{last_col}{4+len(times)}", FormulaRule(formula=[f'B{4+len(times)}>=6'], fill=PatternFill("solid", fgColor="FFCDD2")))
sl.column_dimensions["A"].width = 9
for j in range(2, len(days) + 2):
    sl.column_dimensions[sl.cell(row=3, column=j).column_letter].width = 16
sl.freeze_panes = "B4"

# ---------- Demo ----------
dm = wb.create_sheet("Demo")
demo_cols = ["Data inserimento", "Setter", "Attività", "Titolare", "Telefono", "Città", "Settore",
             "Giorno demo", "Ora", "Decide", "Confermato (Egor)", "Svolto (Egor)", "Vendita (Egor)",
             "Acconto €", "Note"]
hdr(dm, 1, demo_cols)
dm.row_dimensions[1].height = 32
for r in range(2, DEMO_ROWS + 2):
    dm.cell(row=r, column=1).number_format = "dd.mm"
    dm.cell(row=r, column=8).number_format = "dd.mm"
    dm.cell(row=r, column=14).number_format = '#,##0 "€"'
    for c in (11, 12, 13):
        dm.cell(row=r, column=c, value=False)
for col, w in zip("ABCDEFGHIJKLMNO", [11, 14, 28, 18, 15, 14, 18, 11, 8, 14, 12, 11, 11, 11, 30]):
    dm.column_dimensions[col].width = w
dm.freeze_panes = "D2"
dm.conditional_formatting.add(f"A2:O{DEMO_ROWS+1}", FormulaRule(formula=['$M2=TRUE'], fill=PatternFill("solid", fgColor="C8E6C9")))
dm.conditional_formatting.add(f"A2:O{DEMO_ROWS+1}", FormulaRule(formula=['AND($L2=TRUE,$M2=FALSE)'], fill=PatternFill("solid", fgColor="E3F2FD")))

# ---------- Ore ----------
ho = wb.create_sheet("Ore")
hdr(ho, 1, ["Data", "Setter", "Ore lavorate", "Note"])
for r in range(2, 202):
    ho.cell(row=r, column=1).number_format = "dd.mm"
for col, w in zip("ABCD", [11, 16, 12, 40]):
    ho.column_dimensions[col].width = w
ho.freeze_panes = "A2"

# ---------- Pagamenti ----------
pg = wb.create_sheet("Pagamenti")
pg["A1"] = "PAGAMENTI — calcolati in automatico"
pg["A1"].font = H
pg["A3"] = "Settimana dal (lunedì):"
pg["A3"].font = B
pg["B3"] = dt.date(2026, 11, 16)
pg["B3"].number_format = "dd.mm.yyyy"
pg["B3"].fill = INPUT
pg["C3"] = "← cambia la data per vedere un'altra settimana"
cols = ["Setter", "Ore", "Chiamate (tutto)", "Demo fissate", "Confermate", "Svolte", "Vendite",
        "Da pagare questa settimana €", "Totale progetto €"]
hdr(pg, 5, cols)
S = "Impostazioni"
for i in range(N_SETTERS):
    r = 6 + i
    name = f"{S}!$A${10+i}"
    wk = "$B$3"
    rng = f'">="&{wk},'
    pg.cell(row=r, column=1, value=f"={name}").font = B
    pg.cell(row=r, column=2, value=f'=SUMIFS(Ore!$C:$C,Ore!$B:$B,A{r},Ore!$A:$A,">="&{wk},Ore!$A:$A,"<"&({wk}+7))')
    pg.cell(row=r, column=3, value=f"=COUNTA('Setter {i+1}'!H2:H{LEAD_ROWS+1})")
    pg.cell(row=r, column=4, value=f'=COUNTIFS(Demo!$B:$B,A{r},Demo!$A:$A,">="&{wk},Demo!$A:$A,"<"&({wk}+7))')
    pg.cell(row=r, column=5, value=f'=COUNTIFS(Demo!$B:$B,A{r},Demo!$K:$K,TRUE,Demo!$A:$A,">="&{wk},Demo!$A:$A,"<"&({wk}+7))')
    pg.cell(row=r, column=6, value=f'=COUNTIFS(Demo!$B:$B,A{r},Demo!$L:$L,TRUE,Demo!$H:$H,">="&{wk},Demo!$H:$H,"<"&({wk}+7))')
    pg.cell(row=r, column=7, value=f'=COUNTIFS(Demo!$B:$B,A{r},Demo!$M:$M,TRUE,Demo!$H:$H,">="&{wk},Demo!$H:$H,"<"&({wk}+7))')
    pg.cell(row=r, column=8, value=f"=B{r}*{S}!$B$3+E{r}*{S}!$B$4+F{r}*{S}!$B$5+G{r}*{S}!$B$6")
    pg.cell(row=r, column=9, value=(
        f"=SUMIFS(Ore!$C:$C,Ore!$B:$B,A{r})*{S}!$B$3"
        f"+COUNTIFS(Demo!$B:$B,A{r},Demo!$K:$K,TRUE)*{S}!$B$4"
        f"+COUNTIFS(Demo!$B:$B,A{r},Demo!$L:$L,TRUE)*{S}!$B$5"
        f"+COUNTIFS(Demo!$B:$B,A{r},Demo!$M:$M,TRUE)*{S}!$B$6"))
    for c in range(1, 10):
        cell = pg.cell(row=r, column=c)
        cell.border = BOX
        if c > 1:
            cell.alignment = CENTER
    pg.cell(row=r, column=8).number_format = '#,##0 "€"'
    pg.cell(row=r, column=9).number_format = '#,##0 "€"'
    pg.cell(row=r, column=8).fill = SOFT
tr = 6 + N_SETTERS
pg.cell(row=tr, column=1, value="TOTALE").font = B
for c in range(2, 10):
    L = pg.cell(row=5, column=c).column_letter
    cell = pg.cell(row=tr, column=c, value=f"=SUM({L}6:{L}{tr-1})")
    cell.font, cell.border, cell.alignment = B, BOX, CENTER
pg.cell(row=tr, column=8).number_format = '#,##0 "€"'
pg.cell(row=tr, column=9).number_format = '#,##0 "€"'

pg.cell(row=tr + 2, column=1, value="Acconti incassati (tutto il progetto)").font = B
pg.cell(row=tr + 2, column=2, value="=SUM(Demo!N2:N1000)").number_format = '#,##0 "€"'
pg.cell(row=tr + 3, column=1, value="Costo setter (tutto il progetto)").font = B
pg.cell(row=tr + 3, column=2, value=f"=I{tr}").number_format = '#,##0 "€"'
pg.cell(row=tr + 4, column=1, value="Saldo: acconti − costo setter").font = B
pg.cell(row=tr + 4, column=2, value=f"=B{tr+2}-B{tr+3}").number_format = '#,##0 "€"'
pg.conditional_formatting.add(f"B{tr+4}", FormulaRule(formula=[f"B{tr+4}<0"], fill=PatternFill("solid", fgColor="FFCDD2")))
pg.conditional_formatting.add(f"B{tr+4}", FormulaRule(formula=[f"B{tr+4}>=0"], fill=PatternFill("solid", fgColor="C8E6C9")))
pg.cell(row=tr + 5, column=1, value="Regola: se il saldo è rosso, non aggiungere nuovi setter.").font = Font(italic=True, color="B22222")
for col, w in zip("ABCDEFGHI", [34, 10, 14, 12, 12, 10, 10, 18, 16]):
    pg.column_dimensions[col].width = w

wb.move_sheet("Pagamenti", offset=-(len(wb.sheetnames) - 2))
wb.save("/home/user/Claude/vendita/Setter_Tabella.xlsx")
print("ok", wb.sheetnames, len(days))
