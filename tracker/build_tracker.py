"""Genera il tracker delle chiamate (xlsx) da caricare su Google Sheets."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule

wb = Workbook()
S = "'Настройки'"
D = "'Дневник'"
N_DAYS = 24
LAST = 1 + N_DAYS  # ultima riga dati in Дневник

DARK = PatternFill("solid", fgColor="12355B")
INPUT = PatternFill("solid", fgColor="FFF4C2")
CALC = PatternFill("solid", fgColor="E8EEF5")
WHITE_B = Font(bold=True, color="FFFFFF")
B = Font(bold=True)
thin = Side(style="thin", color="B9C4D0")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="center", wrap_text=True)

def hdr(ws, row, cols):
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = DARK, WHITE_B, CENTER, BOX

# ---------------- Настройки ----------------
st = wb.active
st.title = "Настройки"
st["A1"] = "НАСТРОЙКИ ПЛАНА — цель и конверсии (пессимистичный сценарий)"
st["A1"].font = Font(bold=True, size=14, color="12355B")
st["A2"] = "Жёлтые ячейки можно менять. Когда узнаешь реальную конверсию (после 5–10 проведённых демо), впиши её — план пересчитается сам."
st["A2"].alignment = WRAP
st.merge_cells("A2:D2")
st.row_dimensions[2].height = 32

inputs = [
    (4, "Цель, €", 10000, '#,##0 "€"'),
    (5, "Средний чек, €", 900, '#,##0 "€"'),
    (6, "% демо от звонков", 0.03, "0.0%"),
    (7, "% явки на демо", 0.65, "0%"),
    (8, "% оплат от проведённых демо", 0.25, "0%"),
    (9, "% живых номеров (мобильные)", 0.91, "0%"),
    (10, "Первый день обзвона", dt.date(2026, 11, 16), "dd.mm.yyyy"),
    (11, "Праздник (выходной)", dt.date(2026, 12, 8), "dd.mm.yyyy"),
]
st["A3"], st["B3"], st["C3"] = "Параметр", "Значение", "Комментарий"
hdr(st, 3, [1, 2, 3])
notes = {
    5: "Пессимистично. При 1 000 € звонков нужно на ~10% меньше",
    6: "Твои мобильные сейчас ≈ 4%; закладываем 3%",
    7: "1 из 2 первых демо сорвалось и перенесено",
    8: "Пока не измерено — самая неопределённая цифра",
    9: "По твоей таблице: 49 из 54 мобильных живые",
    10: "Понедельник. Даты в «Дневнике» считаются от этой даты",
    11: "8 декабря — Immacolata, не звоним",
}
for r, label, val, fmt in inputs:
    st.cell(row=r, column=1, value=label).border = BOX
    c = st.cell(row=r, column=2, value=val)
    c.fill, c.number_format, c.border, c.font = INPUT, fmt, BOX, B
    n = st.cell(row=r, column=3, value=notes.get(r, ""))
    n.border, n.alignment = BOX, WRAP

st["A13"], st["B13"], st["C13"] = "Расчёт плана", "Значение", "Как считается"
hdr(st, 13, [1, 2, 3])
calc = [
    (14, "Нужно продаж", "=ROUNDUP(B4/B5,0)", "0", "Цель ÷ средний чек"),
    (15, "Продаж на 100 звонков", "=100*B6*B7*B8", "0.00", "демо × явка × оплата"),
    (16, "Нужно звонков ВСЕГО", "=ROUNDUP(B14/(B6*B7*B8),0)", "#,##0", "Главная цифра: сколько наборов номера за период"),
    (17, "Дней обзвона (пн–пт)", f"=COUNT({D}!A2:A{LAST})", "0", "Рабочие дни в «Дневнике», без праздника"),
    (18, "ПЛАН ЗВОНКОВ В ДЕНЬ", "=ROUNDUP(B16/B17,0)", "0", "При ~100 звонках ≈ 3–4 часа"),
    (19, "Демо назначить (всего)", "=ROUNDUP(B16*B6,0)", "0", ""),
    (20, "Демо провести (всего)", "=ROUNDUP(B19*B7,0)", "0", ""),
    (21, "Мобильных номеров в базе (≈)", "=ROUNDUP(B16*0.95,-2)", "#,##0", "~5% звонков — повторные наборы"),
    (22, "Последний день обзвона", f"={D}!A{LAST}", "dd.mm.yyyy", "Потом 2–3 дня на оплаты и сдачу сайтов"),
]
for r, label, f, fmt, note in calc:
    st.cell(row=r, column=1, value=label).border = BOX
    c = st.cell(row=r, column=2, value=f)
    c.fill, c.number_format, c.border = CALC, fmt, BOX
    n = st.cell(row=r, column=3, value=note)
    n.border, n.alignment = BOX, WRAP
for r in (16, 18):
    st.cell(row=r, column=1).font = B
    st.cell(row=r, column=2).font = Font(bold=True, size=13, color="B22222")

st["A25"] = "Сравнение сценариев"
st["A25"].font = Font(bold=True, size=12, color="12355B")
heads = ["Сценарий", "% демо", "Явка", "Оплата", "Чек, €", "Звонков всего", "Звонков в день"]
for i, h in enumerate(heads, 1):
    st.cell(row=26, column=i, value=h)
hdr(st, 26, range(1, 8))
scen = [("Пессимистичный (план)", "=B6", "=B7", "=B8", "=B5"),
        ("Базовый", 0.04, 0.70, 0.30, 1000),
        ("Хороший", 0.05, 0.75, 0.35, 1100)]
for i, (name, a, b, c_, d) in enumerate(scen):
    r = 27 + i
    vals = [name, a, b, c_, d,
            f"=ROUNDUP(ROUNDUP($B$4/E{r},0)/(B{r}*C{r}*D{r}),0)",
            f"=ROUNDUP(F{r}/$B$17,0)"]
    fmts = [None, "0%", "0%", "0%", '#,##0 "€"', "#,##0", "0"]
    for j, (v, fm) in enumerate(zip(vals, fmts), 1):
        cell = st.cell(row=r, column=j, value=v)
        cell.border = BOX
        if fm:
            cell.number_format = fm
st.column_dimensions["A"].width = 32
st.column_dimensions["B"].width = 16
st.column_dimensions["C"].width = 52
for col in "DEFG":
    st.column_dimensions[col].width = 14

# ---------------- Дневник ----------------
lg = wb.create_sheet("Дневник")
cols = ["Дата", "День", "План звонков", "Звонков", "Мёртвых номеров", "Разговоров с владельцем",
        "Демо назначено", "Демо проведено", "Продаж (оплата получена)", "Получено €",
        "% плана", "Результат дня", "Звонков накоплено", "План накопл.", "€ накоплено", "План € накопл.", "Заметки"]
for i, h in enumerate(cols, 1):
    lg.cell(row=1, column=i, value=h)
hdr(lg, 1, range(1, len(cols) + 1))
lg.row_dimensions[1].height = 42
for r in range(2, LAST + 1):
    lg.cell(row=r, column=1, value=f"={S}!B10" if r == 2 else f"=WORKDAY(A{r-1},1,{S}!$B$11)").number_format = "dd.mm"
    lg.cell(row=r, column=2, value=f'=CHOOSE(WEEKDAY(A{r},2),"Пн","Вт","Ср","Чт","Пт","Сб","Вс")')
    lg.cell(row=r, column=3, value=f"={S}!$B$18")
    for c in range(4, 11):
        lg.cell(row=r, column=c).fill = INPUT
    lg.cell(row=r, column=10).number_format = '#,##0 "€"'
    lg.cell(row=r, column=11, value=f'=IF(D{r}="","",D{r}/C{r})').number_format = "0%"
    lg.cell(row=r, column=12, value=f'=IF(D{r}="","",IF(D{r}>=C{r},"✅ Успех","❌ −"&(C{r}-D{r})))')
    lg.cell(row=r, column=13, value=f'=IF(D{r}="","",SUM($D$2:D{r}))')
    lg.cell(row=r, column=14, value=f"=SUM($C$2:C{r})")
    lg.cell(row=r, column=15, value=f'=IF(D{r}="","",SUM($J$2:J{r}))').number_format = '#,##0 "€"'
    lg.cell(row=r, column=16, value=f"=ROUND(MIN({S}!$B$4,{S}!$B$4*N{r}/{S}!$B$16),0)").number_format = '#,##0 "€"'
    for c in range(1, len(cols) + 1):
        lg.cell(row=r, column=c).border = BOX
        if c < 17:
            lg.cell(row=r, column=c).alignment = Alignment(horizontal="center")
tot = LAST + 1
lg.cell(row=tot, column=1, value="ИТОГО").font = B
for c, L in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J")]:
    cell = lg.cell(row=tot, column=c, value=f"=SUM({L}2:{L}{LAST})")
    cell.font, cell.fill, cell.border = B, CALC, BOX
lg.cell(row=tot, column=10).number_format = '#,##0 "€"'
widths = [9, 6, 10, 10, 10, 12, 10, 10, 12, 11, 9, 13, 12, 11, 12, 12, 40]
for i, w in enumerate(widths, 1):
    lg.column_dimensions[lg.cell(row=1, column=i).column_letter].width = w
lg.freeze_panes = "D2"
rng = f"L2:L{LAST}"
lg.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(L2,1)="✅"'], fill=PatternFill("solid", fgColor="C8E6C9")))
lg.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(L2,1)="❌"'], fill=PatternFill("solid", fgColor="FFCDD2")))
lg.conditional_formatting.add(f"K2:K{LAST}", ColorScaleRule(start_type="num", start_value=0.5, start_color="F8696B",
                                                             mid_type="num", mid_value=0.85, mid_color="FFEB84",
                                                             end_type="num", end_value=1, end_color="63BE7B"))

# ---------------- Дашборд ----------------
db = wb.create_sheet("Дашборд", 0)
db["A1"] = "📞 ТРЕКЕР ОБЗВОНА — ЦЕЛЬ 10 000 €"
db["A1"].font = Font(bold=True, size=18, color="12355B")
db["A2"] = f'="Период обзвона: "&TEXT({S}!B10,"dd.mm")&" – "&TEXT({S}!B22,"dd.mm.yyyy")&"   ·   план: "&{S}!B18&" звонков в день"'
db["A2"].font = Font(size=11, color="555555")

db["A4"] = "ГЛАВНОЕ"
db["A4"].font = Font(bold=True, size=12, color="12355B")
for i, h in enumerate(["Показатель", "Факт", "План", "% плана", "Прогресс"], 1):
    db.cell(row=5, column=i, value=h)
hdr(db, 5, range(1, 6))
kpi = [
    (6, "📞 Звонков всего", f"={D}!D{tot}", f"={S}!B16", "#,##0"),
    (7, "💶 Получено €", f"={D}!J{tot}", f"={S}!B4", '#,##0 "€"'),
    (8, "🤝 Продаж", f"={D}!I{tot}", f"={S}!B14", "0"),
    (9, "📅 Демо назначено", f"={D}!G{tot}", f"={S}!B19", "0"),
    (10, "🎯 Демо проведено", f"={D}!H{tot}", f"={S}!B20", "0"),
    (11, "✅ Успешных дней (план звонков)", f'=COUNTIF({D}!L2:L{LAST},"✅*")', f"={S}!B17", "0"),
]
for r, label, fact, plan, fmt in kpi:
    db.cell(row=r, column=1, value=label).font = B
    for c, v in [(2, fact), (3, plan)]:
        cell = db.cell(row=r, column=c, value=v)
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="center")
    db.cell(row=r, column=4, value=f"=IFERROR(B{r}/C{r},0)").number_format = "0%"
    db.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    for c in range(1, 6):
        db.cell(row=r, column=c).border = BOX
db.conditional_formatting.add("D6:D11", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                                                       mid_type="num", mid_value=0.5, mid_color="FFEB84",
                                                       end_type="num", end_value=1, end_color="63BE7B"))

db["A13"] = "СЕГОДНЯ"
db["A13"].font = Font(bold=True, size=12, color="12355B")
today = [
    (14, "Дней обзвона прошло", f"=COUNT({D}!D2:D{LAST})", "0"),
    (15, "Дней осталось", f"={S}!B17-B14", "0"),
    (16, "Опережение (+) / отставание (−) по звонкам",
     f'=SUM({D}!D2:D{LAST})-SUMPRODUCT(({D}!D2:D{LAST}<>"")*{D}!C2:C{LAST})', "+#,##0;−#,##0;0"),
    (17, "Нужно звонков в день, чтобы выйти на план",
     f"=IF(B15<=0,0,ROUNDUP(MAX(0,{S}!B16-B6)/B15,0))", "0"),
    (18, "До цели осталось €", f"=MAX(0,{S}!B4-B7)", '#,##0 "€"'),
    (19, "Мёртвых номеров", f"=IFERROR({D}!E{tot}/{D}!D{tot},0)", "0%"),
    (20, "Разговоров с владельцем от звонков", f"=IFERROR({D}!F{tot}/{D}!D{tot},0)", "0%"),
]
for r, label, f, fmt in today:
    db.cell(row=r, column=1, value=label).font = B
    cell = db.cell(row=r, column=2, value=f)
    cell.number_format, cell.alignment, cell.fill = fmt, Alignment(horizontal="center"), CALC
    for c in (1, 2):
        db.cell(row=r, column=c).border = BOX
db.conditional_formatting.add("B16", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="FFCDD2")))
db.conditional_formatting.add("B16", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PatternFill("solid", fgColor="C8E6C9")))

db["A22"] = "КОНВЕРСИЯ: ФАКТ vs ДОПУЩЕНИЕ"
db["A22"].font = Font(bold=True, size=12, color="12355B")
for i, h in enumerate(["Этап", "Факт", "Допущение", "Статус"], 1):
    db.cell(row=23, column=i, value=h)
hdr(db, 23, range(1, 5))
conv = [
    (24, "Демо / звонки", f'=IF({D}!D{tot}=0,"—",{D}!G{tot}/{D}!D{tot})', f"={S}!B6", "0.0%", f"{D}!D{tot}>=100"),
    (25, "Явка на демо", f'=IF({D}!G{tot}=0,"—",{D}!H{tot}/{D}!G{tot})', f"={S}!B7", "0%", f"{D}!G{tot}>=5"),
    (26, "Оплата / проведённые демо", f'=IF({D}!H{tot}=0,"—",{D}!I{tot}/{D}!H{tot})', f"={S}!B8", "0%", f"{D}!H{tot}>=5"),
    (27, "Средний чек", f'=IF({D}!I{tot}=0,"—",{D}!J{tot}/{D}!I{tot})', f"={S}!B5", '#,##0 "€"', f"{D}!I{tot}>=3"),
]
for r, label, fact, plan, fmt, enough in conv:
    db.cell(row=r, column=1, value=label).font = B
    for c, v in [(2, fact), (3, plan)]:
        cell = db.cell(row=r, column=c, value=v)
        cell.number_format, cell.alignment = fmt, Alignment(horizontal="center")
    db.cell(row=r, column=4, value=f'=IF(B{r}="—","нет данных",IF(NOT({enough}),"мало данных",IF(B{r}>=C{r},"✅ не хуже плана","⚠️ ниже плана")))')
    for c in range(1, 5):
        db.cell(row=r, column=c).border = BOX
db.cell(row=28, column=1, value="🔮 Прогноз звонков всего при фактической конверсии").font = B
db.cell(row=28, column=2, value=f'=IF({D}!H{tot}<5,"после 5 проведённых демо",IFERROR(ROUNDUP({S}!B4/B27/(B24*B25*B26),0),"нужно больше данных"))')
db.cell(row=28, column=2).number_format = "#,##0"
db.merge_cells("B28:D28")
for c in range(1, 5):
    db.cell(row=28, column=c).border = BOX
db.conditional_formatting.add("D24:D27", FormulaRule(formula=['LEFT(D24,1)="✅"'], fill=PatternFill("solid", fgColor="C8E6C9")))
db.conditional_formatting.add("D24:D27", FormulaRule(formula=['LEFT(D24,1)="⚠"'], fill=PatternFill("solid", fgColor="FFE0B2")))

db["A30"] = "ПО НЕДЕЛЯМ"
db["A30"].font = Font(bold=True, size=12, color="12355B")
wk = ["Неделя", "План звонков", "Звонков", "% плана", "Демо проведено", "Продаж", "Получено €"]
for i, h in enumerate(wk, 1):
    db.cell(row=31, column=i, value=h)
hdr(db, 31, range(1, 8))
for n in range(1, 6):
    r = 31 + n
    start = f"({D}!$A$2+{7*(n-1)})"
    crit = f'{D}!$A$2:$A${LAST},">="&{start},{D}!$A$2:$A${LAST},"<"&({start}+7)'
    db.cell(row=r, column=1, value=f'="Неделя {n} ("&TEXT({start},"dd.mm")&")"')
    db.cell(row=r, column=2, value=f"=SUMIFS({D}!$C$2:$C${LAST},{crit})")
    db.cell(row=r, column=3, value=f"=SUMIFS({D}!$D$2:$D${LAST},{crit})")
    db.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)").number_format = "0%"
    db.cell(row=r, column=5, value=f"=SUMIFS({D}!$H$2:$H${LAST},{crit})")
    db.cell(row=r, column=6, value=f"=SUMIFS({D}!$I$2:$I${LAST},{crit})")
    db.cell(row=r, column=7, value=f"=SUMIFS({D}!$J$2:$J${LAST},{crit})").number_format = '#,##0 "€"'
    for c in range(1, 8):
        db.cell(row=r, column=c).border = BOX
        if c > 1:
            db.cell(row=r, column=c).alignment = Alignment(horizontal="center")
db.conditional_formatting.add("D32:D36", ColorScaleRule(start_type="num", start_value=0.5, start_color="F8696B",
                                                        mid_type="num", mid_value=0.85, mid_color="FFEB84",
                                                        end_type="num", end_value=1, end_color="63BE7B"))
for col, w in zip("ABCDEFG", [44, 16, 14, 16, 16, 10, 13]):
    db.column_dimensions[col].width = w

# ---------------- Инструкция ----------------
ins = wb.create_sheet("Инструкция")
lines = [
    ("КАК ПОЛЬЗОВАТЬСЯ ТРЕКЕРОМ", True),
    ("1. Каждый вечер (5 минут) открой лист «Дневник» и заполни жёлтые ячейки за сегодня из своей таблицы лидов.", False),
    ("2. «Звонков» — все наборы номера, включая мёртвые и неотвеченные.", False),
    ("3. «Разговоров с владельцем» — реальный разговор с тем, кто принимает решение (да, нет, есть сайт, перезвонить).", False),
    ("4. «Продаж» — только когда ПОЛУЧЕНА предоплата. «Да» без денег не считается.", False),
    ("5. «Получено €» — деньги, пришедшие в этот день (предоплаты и остатки). Демо и оплаты в субботу записывай в пятничную строку.", False),
    ("6. День = ✅ Успех, если звонков не меньше плана. Это единственное, что ты полностью контролируешь.", False),
    ("7. Лист «Дашборд» считает всё сам: прогресс, отставание, сколько звонков в день нужно теперь, факт. конверсию и графики.", False),
    ("8. После 5–10 проведённых демо сравни факт с допущениями. Если факт стабильно другой — впиши его в «Настройки».", False),
    ("", False),
    ("ПРАВИЛА РЕШЕНИЙ", True),
    ("• После 300 мобильных звонков: демо < 3% → меняй нишу или первые 20 секунд скрипта.", False),
    ("• После 10 проведённых демо: оплат < 2 → проблема в демо/цене/предложении, а не в количестве звонков.", False),
    ("• Отставание больше одного дневного плана → добавь блок звонков, а не переноси на «потом».", False),
    ("• Вторая половина декабря медленная (праздники): старайся набрать опережение в ноябре.", False),
]
for i, (t, bold) in enumerate(lines, 1):
    c = ins.cell(row=i, column=1, value=t)
    c.font = Font(bold=True, size=12, color="12355B") if bold else Font(size=11)
    c.alignment = WRAP
ins.column_dimensions["A"].width = 120

wb.save("/home/user/Claude/tracker/Tracker_Obzvon_10000.xlsx")
print("ok")
