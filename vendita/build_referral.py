"""Genera il file «Рефералы и тёплые контакты» (xlsx) da caricare su Google Sheets."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
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

wb = Workbook()

# ---------------- 1. План ----------------
pl = wb.active
pl.title = "План 15 мин"
pl["A1"] = "РЕФЕРАЛЫ И ТЁПЛЫЕ КОНТАКТЫ — 10–15 минут в день"
pl["A1"].font = H
pl["A2"] = ("Цель: 2–3 клиента из знакомых и рекомендаций к 18.12. Отмечай галочки. "
            "Все тексты — на листе «Шаблоны», учёт — на листе «Контакты».")
pl["A2"].alignment = WRAP
pl.merge_cells("A2:G2")
pl.row_dimensions[2].height = 32

row = 4
def section(title):
    global row
    pl.cell(row=row, column=1, value=title).font = Font(bold=True, size=12, color="12355B")
    row += 1

def header(cols):
    global row
    for i, h in enumerate(cols, 1):
        c = pl.cell(row=row, column=i, value=h)
        c.fill, c.font, c.alignment, c.border = DARK, WB, CENTER, BOX
    row += 1

checkbox_ranges = []

section("СТАРТ — один раз (≈ 45 минут, можно разбить на 3 дня)")
header(["✓", "Задача", "Зачем", "Время"])
start_tasks = [
    ("Выгрузить контакты телефона и выписать в лист «Контакты» тех, кому можно написать", "Это твоя база тёплых контактов", "20 мин"),
    ("Пометить каждого: A — сам владелец бизнеса или знает многих бизнесменов; B — обычный знакомый; C — давно не общались", "Начинаешь с A — там выше шанс", "10 мин"),
    ("Решить вознаграждение: 50 € знакомым за каждого, кто заплатил; клиентам — месяц поддержки или 50 €", "Конкретная награда = в разы больше рекомендаций", "2 мин"),
    ("Попросить рекомендации и отзыв в Google у медцентра, эстетического центра и G.R (шаблон 6)", "Самые тёплые люди: уже видели твою работу", "10 мин"),
    ("Подготовить 3 скриншота лучших сайтов (медцентр и др.) — прикладывать к сообщениям", "Картинка продаёт лучше слов", "5 мин"),
]
first = row
for t, why, tm in start_tasks:
    pl.cell(row=row, column=1, value=False)
    pl.cell(row=row, column=2, value=t)
    pl.cell(row=row, column=3, value=why)
    pl.cell(row=row, column=4, value=tm)
    for c in range(1, 5):
        pl.cell(row=row, column=c).border = BOX
        pl.cell(row=row, column=c).alignment = WRAP
    row += 1
checkbox_ranges.append((first, row - 1, 1, 1))
row += 1

section("КАЖДЫЙ ДЕНЬ — 10–15 минут (после блока звонков)")
header(["Задача", "Сколько", "Пн", "Вт", "Ср", "Чт", "Пт"])
daily = [
    ("Отправить сообщения из списка «Контакты» (сначала группа A) — по одному, с именем, шаблон 1 или 2", "15–20 шт."),
    ("Ответить тем, кто написал; если дали имя — сразу позвонить или записать в «Контакты»", "5 мин"),
    ("Напоминание тем, кто молчит 3–4 дня (шаблон 4) — один раз, не больше", "5–10 шт."),
    ("Обновить статусы в «Контактах»", "2 мин"),
]
first = row
for t, how in daily:
    pl.cell(row=row, column=1, value=t)
    pl.cell(row=row, column=2, value=how)
    for c in range(3, 8):
        pl.cell(row=row, column=c, value=False)
    for c in range(1, 8):
        pl.cell(row=row, column=c).border = BOX
        pl.cell(row=row, column=c).alignment = WRAP if c < 3 else CENTER
    row += 1
checkbox_ranges.append((first, row - 1, 3, 7))
row += 1

section("КАЖДУЮ НЕДЕЛЮ — 20–30 минут в сумме")
header(["Задача", "Шаблон", "Нед 1\n16.11", "Нед 2\n23.11", "Нед 3\n30.11", "Нед 4\n07.12", "Нед 5\n14.12"])
weekly = [
    ("Пост/сторис в WhatsApp-статусе и Instagram: готовый сайт клиента «до/после» (с разрешения)", "7"),
    ("Один полезный пост в 1–2 русско-/украиноязычных группах Италии (Telegram, Facebook)", "8"),
    ("Написать или зайти к 2 партнёрам: коммерчиалист, типография, фотограф, графический дизайнер", "9"),
    ("Отдать награду за каждого оплатившего клиента — в тот же день, и поблагодарить (шаблон 5)", "5"),
    ("Попросить рекомендацию у каждого нового клиента в момент сдачи сайта (шаблон 6)", "6"),
]
first = row
for t, tpl in weekly:
    pl.cell(row=row, column=1, value=t)
    pl.cell(row=row, column=2, value=tpl)
    for c in range(3, 8):
        pl.cell(row=row, column=c, value=False)
    for c in range(1, 8):
        pl.cell(row=row, column=c).border = BOX
        pl.cell(row=row, column=c).alignment = WRAP if c < 2 else CENTER
    row += 1
checkbox_ranges.append((first, row - 1, 3, 7))
row += 1

section("РЕЗУЛЬТАТ (считается сам из листа «Контакты»)")
header(["Показатель", "Факт", "Ориентир к 18.12"])
C = "'Контакты'"
kpis = [
    ("Сообщений отправлено", f'=COUNTA({C}!F2:F2000)', 300),
    ("Ответили", f'=COUNTIF({C}!G2:G2000,"Да*")', "60 (≈ 20%)"),
    ("Дали имя / контакт", f'=COUNTA({C}!H2:H2000)', "15–25"),
    ("Демо с рекомендованными", f'=COUNTIF({C}!I2:I2000,"Демо*")+COUNTIF({C}!I2:I2000,"Клиент*")', "8–10"),
    ("Клиентов из рекомендаций", f'=COUNTIF({C}!I2:I2000,"Клиент*")', "2–3"),
    ("Наград выплачено, €", f'=SUM({C}!K2:K2000)', ""),
]
for label, f, target in kpis:
    pl.cell(row=row, column=1, value=label).font = B
    pl.cell(row=row, column=2, value=f).alignment = CENTER
    pl.cell(row=row, column=3, value=target).alignment = CENTER
    for c in range(1, 4):
        pl.cell(row=row, column=c).border = BOX
    pl.cell(row=row, column=2).fill = SOFT
    row += 1
row += 1
pl.cell(row=row, column=1, value=(
    "Почему это сработает: 300 сообщений → ~20% ответят → ~15–25 имён → по ним ты звонишь уже «от знакомого». "
    "Такие звонки закрываются в разы чаще холодных: 2–3 клиента = 2 000–3 000 € за ~15 минут в день."))
pl.cell(row=row, column=1).alignment = WRAP
pl.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
pl.row_dimensions[row].height = 45

pl.column_dimensions["A"].width = 62
pl.column_dimensions["B"].width = 34
pl.column_dimensions["C"].width = 30
for col in "DEFG":
    pl.column_dimensions[col].width = 10

# ---------------- 2. Шаблоны ----------------
tp = wb.create_sheet("Шаблоны")
tp["A1"] = "ШАБЛОНЫ СООБЩЕНИЙ — копируй, подставь имя, отправляй по одному (не рассылкой)"
tp["A1"].font = H
heads = ["№", "Кому / когда", "Текст 🇮🇹 / основной", "Перевод 🇷🇺 / версия для русскоязычных"]
for i, h in enumerate(heads, 1):
    c = tp.cell(row=3, column=i, value=h)
    c.fill, c.font, c.alignment, c.border = DARK, WB, CENTER, BOX

templates = [
    ("1", "Итальянские знакомые и друзья",
     "Ciao [Nome], come stai? 🙂 Ti scrivo per un piccolo favore: in questo periodo realizzo siti web per piccole attività "
     "(artigiani, imprese edili, officine, negozi, ristoranti) — belli da telefono e collegati a Google e WhatsApp. "
     "Conosci qualcuno che ha un'attività senza sito o con un sito vecchio? Chi ti viene in mente per primo? "
     "Se diventa cliente, ti ringrazio con 50 euro. Grazie!",
     "Привет, [Имя], как дела? 🙂 Пишу с небольшой просьбой: сейчас делаю сайты для малого бизнеса (мастера, строительные "
     "фирмы, автосервисы, магазины, рестораны) — удобные с телефона и связанные с Google и WhatsApp. Знаешь кого-то с бизнесом "
     "без сайта или со старым сайтом? Кто первым приходит в голову? Если станет клиентом — отблагодарю 50 евро. Спасибо!"),
    ("2", "Русско-/украиноязычные знакомые",
     "(для пересылки итальянцам — шаблон 3)",
     "Привет, [Имя]! Как ты? 🙂 У меня небольшая просьба. Я сейчас делаю сайты для бизнеса в Италии: быстро, красиво с "
     "телефона, с кнопками звонка и WhatsApp и с настройкой карточки в Google. Если знаешь кого-то с бизнесом без нормального "
     "сайта — мастер, строитель, автосервис, салон, ресторан, кто угодно, — подскажешь? Кто первый приходит на ум? "
     "За каждого, кто станет клиентом, отблагодарю 50 €. Могу скинуть короткий текст, чтобы просто переслать."),
    ("3", "Текст, который знакомый пересылает дальше",
     "Ciao! Ti giro il contatto di Egor: fa siti web per attività locali, veloci e fatti bene da telefono, collegati a "
     "Google e WhatsApp. Ti prepara anche una bozza gratuita per vedere come verrebbe, senza impegno. "
     "Il suo numero: [numero]",
     "Привет! Пересылаю контакт Егора: делает сайты для местного бизнеса, быстро и качественно, удобные с телефона, "
     "связанные с Google и WhatsApp. Бесплатно сделает черновик, чтобы посмотреть, как будет выглядеть, без обязательств. "
     "Его номер: [номер]"),
    ("4", "Напоминание через 3–4 дня (один раз)",
     "Ciao [Nome]! Ti è venuto in mente qualcuno? Anche solo un nome va benissimo, lo contatto io con discrezione 🙂",
     "Привет, [Имя]! Никто не пришёл на ум? Даже одно имя — уже отлично, я сам аккуратно свяжусь 🙂"),
    ("5", "Благодарность и награда",
     "Ciao [Nome]! [Amico] ha appena iniziato il sito con me — come promesso, ecco i tuoi 50 euro. Grazie davvero, "
     "se ti viene in mente qualcun altro, sono qui!",
     "Привет, [Имя]! [Друг] только что начал сайт со мной — как и обещал, вот твои 50 евро. Огромное спасибо! "
     "Если вспомнишь ещё кого-то — я на связи!"),
    ("6", "Клиенту при сдаче сайта (самый «счастливый» момент)",
     "Sono contento che il sito le piaccia! Due piccole cose: mi lascerebbe una recensione su Google? Ecco il link: [link]. "
     "E se conosce un altro titolare che ha bisogno di un sito, per ogni cliente che arriva da lei le regalo un mese "
     "di assistenza oppure 50 euro.",
     "Рад, что сайт понравился! Две небольшие просьбы: оставите отзыв в Google? Вот ссылка: [ссылка]. И если знаете "
     "другого владельца, которому нужен сайт, — за каждого клиента от вас дарю месяц поддержки или 50 евро."),
    ("7", "Статус WhatsApp / сторис раз в неделю",
     "Nuovo sito online! 🚀 [Attività] di [città] adesso si trova su Google e riceve richieste da WhatsApp. "
     "Hai un'attività senza sito? Scrivimi: ti preparo una bozza gratuita.",
     "Новый сайт онлайн! 🚀 [Бизнес] из [город] теперь находится в Google и получает заявки через WhatsApp. "
     "У тебя бизнес без сайта? Напиши — сделаю бесплатный черновик."),
    ("8", "Пост в русско-/украиноязычной группе",
     "(группы: «Русские в Вероне/Италии», «Украинцы в Италии» и т. п.)",
     "Друзья, привет! Я Егор, живу в Вероне и делаю сайты для бизнеса в Италии — для мастеров, строительных фирм, "
     "автосервисов, салонов, ресторанов. Сайт удобный с телефона, с кнопками звонка/WhatsApp, на итальянском (можно и "
     "на русском/английском), плюс настройка карточки в Google, чтобы вас находили клиенты. Делаю бесплатный черновик, "
     "чтобы вы увидели результат до оплаты. Пишите в личку — отвечу на любые вопросы про сайты и Google 🙂"),
    ("9", "Партнёру: коммерчиалист, типография, фотограф",
     "Buongiorno [Nome], sono Egor, realizzo siti web per piccole attività qui a Verona. Molti dei suoi clienti "
     "aprono o gestiscono un'attività e prima o poi hanno bisogno di un sito. Le propongo una collaborazione semplice: "
     "se mi presenta un cliente che poi lavora con me, le riconosco il 15% del progetto. Le andrebbe un caffè di "
     "10 minuti questa settimana?",
     "Добрый день, [Имя], я Егор, делаю сайты для малого бизнеса здесь, в Вероне. Многие ваши клиенты открывают или ведут "
     "бизнес, и рано или поздно им нужен сайт. Предлагаю простое сотрудничество: если познакомите меня с клиентом, который "
     "потом начнёт работать со мной, я плачу вам 15% от проекта. Как насчёт кофе на 10 минут на этой неделе?"),
    ("10", "Звонок по рекомендации (первые слова)",
     "Buongiorno, sono Egor, mi ha dato il suo numero [Nome amico]. Mi ha detto che forse le serve un sito per [attività]: "
     "ha due minuti?",
     "Добрый день, я Егор, ваш номер мне дал [Имя друга]. Он сказал, что вам, возможно, нужен сайт для [бизнес]: "
     "есть пара минут?"),
]
for i, t in enumerate(templates):
    r = 4 + i
    for j, v in enumerate(t, 1):
        c = tp.cell(row=r, column=j, value=v)
        c.alignment, c.border = WRAP, BOX
    tp.cell(row=r, column=1).alignment = CENTER
    tp.cell(row=r, column=2).font = B
    tp.row_dimensions[r].height = 150
tp.column_dimensions["A"].width = 5
tp.column_dimensions["B"].width = 26
tp.column_dimensions["C"].width = 70
tp.column_dimensions["D"].width = 70

r = 4 + len(templates) + 1
rules = [
    "ПРАВИЛА, КОТОРЫЕ ПОДНИМАЮТ ОТДАЧУ",
    "• Только лично, по одному, с именем. Массовая рассылка выглядит как спам и получает в разы меньше ответов.",
    "• Просить одно конкретное имя («кто первым приходит в голову?»), а не «если вдруг кто-то…».",
    "• Сразу давать готовый текст для пересылки (шаблон 3) — человеку не нужно ничего придумывать.",
    "• Награду платить в тот же день, когда клиент заплатил, и рассказывать об этом — так слух расходится быстрее.",
    "• Одно напоминание — максимум. Не дави на знакомых: отношения важнее одного лида.",
    "• Звонок по рекомендации — всегда в первые 24 часа и со ссылкой на имя друга (шаблон 10).",
    "• Награды частным лицам: держи суммы скромными и уточни у коммерчиалиста, как их правильно оформлять.",
]
for k, line in enumerate(rules):
    c = tp.cell(row=r + k, column=2, value=line)
    c.font = Font(bold=(k == 0), size=12 if k == 0 else 11, color="12355B" if k == 0 else "000000")
    tp.merge_cells(start_row=r + k, start_column=2, end_row=r + k, end_column=4)
tp.freeze_panes = "A4"

# ---------------- 3. Контакты ----------------
ct = wb.create_sheet("Контакты")
cols = ["Имя", "Телефон", "Язык", "Группа (A/B/C)", "Кто это", "Дата сообщения", "Ответил?",
        "Кого порекомендовал (имя, телефон, бизнес)", "Статус рекомендации", "Дата напоминания", "Награда €", "Заметки"]
for i, h in enumerate(cols, 1):
    c = ct.cell(row=1, column=i, value=h)
    c.fill, c.font, c.alignment, c.border = DARK, WB, CENTER, BOX
ct.row_dimensions[1].height = 36
widths = [20, 16, 9, 10, 22, 13, 12, 40, 18, 14, 10, 30]
for i, w in enumerate(widths, 1):
    ct.column_dimensions[ct.cell(row=1, column=i).column_letter].width = w
for r in range(2, 302):
    ct.cell(row=r, column=6).number_format = "dd.mm"
    ct.cell(row=r, column=10).number_format = "dd.mm"
    ct.cell(row=r, column=11).number_format = '#,##0 "€"'
dvs = [
    ("C2:C2000", '"IT,RU,UA,EN"'),
    ("D2:D2000", '"A,B,C"'),
    ("E2:E2000", '"Друг/знакомый,Владелец бизнеса,Клиент,Партнёр,Группа/сообщество"'),
    ("G2:G2000", '"Да — дал имя,Да — никого не знает,Нет ответа"'),
    ("I2:I2000", '"Новый,Позвонил,Демо назначено,Демо проведено,Клиент (оплатил),Отказ"'),
]
for rng, formula in dvs:
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ct.add_data_validation(dv)
    dv.add(rng)
ct.freeze_panes = "B2"
ct.conditional_formatting.add("A2:L2000", FormulaRule(formula=['LEFT($I2,6)="Клиент"'], fill=PatternFill("solid", fgColor="C8E6C9")))
ct.conditional_formatting.add("A2:L2000", FormulaRule(formula=['LEFT($I2,4)="Демо"'], fill=PatternFill("solid", fgColor="FFF4C2")))
ct.conditional_formatting.add("A2:L2000", FormulaRule(formula=['LEFT($G2,13)="Да — дал имя"'], fill=PatternFill("solid", fgColor="E3F2FD")))

# ---------------- 4. Награды ----------------
rw = wb.create_sheet("Награды")
rw["A1"] = "ПРОГРАММА РЕКОМЕНДАЦИЙ — что ты обещаешь"
rw["A1"].font = H
rows = [
    ("Кто рекомендует", "Награда", "Когда платишь"),
    ("Друг / знакомый", "50 € за каждого клиента", "В день, когда клиент внёс предоплату"),
    ("Текущий клиент", "1 месяц поддержки бесплатно или 50 €", "В день оплаты нового клиента"),
    ("Партнёр (коммерчиалист, типография, фотограф)", "15% от суммы проекта", "После полной оплаты проекта"),
    ("Новый клиент из рекомендации", "Бесплатная настройка карточки Google (вместо скидки)", "Входит в проект"),
]
for i, rr in enumerate(rows):
    for j, v in enumerate(rr, 1):
        c = rw.cell(row=3 + i, column=j, value=v)
        c.border, c.alignment = BOX, WRAP
        if i == 0:
            c.fill, c.font = DARK, WB
for col, w in zip("ABC", [42, 46, 40]):
    rw.column_dimensions[col].width = w
rw["A10"] = ("Почему не скидка новому клиенту, а бонус: скидка обесценивает твою работу, "
             "а бесплатная настройка Google стоит тебе 1 час и выглядит щедро.")
rw["A10"].alignment = WRAP
rw.merge_cells("A10:C10")
rw.row_dimensions[10].height = 36

wb.save("/home/user/Claude/vendita/Referral_Teplye_Kontakty.xlsx")
print("ok", checkbox_ranges)
