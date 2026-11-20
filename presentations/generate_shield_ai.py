import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BG_DARK = RGBColor(11, 15, 20)       # #0B0F14
SURFACE = RGBColor(18, 24, 33)       # #121821
SURFACE_BORDER = RGBColor(30, 41, 59)# #1E293B
PRIMARY_GREEN = RGBColor(0, 168, 89) # #00A859 (BCC Green)
EMERALD = RGBColor(16, 185, 129)     # #10B981
TEXT_WHITE = RGBColor(248, 250, 252) # #F8FAFC
TEXT_MUTED = RGBColor(148, 163, 184) # #94A3B8
TEXT_DIM = RGBColor(100, 116, 139)   # #64748B
AMBER = RGBColor(245, 158, 11)       # #F59E0B
RED = RGBColor(239, 68, 68)          # #EF4444
BLUE = RGBColor(56, 189, 248)        # #38BDF8
PURPLE = RGBColor(168, 85, 247)      # #A855F7

FONT_HEAD = "Segoe UI"
FONT_BODY = "Segoe UI"
FONT_MONO = "Consolas"

def create_base_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def set_slide_bg(slide, prs):
    blank_layout = prs.slide_layouts[6]
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = BG_DARK
    bg_shape.line.fill.background()
    return bg_shape

def add_header(slide, slide_num, title, subtitle=None, tag="BCC SHIELD AI"):
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8), Inches(0.35))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = f"{tag}  ·  СЛАЙД {slide_num:02d}"
    p_tag.font.name = FONT_MONO
    p_tag.font.size = Pt(11)
    p_tag.font.bold = True
    p_tag.font.color.rgb = PRIMARY_GREEN

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title
    p_title.font.name = FONT_HEAD
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = TEXT_WHITE

    if subtitle:
        p_sub = tf_title.add_paragraph()
        p_sub.text = subtitle
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(13)
        p_sub.font.color.rgb = TEXT_MUTED

def add_footer(slide, disclaimer="Интеграции и данные демо смоделированы"):
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = SURFACE_BORDER
    line.line.fill.background()

    footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.95), Inches(8), Inches(0.4))
    tf = footer_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"🛡 Банк ЦентрКредит  |  {disclaimer}"
    p.font.name = FONT_BODY
    p.font.size = Pt(9.5)
    p.font.color.rgb = TEXT_DIM

def add_card(slide, left, top, width, height, title="", items=None, accent_color=None, bg_color=None, border_color=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color or SURFACE
    card.line.color.rgb = border_color or (accent_color if accent_color else SURFACE_BORDER)
    card.line.width = Pt(1.2 if accent_color else 0.8)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)
    tf.margin_bottom = Inches(0.25)
    
    if title:
        p_title = tf.paragraphs[0]
        p_title.text = title
        p_title.font.name = FONT_HEAD
        p_title.font.size = Pt(15)
        p_title.font.bold = True
        p_title.font.color.rgb = accent_color or TEXT_WHITE
        p_title.space_after = Pt(8)
    
    if items:
        start_idx = 1 if title else 0
        for i, itm in enumerate(items):
            if start_idx == 0 and i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = f"• {itm}"
            p.font.name = FONT_BODY
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_WHITE
            p.space_after = Pt(6)
            
    return card

def build_shield_ai():
    prs = create_base_presentation()
    layout = prs.slide_layouts[6]
    
    # ---------------- SLIDE 1: ТИТУЛЬНЫЙ ----------------
    s1 = prs.slides.add_slide(layout)
    set_slide_bg(s1, prs)
    
    title_box = s1.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11.5), Inches(3.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p_tag = tf.paragraphs[0]
    p_tag.text = "NEXT-GEN BEHAVIORAL ANTIFRAUD · HACKATHON 2026"
    p_tag.font.name = FONT_MONO
    p_tag.font.size = Pt(12)
    p_tag.font.bold = True
    p_tag.font.color.rgb = PRIMARY_GREEN
    p_tag.space_after = Pt(14)
    
    p_main = tf.add_paragraph()
    p_main.text = "BCC Shield AI"
    p_main.font.name = FONT_HEAD
    p_main.font.size = Pt(46)
    p_main.font.bold = True
    p_main.font.color.rgb = TEXT_WHITE
    p_main.space_after = Pt(6)
    
    p_sub = tf.add_paragraph()
    p_sub.text = "Поведенческий антифрод нового поколения · Защита клиента от социальной инженерии 2026"
    p_sub.font.name = FONT_HEAD
    p_sub.font.size = Pt(19)
    p_sub.font.color.rgb = EMERALD
    p_sub.space_after = Pt(14)
    
    p_desc = tf.add_paragraph()
    p_desc.text = "Комплексная экосистема противодействия телефонному мошенничеству, дипфейкам, удалённому доступу и дроп-сетям. Демо-стенд · данные смоделированы."
    p_desc.font.name = FONT_BODY
    p_desc.font.size = Pt(13)
    p_desc.font.color.rgb = TEXT_MUTED

    features = [
        ("Архитектура", "Real-Time SOC", PRIMARY_GREEN),
        ("Когнитивный барьер", "Клиентский опыт", BLUE),
        ("Ловушка для фрода", "«Тихий щит»", AMBER),
        ("Кросс-отраслевой", "Telco-Bank Fusion", PURPLE),
    ]
    for i, (head, val, col) in enumerate(features):
        left = Inches(0.9 + i * 2.9)
        card = add_card(s1, left, Inches(4.7), Inches(2.75), Inches(1.8), bg_color=SURFACE)
        tf_c = card.text_frame
        p1 = tf_c.paragraphs[0]
        p1.text = head.upper()
        p1.font.name = FONT_MONO
        p1.font.size = Pt(10)
        p1.font.color.rgb = TEXT_MUTED
        
        p2 = tf_c.add_paragraph()
        p2.text = val
        p2.font.name = FONT_HEAD
        p2.font.size = Pt(14)
        p2.font.bold = True
        p2.font.color.rgb = col
        p2.space_before = Pt(6)
        
    add_footer(s1)

    # ---------------- SLIDE 2: ПРОБЛЕМА 2026 ----------------
    s2 = prs.slides.add_slide(layout)
    set_slide_bg(s2, prs)
    add_header(s2, 2, "Социальная инженерия стала промышленной", 
               "Злоумышленники комбинируют психологическое давление с AI-инструментами и техническим контролем над устройством.")
    
    cols = [
        ("AI Voice & Deepfake", "Синтез голоса близких или руководителей в Telegram/WhatsApp, звонки с подменой номера (VoIP).", AMBER, "УГРОЗА 1"),
        ("Remote Access & AI-Agents", "Управление через AnyDesk / RustDesk или скрытые AI-агенты и автокликеры прямо в сессии.", RED, "УГРОЗА 2"),
        ("Mule Ring & Дроп-сети", "Мгновенное распыление похищенных средств через сеть дроп-счетов и криптовалютные шлюзы.", RED, "УГРОЗА 3"),
    ]
    for i, (title, desc, color, tag_val) in enumerate(cols):
        x = Inches(0.8 + i * 3.98)
        card = add_card(s2, x, Inches(1.8), Inches(3.78), Inches(3.3), accent_color=color)
        tf_c = card.text_frame
        
        p_t = tf_c.paragraphs[0]
        p_t.text = tag_val
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(10)
        p_t.font.bold = True
        p_t.font.color.rgb = color
        
        p_title = tf_c.add_paragraph()
        p_title.text = title
        p_title.font.name = FONT_HEAD
        p_title.font.size = Pt(16)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE
        p_title.space_after = Pt(8)
        
        p_desc = tf_c.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(11.5)
        p_desc.font.color.rgb = TEXT_MUTED

    # Bottom Callout: Key risks
    bot_card = add_card(s2, Inches(0.8), Inches(5.35), Inches(11.733), Inches(1.3), accent_color=PRIMARY_GREEN)
    tf_b = bot_card.text_frame
    p_b1 = tf_b.paragraphs[0]
    p_b1.text = "КЛЮЧЕВЫЕ ВЫЗОВЫ ДЛЯ БАНКА ЦЕНТРКРЕДИТ:"
    p_b1.font.name = FONT_MONO
    p_b1.font.size = Pt(10)
    p_b1.font.bold = True
    p_b1.font.color.rgb = PRIMARY_GREEN
    
    p_b2 = tf_b.add_paragraph()
    p_b2.text = "• Под давлением жертва САМА авторизуется и нажимает «Подтвердить» — классический антифрод бессилен.\n• Риски: потеря доверия клиентов, колоссальный репутационный ущерб, ужесточение требований регулятора (НБРК)."
    p_b2.font.name = FONT_BODY
    p_b2.font.size = Pt(11.5)
    p_b2.font.color.rgb = TEXT_WHITE
    p_b2.space_before = Pt(4)

    add_footer(s2)

    # ---------------- SLIDE 3: РЕШЕНИЕ ----------------
    s3 = prs.slides.add_slide(layout)
    set_slide_bg(s3, prs)
    add_header(s3, 3, "BCC Shield AI — защита, которая думает вместе с клиентом", 
               "Мы останавливаем не просто подозрительную транзакцию, а весь сценарий атаки.")
    
    # Left Hero Box
    c_left = add_card(s3, Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.8), accent_color=PRIMARY_GREEN)
    tf_l = c_left.text_frame
    p0 = tf_l.paragraphs[0]
    p0.text = "ФИЛОСОФИЯ ЗАЩИТЫ НОВОГО ПОКОЛЕНИЯ"
    p0.font.name = FONT_MONO
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = PRIMARY_GREEN
    
    p_quote = tf_l.add_paragraph()
    p_quote.text = "«Мы останавливаем не транзакцию, а сценарий атаки»"
    p_quote.font.name = FONT_HEAD
    p_quote.font.size = Pt(20)
    p_quote.font.bold = True
    p_quote.font.color.rgb = TEXT_WHITE
    p_quote.space_after = Pt(12)
    
    feats = [
        ("Dynamic Risk Score (0–100)", "Непрерывная адаптивная оценка сессии в реальном времени."),
        ("AI Explainability", "Мгновенное и понятное объяснение причин срабатывания как клиенту, так и офицеру безопасности SOC."),
        ("Адаптивный барьер", "Умная градация: пауза + самопроверка, cool-down, звонок близкому или блокировка."),
        ("Ловушка «Тихий щит»", "Дезориентация мошенника при удалённом доступе с полной сохранностью баланса.")
    ]
    for h, d in feats:
        ph = tf_l.add_paragraph()
        ph.text = f"✔ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(12.5)
        ph.font.bold = True
        ph.font.color.rgb = PRIMARY_GREEN
        ph.space_before = Pt(4)
        
        pd = tf_l.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # Right Box - 3 Ecosystem Layers
    rings = [
        ("1. Уровень клиента", "Когнитивный барьер", "Умный чек-лист, 15-секундная пауза отрезвления, кнопка «Стоп-кредит» и Доверенный контакт.", BLUE),
        ("2. Уровень SOC Банка", "Центр реагирования", "Детальная карточка инцидента, объяснение вклада баллов (+38 AI Voice), граф связей.", PRIMARY_GREEN),
        ("3. Экосистема регулятора", "Национальный контур", "Автоматическая передача сигнатур и дроп-счетов в Антифрод-центр НБРК и операторам связи.", PURPLE),
    ]
    for i, (r_title, r_sub, r_desc, r_col) in enumerate(rings):
        y = Inches(1.8 + i * 1.63)
        c = add_card(s3, Inches(6.5), y, Inches(6.033), Inches(1.48), accent_color=r_col)
        tf_c = c.text_frame
        p_r0 = tf_c.paragraphs[0]
        p_r0.text = r_sub.upper()
        p_r0.font.name = FONT_MONO
        p_r0.font.size = Pt(9.5)
        p_r0.font.bold = True
        p_r0.font.color.rgb = r_col
        
        p_r1 = tf_c.add_paragraph()
        p_r1.text = r_title
        p_r1.font.name = FONT_HEAD
        p_r1.font.size = Pt(14)
        p_r1.font.bold = True
        p_r1.font.color.rgb = TEXT_WHITE
        
        p_r2 = tf_c.add_paragraph()
        p_r2.text = r_desc
        p_r2.font.name = FONT_BODY
        p_r2.font.size = Pt(11)
        p_r2.font.color.rgb = TEXT_MUTED

    add_footer(s3)

    # ---------------- SLIDE 4: КЛИЕНТСКИЙ ОПЫТ ----------------
    s4 = prs.slides.add_slide(layout)
    set_slide_bg(s4, prs)
    add_header(s4, 4, "Как выглядит защита для клиента: Когнитивный барьер", 
               "Клиент видит прозрачную причину вмешательства и выходит из-под гипнотического давления злоумышленника.")
    
    steps_cx = [
        ("Старт операции", "1.5 сек", "Фоновый скоринг сессии, телеметрия движений и анализ сетевого окружения.", BLUE),
        ("Риск 40–75", "Когнитивная пауза", "Экран замораживается на 15 секунд. Выводятся 3 чекбокса самопроверки («Вам звонили из банка?», «Просят оформить кредит?»).", AMBER),
        ("Риск 76+", "Cool-Down (2 часа)", "Операция замораживается. Разбор причин и предложение безопасных шагов защиты сбережений.", RED),
        ("Доверенный контакт", "Помощь близких", "Мгновенная отправка запроса маме/брату/партнеру на подтверждение транзакции.", PRIMARY_GREEN),
    ]
    for i, (title, badge, desc, col) in enumerate(steps_cx):
        x = Inches(0.8 + i * 2.98)
        card = add_card(s4, x, Inches(1.8), Inches(2.82), Inches(3.6), accent_color=col)
        tf_c = card.text_frame
        
        p_b = tf_c.paragraphs[0]
        p_b.text = badge.upper()
        p_b.font.name = FONT_MONO
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = col
        p_b.space_after = Pt(4)
        
        p_t = tf_c.add_paragraph()
        p_t.text = title
        p_t.font.name = FONT_HEAD
        p_t.font.size = Pt(15)
        p_t.font.bold = True
        p_t.font.color.rgb = TEXT_WHITE
        p_t.space_after = Pt(8)
        
        p_d = tf_c.add_paragraph()
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_MUTED

    # Bottom Key Action Pill
    bot_pill = add_card(s4, Inches(0.8), Inches(5.65), Inches(11.733), Inches(1.05), accent_color=PRIMARY_GREEN)
    tf_bp = bot_pill.text_frame
    p_bp = tf_bp.paragraphs[0]
    p_bp.text = "КНОПКА ЭКСТРЕННОЙ БЕЗОПАСНОСТИ В 1 КЛИК:"
    p_bp.font.name = FONT_MONO
    p_bp.font.size = Pt(10)
    p_bp.font.bold = True
    p_bp.font.color.rgb = PRIMARY_GREEN
    
    p_bp2 = tf_bp.add_paragraph()
    p_bp2.text = "Клиент может нажать «Стоп-кредит + Инцидент в Антифрод-центр НБРК», моментально заблокировав кредитный потенциал во всех банках РК."
    p_bp2.font.name = FONT_BODY
    p_bp2.font.size = Pt(11.5)
    p_bp2.font.color.rgb = TEXT_WHITE

    add_footer(s4)

    # ---------------- SLIDE 5: ЛОГИКА ЗАЩИТЫ 2026 ----------------
    s5 = prs.slides.add_slide(layout)
    set_slide_bg(s5, prs)
    add_header(s5, 5, "Логика защиты 2026: Dynamic Risk Score + Explainability", 
               "Интеллектуальная агрегация контекстных факторов и объяснимость каждого балла для офицера безопасности.")
    
    # Left Box: Risk Scale
    c_sc = add_card(s5, Inches(0.8), Inches(1.8), Inches(5.5), Inches(4.8), accent_color=PRIMARY_GREEN)
    tf_sc = c_sc.text_frame
    p0 = tf_sc.paragraphs[0]
    p0.text = "ДИНАМИЧЕСКАЯ ШКАЛА ОЦЕНКИ РИСКА"
    p0.font.name = FONT_MONO
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = PRIMARY_GREEN
    
    scale_items = [
        ("0 – 39", "Безопасно", PRIMARY_GREEN, "Штатный перевод. Профиль соответствует 12-месячной истории клиента."),
        ("40 – 75", "Подозрительно", AMBER, "Когнитивная пауза (15 сек), чек-лист самопроверки, подтверждение через доверенное лицо."),
        ("76 – 100", "Критический фрод", RED, "Блокировка / Cool-down на 2 часа, активация ловушки «Тихий щит», сигнал в SOC.")
    ]
    for rng, name, col, desc in scale_items:
        pr = tf_sc.add_paragraph()
        pr.text = f"{rng} · {name.upper()}"
        pr.font.name = FONT_MONO
        pr.font.size = Pt(13)
        pr.font.bold = True
        pr.font.color.rgb = col
        pr.space_before = Pt(8)
        
        pd = tf_sc.add_paragraph()
        pd.text = desc
        pd.font.name = FONT_BODY
        pd.font.size = Pt(11)
        pd.font.color.rgb = TEXT_MUTED

    # Right Box: Factor Weights Breakdown
    c_fc = add_card(s5, Inches(6.5), Inches(1.8), Inches(6.033), Inches(4.8), accent_color=BLUE)
    tf_fc = c_fc.text_frame
    p_ft = tf_fc.paragraphs[0]
    p_ft.text = "ПРИМЕР РАСКЛАДКИ ФАКТОРОВ В КАРТОЧКЕ ИНЦИДЕНТА (EXPLAINABILITY)"
    p_ft.font.name = FONT_MONO
    p_ft.font.size = Pt(10)
    p_ft.font.bold = True
    p_ft.font.color.rgb = BLUE
    
    factors = [
        ("+38", "AI Voice & Deepfake Detection", "Выявлены маркеры синтеза речи в сопутствующем VoIP-вызове (вероятность 89%)", RED),
        ("+34", "Biometrics 2.0 (Gyro & Touch)", "Аномальный наклон устройства (строго вертикально) + задержка касаний под диктовку", RED),
        ("+32", "Telco-Bank Fusion (Операторы)", "Активный нетарифицируемый звонок параллельно с отправкой формы перевода", AMBER),
        ("+28", "Mule Ring Detection", "Получатель зарегистрирован 3 дня назад, числится в серой базе транзитных дропов", AMBER),
        ("−15", "Trusted History Compensator", "Постоянное доверенное устройство, используемое клиентом > 180 дней", PRIMARY_GREEN),
    ]
    for pts, fname, fdesc, fcol in factors:
        p_pt = tf_fc.add_paragraph()
        p_pt.text = f"{pts} баллов  ·  {fname}"
        p_pt.font.name = FONT_HEAD
        p_pt.font.size = Pt(12.5)
        p_pt.font.bold = True
        p_pt.font.color.rgb = fcol
        p_pt.space_before = Pt(6)
        
        p_pdesc = tf_fc.add_paragraph()
        p_pdesc.text = fdesc
        p_pdesc.font.name = FONT_BODY
        p_pdesc.font.size = Pt(10)
        p_pdesc.font.color.rgb = TEXT_MUTED

    add_footer(s5)

    # ---------------- SLIDE 6: ТЕХНОЛОГИИ 2026 ----------------
    s6 = prs.slides.add_slide(layout)
    set_slide_bg(s6, prs)
    add_header(s6, 6, "Технологический стек защиты нового поколения", 
               "Глубокая синергия мобильной телеметрии, AI-анализа речи, графовых баз и кросс-индустриальных данных.")
    
    techs = [
        ("Telco-Bank Fusion", "Связка с операторами", "Сигналы о параллельном звонке, переадресации, подмене SIM-карты (SIM-swap) и вызовах с поддельных номеров (1414 / полиция).", PRIMARY_GREEN),
        ("AI Voice & Deepfake", "Анализ синтеза речи", "Детекция синтезированного аудио и клонированных голосов родственников с точностью до 89% прямо во время диалога.", RED),
        ("Biometrics 2.0", "Поведенческая физика", "Снятие гироскопа, угла наклона смартфона, давления пальца, интервалов между нажатиями (Touch Δt), детекция диктовки.", BLUE),
        ("Граф Mule Ring", "Анализ связей", "Мгновенное обнаружение топологии транзита: «жертвы → узел-аккумулятор → банкоматный вывод» через графовые алгоритмы.", AMBER),
        ("Поведенческий профиль", "12-месячный бейзлайн", "Сравнение текущей операции с привычным цифровым почерком: время суток, типичные суммы, категории получателей.", PURPLE),
        ("Zero-Trust Device SDK", "Аттестация среды", "Проверка целостности ОС, детект Magisk, Frida, скрытых сервисов удалённого доступа (RustDesk / AnyDesk).", EMERALD),
    ]
    for i, (t_name, t_sub, t_desc, t_col) in enumerate(techs):
        x = Inches(0.8 + (i % 3) * 3.98)
        y = Inches(1.8 + (i // 3) * 2.45)
        card = add_card(s6, x, y, Inches(3.78), Inches(2.25), accent_color=t_col)
        tf_c = card.text_frame
        
        p_sub = tf_c.paragraphs[0]
        p_sub.text = t_sub.upper()
        p_sub.font.name = FONT_MONO
        p_sub.font.size = Pt(9.5)
        p_sub.font.bold = True
        p_sub.font.color.rgb = t_col
        
        p_title = tf_c.add_paragraph()
        p_title.text = t_name
        p_title.font.name = FONT_HEAD
        p_title.font.size = Pt(14)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE
        p_title.space_after = Pt(4)
        
        p_desc = tf_c.add_paragraph()
        p_desc.text = t_desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = TEXT_MUTED

    add_footer(s6)

    # ---------------- SLIDE 7: ТИХИЙ ЩИТ ----------------
    s7 = prs.slides.add_slide(layout)
    set_slide_bg(s7, prs)
    add_header(s7, 7, "«Тихий щит» — ловушка для мошенника", 
               "Инновационный механизм обмана атакующего при попытке захвата сессии или удалённого доступа.")
    
    # Left Card: Attacker view
    c_att = add_card(s7, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), accent_color=RED)
    tf_att = c_att.text_frame
    p_a0 = tf_att.paragraphs[0]
    p_a0.text = "ЧТО ВИДИТ МОШЕННИК (ANYDESK / AI-ОВЕРЛЕЙ)"
    p_a0.font.name = FONT_MONO
    p_a0.font.size = Pt(10)
    p_a0.font.bold = True
    p_a0.font.color.rgb = RED
    
    att_points = [
        ("Фейковый экран успеха", "Злоумышленнику транслируется реалистичный экран «Перевод успешно выполнен. Чек отправлен на почту»."),
        ("Иллюзия победы", "Мошенник уверен, что цель достигнута, и завершает звонок / сессию, не предпринимая агрессивных повторных попыток."),
        ("Защита баланса", "Фактическая транзакция НЕ отправляется в платёжный шлюз. Средства клиента заблокированы в безопасном резерве."),
        ("Фиксация улик", "Пока мошенник спокоен, система фиксирует IP-адрес, параметры соединения, идентификатор удалённой программы и реквизиты дропа.")
    ]
    for h, d in att_points:
        ph = tf_att.add_paragraph()
        ph.text = f"✖ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)
        
        pd = tf_att.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # Right Card: Real SOC & Client view
    c_soc = add_card(s7, Inches(6.6), Inches(1.8), Inches(5.933), Inches(4.8), accent_color=PRIMARY_GREEN)
    tf_soc = c_soc.text_frame
    p_s0 = tf_soc.paragraphs[0]
    p_s0.text = "ЧТО ПРОИСХОДИТ В РЕАЛЬНОСТИ (КЛИЕНТ & SOC)"
    p_s0.font.name = FONT_MONO
    p_s0.font.size = Pt(10)
    p_s0.font.bold = True
    p_s0.font.color.rgb = PRIMARY_GREEN
    
    soc_points = [
        ("Владелец видит правду", "По специальному жесту или физической кнопке на смартфоне открывается экран реальной защиты: «Мы спасли ваши деньги»."),
        ("Мгновенный Safe Mode", "Автоматически замораживаются цифровые карты и доступ в мобильный банк до подтверждения личности."),
        ("Оповещение близких", "Доверенный контакт получает SMS/Push: «С аккаунтом родственника зафиксирована подозрительная активность, свяжитесь с ним»."),
        ("Результат", "Деньги клиента сохранены на 100%, сценарий атаки сорван, банк предотвратил многомиллионный ущерб.")
    ]
    for h, d in soc_points:
        ph = tf_soc.add_paragraph()
        ph.text = f"✔ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)
        
        pd = tf_soc.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    add_footer(s7)

    # ---------------- SLIDE 8: ЭКОСИСТЕМА И РЕГУЛЯТОР ----------------
    s8 = prs.slides.add_slide(layout)
    set_slide_bg(s8, prs)
    add_header(s8, 8, "Экосистема: От клиента до Антифрод-центра НБРК", 
               "Комплексный контур безопасности объединяет пользователя, банк, операторов связи и регулятора.")
    
    blocks = [
        ("Доверенный контакт", "Человеческий фактор доверия", "При обнаружении давления или синтеза речи система запрашивает одобрение близкого родственника (мамы, супруга).", PRIMARY_GREEN),
        ("Кнопка «Стоп-кредит»", "Мгновенная блокировка займов", "Клиент или SOC активирует добровольный запрет на кредиты во всех БВУ через интеграцию с ПКБ / ГКБ.", BLUE),
        ("Граф Mule Ring", "Аналитика связей", "Выявление транзитных узлов и передача скомпрометированных карт в единый реестр дропперов Казахстана.", AMBER),
        ("Антифрод-центр НБРК", "Национальный обмен", "Real-time синхронизация инцидентов с Национальным Банком РК и операторами связи для блокировки номеров.", PURPLE),
    ]
    for i, (b_title, b_sub, b_desc, b_col) in enumerate(blocks):
        x = Inches(0.8 + (i % 2) * 5.95)
        y = Inches(1.8 + (i // 2) * 2.45)
        card = add_card(s8, x, y, Inches(5.75), Inches(2.25), accent_color=b_col)
        tf_c = card.text_frame
        
        p_t = tf_c.paragraphs[0]
        p_t.text = b_sub.upper()
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(10)
        p_t.font.bold = True
        p_t.font.color.rgb = b_col
        
        p_h = tf_c.add_paragraph()
        p_h.text = b_title
        p_h.font.name = FONT_HEAD
        p_h.font.size = Pt(16)
        p_h.font.bold = True
        p_h.font.color.rgb = TEXT_WHITE
        p_h.space_after = Pt(4)
        
        p_d = tf_c.add_paragraph()
        p_d.text = b_desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(11)
        p_d.font.color.rgb = TEXT_MUTED

    add_footer(s8)

    # ---------------- SLIDE 9: ЭФФЕКТ ДЕМО ----------------
    s9 = prs.slides.add_slide(layout)
    set_slide_bg(s9, prs)
    add_header(s9, 9, "Эффективность решения (смоделированные данные стенда)", 
               "Результаты симуляции защиты на демонстрационном стенде Банка ЦентрКредит.")
    
    # 3 Big Metrics
    kpis = [
        ("14 250 000 ₸", "Сэкономлено клиентам", "Предотвращенный ущерб за смену", PRIMARY_GREEN),
        ("38 операций", "Остановлено атак", "Заморожено через Cool-down и блокировки", RED),
        ("1 240 транзакций", "Проверено потока", "97.2% пропущены без малейшего трения", BLUE),
    ]
    for i, (val, title, sub, col) in enumerate(kpis):
        x = Inches(0.8 + i * 3.98)
        c = add_card(s9, x, Inches(1.8), Inches(3.78), Inches(2.1), accent_color=col)
        tf_k = c.text_frame
        p_kv = tf_k.paragraphs[0]
        p_kv.text = val
        p_kv.font.name = FONT_MONO
        p_kv.font.size = Pt(22)
        p_kv.font.bold = True
        p_kv.font.color.rgb = col
        p_kv.space_after = Pt(4)
        
        p_kt = tf_k.add_paragraph()
        p_kt.text = title
        p_kt.font.name = FONT_HEAD
        p_kt.font.size = Pt(14)
        p_kt.font.bold = True
        p_kt.font.color.rgb = TEXT_WHITE
        
        p_ks = tf_k.add_paragraph()
        p_ks.text = sub
        p_ks.font.name = FONT_BODY
        p_ks.font.size = Pt(10.5)
        p_ks.font.color.rgb = TEXT_MUTED

    # Bottom Split Chart info
    c_split = add_card(s9, Inches(0.8), Inches(4.15), Inches(11.733), Inches(2.5), accent_color=SURFACE_BORDER)
    tf_sp = c_split.text_frame
    p_spt = tf_sp.paragraphs[0]
    p_spt.text = "РАСПРЕДЕЛЕНИЕ ТИПОВ АТАК ЗА СМЕНУ (СИМУЛЯЦИЯ ДЕМО):"
    p_spt.font.name = FONT_MONO
    p_spt.font.size = Pt(10)
    p_spt.font.bold = True
    p_spt.font.color.rgb = PRIMARY_GREEN
    
    categories = [
        ("42%", "Telco + VoIP давление", "Звонки якобы от службы безопасности / полиции параллельно с мобильным приложением."),
        ("26%", "AI Voice & Deepfake", "Синтезированные аудиосообщения от имени близких с просьбой срочного перевода."),
        ("18%", "AnyDesk & Управление экраном", "Попытки перехвата управления сессией и автокликеры — нейтрализованы «Тихим щитом»."),
        ("14%", "Кредит под давлением", "Попытки снятия кредитных лимитов — остановлены тайм-аутом и кнопкой «Стоп-кредит».")
    ]
    for pct, cname, cdesc in categories:
        p_row = tf_sp.add_paragraph()
        p_row.text = f"{pct}  ·  {cname} — {cdesc}"
        p_row.font.name = FONT_BODY
        p_row.font.size = Pt(11)
        p_row.font.color.rgb = TEXT_WHITE
        p_row.space_before = Pt(5)

    add_footer(s9)

    # ---------------- SLIDE 10: СЛЕДУЮЩИЕ ШАГИ ----------------
    s10 = prs.slides.add_slide(layout)
    set_slide_bg(s10, prs)
    add_header(s10, 10, "BCC Shield AI готов к пилотному внедрению", 
               "Пошаговый план интеграции и калибровки решений на инфраструктуре Банка ЦентрКредит.")
    
    # 4 Steps Cards
    roadmap = [
        ("01", "Ограниченный контур", "Запуск пилота на сегменте операций повышенного риска (первые переводы новым получателям, крупные суммы).", PRIMARY_GREEN),
        ("02", "Telco & НБРК интеграция", "Подключение обмена метаданными операторов (активный вызов 1414) и API Антифрод-центра НБРК.", BLUE),
        ("03", "Калибровка порогов", "Тонкая настройка порогов когнитивных пауз и тайм-аутов на обезличенных исторических данных банка.", AMBER),
        ("04", "Развёртывание фич", "Активация «Тихого щита», Доверенного контакта и сервиса «Стоп-кредит» в клиентском приложении BCC.", PURPLE),
    ]
    for i, (num, r_title, r_desc, r_col) in enumerate(roadmap):
        x = Inches(0.8 + i * 2.98)
        card = add_card(s10, x, Inches(1.8), Inches(2.82), Inches(3.6), accent_color=r_col)
        tf_r = card.text_frame
        
        p_num = tf_r.paragraphs[0]
        p_num.text = num
        p_num.font.name = FONT_MONO
        p_num.font.size = Pt(24)
        p_num.font.bold = True
        p_num.font.color.rgb = r_col
        p_num.space_after = Pt(2)
        
        p_rt = tf_r.add_paragraph()
        p_rt.text = r_title
        p_rt.font.name = FONT_HEAD
        p_rt.font.size = Pt(14)
        p_rt.font.bold = True
        p_rt.font.color.rgb = TEXT_WHITE
        p_rt.space_after = Pt(6)
        
        p_rd = tf_r.add_paragraph()
        p_rd.text = r_desc
        p_rd.font.name = FONT_BODY
        p_rd.font.size = Pt(11)
        p_rd.font.color.rgb = TEXT_MUTED

    # Bottom Contact Box
    bot_c = add_card(s10, Inches(0.8), Inches(5.65), Inches(11.733), Inches(1.05), accent_color=PRIMARY_GREEN)
    tf_bc = bot_c.text_frame
    p_bc1 = tf_bc.paragraphs[0]
    p_bc1.text = "СЛЕДУЮЩИЙ ШАГ: ОБСУЖДЕНИЕ ПИЛОТНОГО КОНТУРА И ТЕХНИЧЕСКОЙ КАРТЫ ИНТЕГРАЦИИ"
    p_bc1.font.name = FONT_MONO
    p_bc1.font.size = Pt(10)
    p_bc1.font.bold = True
    p_bc1.font.color.rgb = PRIMARY_GREEN
    
    p_bc2 = tf_bc.add_paragraph()
    p_bc2.text = "Команда проекта готова предоставить демо-стенд, спецификации API и провести технический воркшоп для команды информационной безопасности BCC. Q&A."
    p_bc2.font.name = FONT_BODY
    p_bc2.font.size = Pt(11.5)
    p_bc2.font.color.rgb = TEXT_WHITE

    add_footer(s10)

    out_path = r"c:\Users\user\Downloads\bcc-hackathon\bcc-hackathon\presentations\BCC_Shield_AI_PitchDeck.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}")

build_shield_ai()
