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
CYAN = RGBColor(6, 182, 212)         # #06B6D4

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

def add_header(slide, slide_num, title, subtitle=None, tag="BCC CYBERGUARD"):
    tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(8), Inches(0.35))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = f"{tag}  ·  SLIDE {slide_num:02d}"
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

def build_cyberguard():
    prs = create_base_presentation()
    layout = prs.slide_layouts[6]
    
    # ---------------- SLIDE 1 ----------------
    s1 = prs.slides.add_slide(layout)
    set_slide_bg(s1, prs)
    
    # Hero Title Box
    title_box = s1.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11.5), Inches(3.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    
    p_tag = tf.paragraphs[0]
    p_tag.text = "NEXT-GEN FINTECH SECURITY · HACKATHON PROTOTYPE"
    p_tag.font.name = FONT_MONO
    p_tag.font.size = Pt(12)
    p_tag.font.bold = True
    p_tag.font.color.rgb = PRIMARY_GREEN
    p_tag.space_after = Pt(14)
    
    p_main = tf.add_paragraph()
    p_main.text = "BCC CyberGuard"
    p_main.font.name = FONT_HEAD
    p_main.font.size = Pt(46)
    p_main.font.bold = True
    p_main.font.color.rgb = TEXT_WHITE
    p_main.space_after = Pt(6)
    
    p_sub = tf.add_paragraph()
    p_sub.text = "Real-Time Threat & Payment Integrity Engine"
    p_sub.font.name = FONT_HEAD
    p_sub.font.size = Pt(20)
    p_sub.font.color.rgb = BLUE
    p_sub.space_after = Pt(14)
    
    p_desc = tf.add_paragraph()
    p_desc.text = "Высокоскоростной движок защиты платёжных транзакций, целостности реквизитов и API-сессий Банка ЦентрКредит с задержкой принятия решений < 20 мс."
    p_desc.font.name = FONT_BODY
    p_desc.font.size = Pt(13)
    p_desc.font.color.rgb = TEXT_MUTED

    # 4 Feature Pills
    features = [
        ("Мгновенный вердикт", "< 20 мс задержка", PRIMARY_GREEN),
        ("Три действия", "ALLOW · STEP-UP · BLOCK", BLUE),
        ("Анализ угроз", "4 модуля в реальном времени", AMBER),
        ("Готовность", "Рабочий стенд (Demo)", EMERALD),
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

    # ---------------- SLIDE 2: ПРОБЛЕМА ----------------
    s2 = prs.slides.add_slide(layout)
    set_slide_bg(s2, prs)
    add_header(s2, 2, "Атаки на платёжный контур происходят до и во время операции", 
               "Традиционный антифрод видит только финальную транзакцию, пропуская подмену данных на клиенте и манипуляции сессией.")
    
    threats = [
        ("Подмена реквизитов & QR", "Инъекция чужого IBAN или QR-кода на этапе подтверждения клиентом в браузере или приложении.", RED, "КРИТИЧНО"),
        ("Фермы дропперов & Mule", "Массовые серии мелких переводов, дробление сумм ниже лимитов и моментальный транзит средств.", RED, "МАСШТАБ"),
        ("Эмуляторы & Headless", "Запуск сессий через эмуляторы, скомпрометированный runtime (Frida, Magisk, обход Play Integrity).", AMBER, "СКРЫТНОСТЬ"),
        ("Смена TLS-отпечатка (JA3)", "Ротация прокси, VPN и смена фингерпринта внутри одной активной банковской сессии.", BLUE, "АНАЛИЗ")
    ]
    for i, (title, desc, color, tag_val) in enumerate(threats):
        x = Inches(0.8 + (i % 2) * 5.95)
        y = Inches(1.8 + (i // 2) * 2.4)
        card = add_card(s2, x, y, Inches(5.75), Inches(2.15), accent_color=color)
        tf_c = card.text_frame
        
        # Tag
        p_t = tf_c.paragraphs[0]
        p_t.text = tag_val
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(10)
        p_t.font.bold = True
        p_t.font.color.rgb = color
        
        p_title = tf_c.add_paragraph()
        p_title.text = title
        p_title.font.name = FONT_HEAD
        p_title.font.size = Pt(15)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE
        p_title.space_after = Pt(6)
        
        p_desc = tf_c.add_paragraph()
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(11.5)
        p_desc.font.color.rgb = TEXT_MUTED
        
    add_footer(s2)

    # ---------------- SLIDE 3: РЕШЕНИЕ В ОДНОМ ВЗГЛЯДЕ ----------------
    s3 = prs.slides.add_slide(layout)
    set_slide_bg(s3, prs)
    add_header(s3, 3, "BCC CyberGuard — единый real-time движок", 
               "Консолидирует независимые сигналы безопасности в единый индекс опасности за миллисекунды.")
    
    # Left Box - Big Score Engine
    left_card = add_card(s3, Inches(0.8), Inches(1.8), Inches(4.5), Inches(4.8), accent_color=PRIMARY_GREEN)
    tf_l = left_card.text_frame
    p0 = tf_l.paragraphs[0]
    p0.text = "CORE METRICS"
    p0.font.name = FONT_MONO
    p0.font.size = Pt(11)
    p0.font.color.rgb = PRIMARY_GREEN
    
    p1 = tf_l.add_paragraph()
    p1.text = "Threat Score"
    p1.font.name = FONT_HEAD
    p1.font.size = Pt(22)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_WHITE
    
    p_score = tf_l.add_paragraph()
    p_score.text = "0 – 100"
    p_score.font.name = FONT_MONO
    p_score.font.size = Pt(44)
    p_score.font.bold = True
    p_score.font.color.rgb = PRIMARY_GREEN
    p_score.space_after = Pt(10)
    
    bullets = [
        "4 независимых модуля риска",
        "Задержка принятия решения: 6–20 мс",
        "Прозрачность для SOC: точная причина",
        "Интеграция в каналы: QR, P2P, IBAN, API"
    ]
    for b in bullets:
        pb = tf_l.add_paragraph()
        pb.text = f"✔ {b}"
        pb.font.name = FONT_BODY
        pb.font.size = Pt(11.5)
        pb.font.color.rgb = TEXT_WHITE
        pb.space_after = Pt(5)

    # Right 3 Action Cards
    actions = [
        ("ALLOW (Score < 30)", "Зелёный коридор", "Операция пропускается мгновенно без трения для добросовестного клиента. 95%+ трафика.", PRIMARY_GREEN),
        ("STEP-UP CHALLENGE (30–74)", "Адаптивная проверка", "Требуется подтверждение: аппаратный ключ FIDO2, WebAuthn или биометрия. Трение только при подозрении.", AMBER),
        ("AUTO-BLOCK & FREEZE (≥ 75)", "Мгновенная блокировка", "Списание отменяется, сессия замораживается, инцидент с полным отпечатком отправляется в SIEM/SOC.", RED),
    ]
    for i, (act_title, badge, act_desc, act_col) in enumerate(actions):
        y = Inches(1.8 + i * 1.63)
        c = add_card(s3, Inches(5.5), y, Inches(7.0), Inches(1.48), accent_color=act_col)
        tf_c = c.text_frame
        p_a0 = tf_c.paragraphs[0]
        p_a0.text = badge.upper()
        p_a0.font.name = FONT_MONO
        p_a0.font.size = Pt(9.5)
        p_a0.font.bold = True
        p_a0.font.color.rgb = act_col
        
        p_a1 = tf_c.add_paragraph()
        p_a1.text = act_title
        p_a1.font.name = FONT_HEAD
        p_a1.font.size = Pt(14)
        p_a1.font.bold = True
        p_a1.font.color.rgb = TEXT_WHITE
        
        p_a2 = tf_c.add_paragraph()
        p_a2.text = act_desc
        p_a2.font.name = FONT_BODY
        p_a2.font.size = Pt(11)
        p_a2.font.color.rgb = TEXT_MUTED

    add_footer(s3)

    # ---------------- SLIDE 4: АРХИТЕКТУРА 4 МОДУЛЯ ----------------
    s4 = prs.slides.add_slide(layout)
    set_slide_bg(s4, prs)
    add_header(s4, 4, "Архитектура анализа: 4 специализированных модуля", 
               "Комплексный профиль риска формируется по 4 независимым векторам безопасности.")
    
    modules = [
        ("Payment Payload Integrity", "Вес: 0.95", RED, [
            "Сверка криптографических SHA-256 хэшей экрана подтверждения и банковского ядра",
            "Защита от Man-in-the-Browser и подмены QR-кодов",
            "Расхождение хэшей → немедленный BLOCK"
        ]),
        ("Drop-Network Velocity", "Вес: 0.90", RED, [
            "Анализ частоты серийных переводов (RPM)",
            "Выявление веерных переводов и дробления сумм",
            "Мониторинг связанных счетов-получателей"
        ]),
        ("Device & Runtime Attestation", "Вес: 0.85", AMBER, [
            "Детект эмуляторов, рута, Magisk, Frida hook-ов",
            "Проверка аппаратной аттестации Play Integrity / Keystore",
            "Контроль смены SIM-карты и статуса девайса"
        ]),
        ("Network & TLS Fingerprint", "Вес: 0.80", BLUE, [
            "Сверка отпечатка TLS Client Hello (JA3/JA4)",
            "Детект смены сетевого профиля внутри активной сессии",
            "Анализ резидентских прокси, VPN и дата-центров (ASN)"
        ]),
    ]
    for i, (m_title, m_weight, m_col, m_points) in enumerate(modules):
        x = Inches(0.8 + (i % 2) * 5.95)
        y = Inches(1.8 + (i // 2) * 2.45)
        card = add_card(s4, x, y, Inches(5.75), Inches(2.25), accent_color=m_col)
        tf_c = card.text_frame
        
        p_w = tf_c.paragraphs[0]
        p_w.text = m_weight.upper()
        p_w.font.name = FONT_MONO
        p_w.font.size = Pt(10)
        p_w.font.bold = True
        p_w.font.color.rgb = m_col
        
        p_t = tf_c.add_paragraph()
        p_t.text = m_title
        p_t.font.name = FONT_HEAD
        p_t.font.size = Pt(15)
        p_t.font.bold = True
        p_t.font.color.rgb = TEXT_WHITE
        p_t.space_after = Pt(4)
        
        for pt in m_points:
            p_p = tf_c.add_paragraph()
            p_p.text = f"• {pt}"
            p_p.font.name = FONT_BODY
            p_p.font.size = Pt(10.5)
            p_p.font.color.rgb = TEXT_MUTED

    add_footer(s4)

    # ---------------- SLIDE 5: ФОРМУЛА И РЕШЕНИЯ ----------------
    s5 = prs.slides.add_slide(layout)
    set_slide_bg(s5, prs)
    add_header(s5, 5, "Формула объединения рисков и логика решений", 
               "Вероятностное объединение независимых свидетельств атаки без слепых зон.")
    
    # Formula Box
    form_card = add_card(s5, Inches(0.8), Inches(1.8), Inches(11.733), Inches(1.8), accent_color=PRIMARY_GREEN)
    tf_f = form_card.text_frame
    p_ft = tf_f.paragraphs[0]
    p_ft.text = "ВЕРОЯТНОСТНАЯ МОДЕЛЬ ОБЪЕДИНЕНИЯ НЕЗАВИСИМЫХ СИГНАЛОВ (PROBABILISTIC UNION)"
    p_ft.font.name = FONT_MONO
    p_ft.font.size = Pt(10)
    p_ft.font.bold = True
    p_ft.font.color.rgb = PRIMARY_GREEN
    
    p_fm = tf_f.add_paragraph()
    p_fm.text = "Threat Score = 100 × ( 1 − ∏ᵢ (1 − wᵢ × rᵢ) )"
    p_fm.font.name = FONT_MONO
    p_fm.font.size = Pt(24)
    p_fm.font.bold = True
    p_fm.font.color.rgb = TEXT_WHITE
    p_fm.space_after = Pt(6)
    
    p_fd = tf_f.add_paragraph()
    p_fd.text = "где wᵢ — калиброванный вес модуля (0.80 – 0.95), rᵢ — нормализованная оценка риска модуля (0.0 – 1.0). Даже один критический сигнал способен перевести операцию в BLOCK."
    p_fd.font.name = FONT_BODY
    p_fd.font.size = Pt(11.5)
    p_fd.font.color.rgb = TEXT_MUTED

    # 3 Decision Columns below
    decisions = [
        ("< 30", "ALLOW", PRIMARY_GREEN, "Штатный платёж", [
            "Платёж подтверждается прозрачно",
            "Задержка проверки < 10 мс",
            "Нулевое трение для 95%+ операций"
        ]),
        ("30 – 74", "STEP-UP", AMBER, "Повышенный риск", [
            "Запрос дополнительной верификации",
            "FIDO2 / WebAuthn или биометрия",
            "Отсеивает автоматизированный фрод"
        ]),
        ("≥ 75", "AUTO-BLOCK", RED, "Критическая угроза", [
            "Мгновенный отказ и заморозка счёта",
            "Логирование инцидента в SOC",
            "Защита от списания и вывода"
        ]),
    ]
    for i, (rng, dec_name, col, dec_sub, pts) in enumerate(decisions):
        x = Inches(0.8 + i * 3.98)
        card = add_card(s5, x, Inches(3.85), Inches(3.78), Inches(2.75), accent_color=col)
        tf_d = card.text_frame
        
        p_r = tf_d.paragraphs[0]
        p_r.text = f"SCORE {rng}"
        p_r.font.name = FONT_MONO
        p_r.font.size = Pt(11)
        p_r.font.bold = True
        p_r.font.color.rgb = col
        
        p_n = tf_d.add_paragraph()
        p_n.text = dec_name
        p_n.font.name = FONT_HEAD
        p_n.font.size = Pt(20)
        p_n.font.bold = True
        p_n.font.color.rgb = TEXT_WHITE
        
        p_s = tf_d.add_paragraph()
        p_s.text = dec_sub
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = TEXT_MUTED
        p_s.space_after = Pt(8)
        
        for pt in pts:
            pp = tf_d.add_paragraph()
            pp.text = f"✔ {pt}"
            pp.font.name = FONT_BODY
            pp.font.size = Pt(10.5)
            pp.font.color.rgb = TEXT_WHITE
            pp.space_after = Pt(3)

    add_footer(s5)

    # ---------------- SLIDE 6: ЦЕЛОСТНОСТЬ PAYLOAD ----------------
    s6 = prs.slides.add_slide(layout)
    set_slide_bg(s6, prs)
    add_header(s6, 6, "Целостность Payload: Экран подтверждения vs Ядро банка", 
               "Устранение уязвимости, когда клиент подтверждает одну сумму и получателя, а в бэкенд уходят другие.")
    
    # Left Card: Problem Mechanism
    c_l = add_card(s6, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), accent_color=BLUE)
    tf_cl = c_l.text_frame
    p0 = tf_cl.paragraphs[0]
    p0.text = "МЕХАНИКА АТАКИ & ЗАЩИТЫ"
    p0.font.name = FONT_MONO
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = BLUE
    
    points_l = [
        ("Уязвимость экрана", "Вредоносное ПО на устройстве (троян, оверлей или расширение) подменяет реквизиты в момент клика «Оплатить»."),
        ("Крипто-хэширование SHA-256", "Фронтенд в защищённом анклаве вычисляет хэш от реквизитов, отображённых на экране (Сумма + IBAN + Валюта + Timestamp)."),
        ("Сверка в ядре банка", "Бэкенд пересчитывает хэш фактически полученного платёжного пакета. Сверка занимает менее 1 мс."),
        ("Гарантия защиты", "Любое расхождение хотя бы в 1 символ приводит к немедленному статусу MISMATCH и сбросу операции.")
    ]
    for h, d in points_l:
        ph = tf_cl.add_paragraph()
        ph.text = f"▸ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)
        
        pd = tf_cl.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # Right Card: Demo states
    c_r1 = add_card(s6, Inches(6.6), Inches(1.8), Inches(5.933), Inches(2.25), accent_color=PRIMARY_GREEN)
    tf_r1 = c_r1.text_frame
    p_m1 = tf_r1.paragraphs[0]
    p_m1.text = "РЕЗУЛЬТАТ: MATCH (ХЭШИ СОВПАДАЮТ)"
    p_m1.font.name = FONT_MONO
    p_m1.font.size = Pt(11)
    p_m1.font.bold = True
    p_m1.font.color.rgb = PRIMARY_GREEN
    
    p_h1 = tf_r1.add_paragraph()
    p_h1.text = "Screen SHA: e3b0c44298fc1c149afbf4c8996fb924..."
    p_h1.font.name = FONT_MONO
    p_h1.font.size = Pt(10)
    p_h1.font.color.rgb = TEXT_WHITE
    
    p_h2 = tf_r1.add_paragraph()
    p_h2.text = "Core SHA:   e3b0c44298fc1c149afbf4c8996fb924..."
    p_h2.font.name = FONT_MONO
    p_h2.font.size = Pt(10)
    p_h2.font.color.rgb = PRIMARY_GREEN
    
    p_d1 = tf_r1.add_paragraph()
    p_d1.text = "✔ Целостность подтверждена. Клиент подтвердил именно то, что пришло в процессинг."
    p_d1.font.name = FONT_BODY
    p_d1.font.size = Pt(11)
    p_d1.font.color.rgb = TEXT_MUTED
    p_d1.space_before = Pt(4)

    c_r2 = add_card(s6, Inches(6.6), Inches(4.35), Inches(5.933), Inches(2.25), accent_color=RED)
    tf_r2 = c_r2.text_frame
    p_m2 = tf_r2.paragraphs[0]
    p_m2.text = "РЕЗУЛЬТАТ: MISMATCH (ПОДМЕНА ОБНАРУЖЕНА)"
    p_m2.font.name = FONT_MONO
    p_m2.font.size = Pt(11)
    p_m2.font.bold = True
    p_m2.font.color.rgb = RED
    
    p_rh1 = tf_r2.add_paragraph()
    p_rh1.text = "Screen SHA: 7f83b1657ff1fc53b92dc18148a1d65d..."
    p_rh1.font.name = FONT_MONO
    p_rh1.font.size = Pt(10)
    p_rh1.font.color.rgb = TEXT_WHITE
    
    p_rh2 = tf_r2.add_paragraph()
    p_rh2.text = "Core SHA:   4b227777d4dd1fc61c6f884f48641d02..."
    p_rh2.font.name = FONT_MONO
    p_rh2.font.size = Pt(10)
    p_rh2.font.color.rgb = RED
    
    p_rd1 = tf_r2.add_paragraph()
    p_rd1.text = "✖ ВНИМАНИЕ: Реквизиты изменены после подтверждения клиентом! Мгновенная блокировка транзакции."
    p_rd1.font.name = FONT_BODY
    p_rd1.font.size = Pt(11)
    p_rd1.font.color.rgb = RED
    p_rd1.space_before = Pt(4)

    add_footer(s6)

    # ---------------- SLIDE 7: ГРАФ СВЯЗЕЙ И СКОРОСТЬ ----------------
    s7 = prs.slides.add_slide(layout)
    set_slide_bg(s7, prs)
    add_header(s7, 7, "Выявление дроп-сетей: Граф связей и скорость (Velocity)", 
               "Обнаружение аккумуляции и транзита средств через корреляцию параметров серии.")
    
    c_g1 = add_card(s7, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8), accent_color=AMBER)
    tf_g1 = c_g1.text_frame
    p0 = tf_g1.paragraphs[0]
    p0.text = "ПАТТЕРНЫ ДРОП-АККУМУЛЯЦИИ"
    p0.font.name = FONT_MONO
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = AMBER
    
    g_points = [
        ("Граф «Многие-к-Одному»", "Несколько скомпрометированных счетов или жертв переводят средства на один дроп-аккумулятор."),
        ("Высокая скорость серии (RPM)", "Резкий всплеск количества операций в минуту (Requests Per Minute > 20) с нового или подозрительного устройства."),
        ("Дробление ниже порогов (Smurfing)", "Разбиение крупной суммы на серию платежей по 40 000 – 50 000 ₸ для обхода стандартных AML-правил."),
        ("Реакция модуля", "Drop-Network Velocity мгновенно повышает Score до 85+, вызывая превентивную блокировку всей цепочки.")
    ]
    for h, d in g_points:
        ph = tf_g1.add_paragraph()
        ph.text = f"▸ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)
        
        pd = tf_g1.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = TEXT_MUTED

    # Right Card: Graph representation metrics
    c_g2 = add_card(s7, Inches(6.6), Inches(1.8), Inches(5.933), Inches(4.8), accent_color=SURFACE_BORDER)
    tf_g2 = c_g2.text_frame
    p_gt = tf_g2.paragraphs[0]
    p_gt.text = "ИНДИКАТОРЫ ДРОП-СЕТИ В РЕАЛЬНОМ ВРЕМЕНИ"
    p_gt.font.name = FONT_MONO
    p_gt.font.size = Pt(10)
    p_gt.font.bold = True
    p_gt.font.color.rgb = BLUE
    
    metrics = [
        ("Скорость потока (Velocity)", "48 RPM", "Порог нормы: < 3 RPM", RED),
        ("Связанные счета-получатели", "14 дропов", "Выявлена группа с общим IP/ASN", RED),
        ("Паттерн сумм", "Дробление", "Серия платежей по 49 500 ₸", AMBER),
        ("Вердикт Velocity Engine", "CRITICAL DROP-RING", "Риск модуля: 0.94 / 1.00", RED),
    ]
    for lbl, val, sub, col in metrics:
        p_l = tf_g2.add_paragraph()
        p_l.text = lbl.upper()
        p_l.font.name = FONT_MONO
        p_l.font.size = Pt(9.5)
        p_l.font.color.rgb = TEXT_MUTED
        p_l.space_before = Pt(10)
        
        p_v = tf_g2.add_paragraph()
        p_v.text = f"{val}  ·  {sub}"
        p_v.font.name = FONT_HEAD
        p_v.font.size = Pt(13)
        p_v.font.bold = True
        p_v.font.color.rgb = col

    add_footer(s7)

    # ---------------- SLIDE 8: ЦЕННОСТЬ ДЛЯ БАНКА И SOC ----------------
    s8 = prs.slides.add_slide(layout)
    set_slide_bg(s8, prs)
    add_header(s8, 8, "Ценность для Банка ЦентрКредит и аналитиков SOC", 
               "Баланс между бескомпромиссной безопасностью и комфортом добросовестных клиентов.")
    
    vals = [
        ("Предотвращение потерь", "Блокировка подмены реквизитов и дроп-ферм ДО момента списания средств со счета.", PRIMARY_GREEN, "FINANCIAL IMPACT"),
        ("Прозрачность для SOC", "Каждое решение снабжено расшифровкой факторов, отпечатком TLS, SHA-хэшами и статусом устройства.", BLUE, "SOC EXPLAINABILITY"),
        ("Минимум трения (Zero Friction)", "95%+ легитимных клиентов проходят по сценарию ALLOW за < 10 мс без подтверждений.", EMERALD, "CUSTOMER RETENTION"),
        ("Омни-канальное покрытие", "Единый движок верификации для QR-эквайринга, P2P-переводов, платежей по IBAN и Open API.", AMBER, "UNIFIED SHIELD")
    ]
    for i, (v_title, v_desc, v_col, v_tag) in enumerate(vals):
        x = Inches(0.8 + (i % 2) * 5.95)
        y = Inches(1.8 + (i // 2) * 2.45)
        card = add_card(s8, x, y, Inches(5.75), Inches(2.25), accent_color=v_col)
        tf_c = card.text_frame
        
        p_t = tf_c.paragraphs[0]
        p_t.text = v_tag
        p_t.font.name = FONT_MONO
        p_t.font.size = Pt(10)
        p_t.font.bold = True
        p_t.font.color.rgb = v_col
        
        p_h = tf_c.add_paragraph()
        p_h.text = v_title
        p_h.font.name = FONT_HEAD
        p_h.font.size = Pt(16)
        p_h.font.bold = True
        p_h.font.color.rgb = TEXT_WHITE
        p_h.space_after = Pt(4)
        
        p_d = tf_c.add_paragraph()
        p_d.text = v_desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = TEXT_MUTED

    add_footer(s8)

    # ---------------- SLIDE 9: СЦЕНАРИЙ ЖИВОГО ДЕМО ----------------
    s9 = prs.slides.add_slide(layout)
    set_slide_bg(s9, prs)
    add_header(s9, 9, "Сценарий живой демонстрации (60–90 секунд)", 
               "Наглядная демонстрация реакции движка на 4 ключевых паттерна поведения.")
    
    steps = [
        ("00 – 15 сек", "Штатный платёж (QR)", "ALLOW", PRIMARY_GREEN, "Клиент сканирует QR в магазине, мобильный оператор Kcell/Tele2. Задержка 7 мс, хэши совпадают, чистый запуск."),
        ("15 – 35 сек", "Подмена реквизитов", "AUTO-BLOCK", RED, "Малварь пытается изменить номер счета в момент отправки. Расхождение SHA-256 → Мгновенный статус MISMATCH."),
        ("35 – 60 сек", "Эмулятор + Дроп-ферма", "AUTO-BLOCK", RED, "Запуск через эмулятор, серия переводов по 40 000 ₸ на 12 счетов. Детект Frida + Velocity Engine поднимает Threat Score до 98."),
        ("60 – 90 сек", "Прокси + Смена JA3", "STEP-UP", AMBER, "Сессия открыта с дата-центрового IP с ротацией TLS-отпечатка. Система запрашивает биометрию/FIDO2.")
    ]
    for i, (t_range, s_title, verdict, col, s_desc) in enumerate(steps):
        x = Inches(0.8 + i * 2.98)
        card = add_card(s9, x, Inches(1.8), Inches(2.82), Inches(4.8), accent_color=col)
        tf_s = card.text_frame
        
        p_time = tf_s.paragraphs[0]
        p_time.text = t_range
        p_time.font.name = FONT_MONO
        p_time.font.size = Pt(11)
        p_time.font.bold = True
        p_time.font.color.rgb = BLUE
        p_time.space_after = Pt(4)
        
        p_st = tf_s.add_paragraph()
        p_st.text = s_title
        p_st.font.name = FONT_HEAD
        p_st.font.size = Pt(15)
        p_st.font.bold = True
        p_st.font.color.rgb = TEXT_WHITE
        p_st.space_after = Pt(6)
        
        p_v = tf_s.add_paragraph()
        p_v.text = f"Вердикт: {verdict}"
        p_v.font.name = FONT_MONO
        p_v.font.size = Pt(11)
        p_v.font.bold = True
        p_v.font.color.rgb = col
        p_v.space_after = Pt(8)
        
        p_sd = tf_s.add_paragraph()
        p_sd.text = s_desc
        p_sd.font.name = FONT_BODY
        p_sd.font.size = Pt(11)
        p_sd.font.color.rgb = TEXT_MUTED

    add_footer(s9)

    # ---------------- SLIDE 10: ИТОГ И СЛЕДУЮЩИЙ ШАГ ----------------
    s10 = prs.slides.add_slide(layout)
    set_slide_bg(s10, prs)
    add_header(s10, 10, "BCC CyberGuard готов к пилотному внедрению", 
               "Минимальное время интеграции и мгновенный эффект для защиты платёжного периметра.")
    
    # Left Card
    c_sum = add_card(s10, Inches(0.8), Inches(1.8), Inches(6.5), Inches(4.8), accent_color=PRIMARY_GREEN)
    tf_sum = c_sum.text_frame
    p_st = tf_sum.paragraphs[0]
    p_st.text = "РЕЗЮМЕ ПРОЕКТА"
    p_st.font.name = FONT_MONO
    p_st.font.size = Pt(11)
    p_st.font.bold = True
    p_st.font.color.rgb = PRIMARY_GREEN
    
    sum_items = [
        ("Real-Time защита нового поколения", "Оценка риска и решение за < 20 мс прямо в потоке транзакции."),
        ("Криптографическая сверка Payload", "Устранение атак подмены реквизитов на уровне SHA-256 экрана и ядра."),
        ("4 оси глубокого анализа", "Payload + Device Attestation + Velocity + TLS/Network Fingerprinting."),
        ("Полная объяснимость для аналитика", "Каждое решение прозрачно раскладывается на факторы и сигнатуры.")
    ]
    for h, d in sum_items:
        ph = tf_sum.add_paragraph()
        ph.text = f"✔ {h}"
        ph.font.name = FONT_HEAD
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = TEXT_WHITE
        ph.space_before = Pt(8)
        
        pd = tf_sum.add_paragraph()
        pd.text = d
        pd.font.name = FONT_BODY
        pd.font.size = Pt(11)
        pd.font.color.rgb = TEXT_MUTED

    # Right Card: Pilot proposal
    c_p = add_card(s10, Inches(7.6), Inches(1.8), Inches(4.933), Inches(4.8), accent_color=BLUE)
    tf_p = c_p.text_frame
    p_pt = tf_p.paragraphs[0]
    p_pt.text = "ПРЕДЛОЖЕНИЕ ПО ПИЛОТУ"
    p_pt.font.name = FONT_MONO
    p_pt.font.size = Pt(11)
    p_pt.font.bold = True
    p_pt.font.color.rgb = BLUE
    
    steps_pilot = [
        "1. Пилот на ограниченном контуре (P2P или QR-платежи)",
        "2. Тестирование в shadow-режиме (без влияния на клиентов)",
        "3. Интеграция с существующим SIEM / SOC Банка ЦентрКредит",
        "4. Оценка предотвращенного ущерба за 30 дней пилота"
    ]
    for sp in steps_pilot:
        p_sp = tf_p.add_paragraph()
        p_sp.text = sp
        p_sp.font.name = FONT_BODY
        p_sp.font.size = Pt(12)
        p_sp.font.color.rgb = TEXT_WHITE
        p_sp.space_before = Pt(10)
        
    p_qa = tf_p.add_paragraph()
    p_qa.text = "Готовы к обсуждению архитектуры и запуску пилота!"
    p_qa.font.name = FONT_HEAD
    p_qa.font.size = Pt(14)
    p_qa.font.bold = True
    p_qa.font.color.rgb = PRIMARY_GREEN
    p_qa.space_before = Pt(18)

    add_footer(s10)

    out_path = r"c:\Users\user\Downloads\bcc-hackathon\bcc-hackathon\presentations\BCC_CyberGuard_PitchDeck.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}")

build_cyberguard()
