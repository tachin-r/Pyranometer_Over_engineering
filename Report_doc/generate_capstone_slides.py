import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    script_dir = os.path.dirname(__file__)
    img_dir = os.path.join(script_dir, "image")

    FONT_THAI = "TH SarabunPSK"
    FONT_HEAD = "Segoe UI"

    # Clean Minimalist Palette (matching the user's sample PDF)
    C_WHITE = RGBColor(255, 255, 255)       # #FFFFFF Pure clean white
    C_DARK_BG = RGBColor(17, 24, 39)        # #111827 Dark navy for closing slide
    C_BLUE_TAG = RGBColor(2, 132, 199)      # #0284C7 Clean cyan-blue tag/highlight
    C_ORANGE_TAG = RGBColor(217, 119, 6)    # #D97706 Warm amber-orange tag
    C_TEXT_BLACK = RGBColor(15, 23, 42)     # #0F172A Bold dark slate/black
    C_TEXT_BODY = RGBColor(51, 65, 85)      # #334155 Clean readable gray-slate
    C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B Muted caption gray
    C_CARD_GRAY = RGBColor(248, 250, 252)   # #F8FAFC Soft minimalist card fill
    C_BORDER_GRAY = RGBColor(226, 232, 240) # #E2E8F0 Subtle 1pt card border
    C_NAVY_HEADER = RGBColor(15, 23, 42)    # #0F172A Table dark header
    C_ROW_ALT = RGBColor(241, 245, 249)     # #F1F5F9 Table alternating row
    C_RED_BG = RGBColor(254, 242, 242)      # #FEF2F2 Soft red alert card
    C_RED_BORDER = RGBColor(254, 202, 202)  # #FECACA
    C_RED_TEXT = RGBColor(153, 27, 27)      # #991B1B
    C_GREEN_BG = RGBColor(240, 253, 244)    # #F0FDF4 Soft green solution card
    C_GREEN_BORDER = RGBColor(187, 247, 208)# #BBF7D0
    C_GREEN_TEXT = RGBColor(22, 101, 52)    # #166534
    C_BLUE_BG = RGBColor(240, 249, 255)     # #F0F9FF Soft blue card
    C_AMBER_BG = RGBColor(254, 252, 232)    # #FEFCE8 Soft yellow card

    def set_bg(slide, color=C_WHITE):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_clean_header(slide, tag_text, title_text, subtitle_text="", tag_color=C_BLUE_TAG):
        tb = slide.shapes.add_textbox(Inches(0.85), Inches(0.55), Inches(11.6), Inches(1.35))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_tag = tf.paragraphs[0]
        p_tag.space_after = Pt(2)
        r_tag = p_tag.add_run()
        r_tag.text = tag_text
        r_tag.font.name = FONT_THAI
        r_tag.font.size = Pt(18)
        r_tag.font.bold = True
        r_tag.font.color.rgb = tag_color

        p_title = tf.add_paragraph()
        p_title.space_after = Pt(4)
        r_title = p_title.add_run()
        r_title.text = title_text
        r_title.font.name = FONT_THAI
        r_title.font.size = Pt(34)
        r_title.font.bold = True
        r_title.font.color.rgb = C_TEXT_BLACK

        if subtitle_text:
            p_sub = tf.add_paragraph()
            r_sub = p_sub.add_run()
            r_sub.text = subtitle_text
            r_sub.font.name = FONT_THAI
            r_sub.font.size = Pt(18)
            r_sub.font.color.rgb = C_TEXT_BODY

    def add_card(slide, left, top, width, height, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY, border_pt=1):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        if border_rgb is not None and border_pt > 0:
            card.line.color.rgb = border_rgb
            card.line.width = Pt(border_pt)
        else:
            card.line.fill.background()
        return card

    def add_pic_safe(slide, path, left, top, width=None, height=None):
        if os.path.exists(path):
            kwargs = {}
            if width is not None:
                kwargs["width"] = Inches(width)
            if height is not None:
                kwargs["height"] = Inches(height)
            return slide.shapes.add_picture(path, Inches(left), Inches(top), **kwargs)
        return None

    # =========================================================================
    # SLIDE 1: Clean Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1, C_WHITE)

    tb1 = s1.shapes.add_textbox(Inches(0.85), Inches(0.65), Inches(11.6), Inches(1.6))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "Development of High-Precision Solar Pyranometer"
    r1.font.name = FONT_HEAD
    r1.font.size = Pt(34)
    r1.font.bold = True
    r1.font.color.rgb = C_BLUE_TAG

    p1_sub = tf1.add_paragraph()
    p1_sub.alignment = PP_ALIGN.CENTER
    p1_sub.space_before = Pt(6)
    r1_sub = p1_sub.add_run()
    r1_sub.text = "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (Capstone Project 1)"
    r1_sub.font.name = FONT_THAI
    r1_sub.font.size = Pt(24)
    r1_sub.font.bold = True
    r1_sub.font.color.rgb = C_TEXT_BLACK

    # Center Hero Visuals (Perfboard Prototype + 3D Case + 4-Layer PCB)
    p_hero1 = os.path.join(img_dir, "perfboard_top_view.jpg")
    p_hero2 = os.path.join(img_dir, "cad_perfboard_case_assy.png")
    p_hero3 = os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png")
    add_pic_safe(s1, p_hero1, 1.15, 2.25, height=3.3)
    add_pic_safe(s1, p_hero2, 5.15, 2.25, height=3.3)
    add_pic_safe(s1, p_hero3, 8.95, 2.25, height=3.3)

    tb1_f = s1.shapes.add_textbox(Inches(0.85), Inches(5.85), Inches(11.6), Inches(1.2))
    tf1_f = tb1_f.text_frame
    tf1_f.word_wrap = True
    pf1 = tf1_f.paragraphs[0]
    pf1.alignment = PP_ALIGN.CENTER
    rf1 = pf1.add_run()
    rf1.text = "นายเตชินท์ รัฐวงษา (66010296)  •  นายณัฐพัชร์ รัตนนวงศ์ (66010256)  •  นายพีรวิช เฉลิมเกียรติ (66010591)"
    rf1.font.name = FONT_THAI
    rf1.font.size = Pt(20)
    rf1.font.bold = True
    rf1.font.color.rgb = C_TEXT_BLACK

    pf2 = tf1_f.add_paragraph()
    pf2.alignment = PP_ALIGN.CENTER
    rf2 = pf2.add_run()
    rf2.text = "อาจารย์ที่ปรึกษา: ผศ.ดร.ธนวิชญ์ อนุวงศ์พินิจ (วิศวกรรมไอโอทีและสารสนเทศ) และ ผศ.ดร.ประธาน บูรณศิริ (ฟิสิกส์อุตสาหกรรม)"
    rf2.font.name = FONT_THAI
    rf2.font.size = Pt(17)
    rf2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: ความสำคัญในอุตสาหกรรมพลังงานแสงอาทิตย์
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_bg(s2, C_WHITE)
    add_clean_header(s2, "ความสำคัญ", "ในอุตสาหกรรมพลังงานแสงอาทิตย์")

    c2_l = add_card(s2, 0.85, 2.0, 4.5, 4.6, bg_rgb=C_BLUE_BG, border_rgb=None)
    tf2_l = c2_l.text_frame
    tf2_l.word_wrap = True
    tf2_l.margin_left = tf2_l.margin_right = tf2_l.margin_top = Inches(0.3)
    p = tf2_l.paragraphs[0]
    r = p.add_run()
    r.text = "Global Horizontal Irradiance (GHI)"
    r.font.name = FONT_THAI
    r.font.size = Pt(24)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG

    p_eq = tf2_l.add_paragraph()
    p_eq.space_before = Pt(14)
    r_eq = p_eq.add_run()
    r_eq.text = "PR = E_actual / (G_GHI × A_PV × η_STC)"
    r_eq.font.name = FONT_HEAD
    r_eq.font.size = Pt(16)
    r_eq.font.bold = True
    r_eq.font.color.rgb = C_TEXT_BLACK

    p_desc = tf2_l.add_paragraph()
    p_desc.space_before = Pt(14)
    r_desc = p_desc.add_run()
    r_desc.text = (
        "ไพราโนมิเตอร์ทำหน้าที่เสมือน 'ไม้บรรทัดมาตรฐานทางแสง' "
        "ในการตรวจวัดพลังงานแสงอาทิตย์ตกกระทบจริง (W/m²) "
        "เพื่อประเมินสุขภาพและประสิทธิภาพของโรงไฟฟ้าโซลาร์เซลล์"
    )
    r_desc.font.name = FONT_THAI
    r_desc.font.size = Pt(19)
    r_desc.font.color.rgb = C_TEXT_BODY

    tb2_r = s2.shapes.add_textbox(Inches(5.75), Inches(2.0), Inches(6.8), Inches(4.8))
    tf2_r = tb2_r.text_frame
    tf2_r.word_wrap = True
    bullets_s2 = [
        (
            "คำนวณอัตราส่วนประสิทธิภาพของระบบ (Performance Ratio - PR): ",
            "นำค่าความเข้มของแสงที่วัดได้จาก Pyranometer มาเปรียบเทียบกับกำลังไฟฟ้าจริงที่ผลิตได้จาก Inverter เพื่อดูว่าโรงไฟฟ้าทำงานได้เต็มประสิทธิภาพหรือไม่"
        ),
        (
            "ตรวจหาความผิดปกติของระบบและแผงโซลาร์เซลล์: ",
            "หากแสงแดดมีความเข้มสูง แต่กระแสไฟฟ้าที่ผลิตได้กลับต่ำกว่าปกติ จะช่วยให้ผู้ดูแลทราบทันทีว่ามีปัญหา เช่น แผงโซลาร์สกปรก มีฝุ่นสะสม มีเงาบัง หรือแผงเสื่อมสภาพ"
        ),
        (
            "ใช้เป็นเกณฑ์ตรวจสอบสัญญาและประกันการผลิตไฟฟ้า: ",
            "ข้อมูลจาก Pyranometer ทำหน้าที่เป็นหลักฐานอ้างอิงทางวิศวกรรมระหว่างเจ้าของโครงการ ผู้รับเหมาติดตั้ง (EPC) หรือการซื้อขายไฟฟ้า"
        )
    ]
    for idx, (b_head, b_body) in enumerate(bullets_s2):
        p = tf2_r.paragraphs[0] if idx == 0 else tf2_r.add_paragraph()
        p.space_after = Pt(16)
        rh = p.add_run()
        rh.text = "• " + b_head
        rh.font.name = FONT_THAI
        rh.font.size = Pt(21)
        rh.font.bold = True
        rh.font.color.rgb = C_TEXT_BLACK
        rb = p.add_run()
        rb.text = b_body
        rb.font.name = FONT_THAI
        rb.font.size = Pt(20)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 3: ตารางมาตรฐาน ISO 9060:2018
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_bg(s3, C_WHITE)
    add_clean_header(s3, "มาตรฐานสากล", "เกณฑ์การจำแนกประเภท Pyranometer ตามมาตรฐาน ISO 9060:2018")

    rows_s3 = [
        ["พารามิเตอร์สมรรถนะ", "Class A (ระดับสูงสุด)", "Class B", "Class C", "เป้าหมายโครงงานนี้"],
        ["พิกัดความเข้มแสงสูงสุด (Max Irradiance)", "2,000 – 4,000 W/m²", "2,000 W/m²", "1,500 W/m²", "ปรับได้ถึง 4,000 W/m²"],
        ["เวลาการตอบสนอง (Response Time, 95%)", "< 0.5 s (Fast Response)", "< 10 s", "< 20 s", "< 0.05 s (Photodiode)"],
        ["การเลื่อนวนจากอุณหภูมิ (Temperature Drift)", "< ±1%", "< ±2%", "< ±4%", "ชดเชยด้วย NTC + สมการ"],
        ["ความไม่เป็นเชิงเส้น (Non-linearity)", "< ±0.2%", "< ±0.5%", "< ±1.0%", "คง Vbias ป้องกันอิ่มตัว"],
        ["ความคลาดเคลื่อนเชิงมุม (Cosine Error)", "< ±10 W/m²", "< ±20 W/m²", "< ±30 W/m²", "ใช้ตัวถังแบน BPW 34"]
    ]
    tab3 = s3.shapes.add_table(len(rows_s3), 5, Inches(0.85), Inches(1.95), Inches(11.63), Inches(4.8)).table
    col_w3 = [3.33, 2.0, 1.9, 1.9, 2.5]
    for j, w in enumerate(col_w3):
        tab3.columns[j].width = Inches(w)

    for i, row in enumerate(rows_s3):
        for j, val in enumerate(row):
            cell = tab3.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = C_NAVY_HEADER
            elif j == 4:
                cell.fill.fore_color.rgb = C_BLUE_BG
            elif i % 2 == 0:
                cell.fill.fore_color.rgb = C_ROW_ALT
            else:
                cell.fill.fore_color.rgb = C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(19 if i == 0 else 18)
            r.font.bold = (i == 0 or j == 4)
            r.font.color.rgb = C_WHITE if i == 0 else (C_BLUE_TAG if j == 4 else C_TEXT_BLACK)

    # =========================================================================
    # SLIDE 4: ข้อจำกัดด้านต้นทุนของไพราโนมิเตอร์มาตรฐาน (Big Numbers)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_bg(s4, C_WHITE)
    add_clean_header(s4, "ที่มาและความสำคัญของปัญหา", "ข้อจำกัดด้านต้นทุนของไพราโนมิเตอร์มาตรฐาน")

    tb4 = s4.shapes.add_textbox(Inches(1.0), Inches(1.95), Inches(11.33), Inches(5.0))
    tf4 = tb4.text_frame
    tf4.word_wrap = True

    p_n1 = tf4.paragraphs[0]
    p_n1.alignment = PP_ALIGN.CENTER
    r_n1 = p_n1.add_run()
    r_n1.text = "150,000 ฿"
    r_n1.font.name = FONT_HEAD
    r_n1.font.size = Pt(72)
    r_n1.font.bold = True
    r_n1.font.color.rgb = C_BLUE_TAG

    p_d1 = tf4.add_paragraph()
    p_d1.alignment = PP_ALIGN.CENTER
    r_d1 = p_d1.add_run()
    r_d1.text = "ราคาต่อเครื่องของไพราโนมิเตอร์แบบเทอร์โมไพล์มาตรฐาน ISO 9060 Class A"
    r_d1.font.name = FONT_THAI
    r_d1.font.size = Pt(24)
    r_d1.font.bold = True
    r_d1.font.color.rgb = C_TEXT_BLACK

    p_s1 = tf4.add_paragraph()
    p_s1.alignment = PP_ALIGN.CENTER
    p_s1.space_after = Pt(28)
    r_s1 = p_s1.add_run()
    r_s1.text = "อุปสรรคสำคัญต่อการวางโครงข่ายเซนเซอร์ตรวจวัดแสงอาทิตย์แบบกระจายตัวหนาแน่น (Distributed Sensor Network)"
    r_s1.font.name = FONT_THAI
    r_s1.font.size = Pt(19)
    r_s1.font.color.rgb = C_TEXT_MUTED

    p_n2 = tf4.add_paragraph()
    p_n2.alignment = PP_ALIGN.CENTER
    r_n2 = p_n2.add_run()
    r_n2.text = "10,000 ฿ ++"
    r_n2.font.name = FONT_HEAD
    r_n2.font.size = Pt(64)
    r_n2.font.bold = True
    r_n2.font.color.rgb = C_BLUE_TAG

    p_d2 = tf4.add_paragraph()
    p_d2.alignment = PP_ALIGN.CENTER
    r_d2 = p_d2.add_run()
    r_d2.text = "ราคาต่อเครื่องของไพราโนมิเตอร์ระดับเริ่มต้น ISO 9060 Class C (ในขณะที่ต้นแบบโครงงานนี้ใช้งบวัสดุไม่เกิน 3,000 บาท)"
    r_d2.font.name = FONT_THAI
    r_d2.font.size = Pt(23)
    r_d2.font.bold = True
    r_d2.font.color.rgb = C_TEXT_BLACK

    # =========================================================================
    # SLIDE 5: เป้าหมายและแนวทางนวัตกรรมของโครงงานนี้ (4 Clean Cards)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_bg(s5, C_WHITE)
    add_clean_header(s5, "เป้าหมายการวิจัย", "เป้าหมายและแนวทางนวัตกรรมของโครงงานนี้")

    cards_s5 = [
        ("1. บูรณาการ 2 ปริญญา", "ผสานองค์ความรู้ด้าน Optical Physics (การสะท้อน, การหักเห, กฎโคไซน์) เข้ากับ IoT & 24-bit Embedded Engineering"),
        ("2. จำลองทัศนศาสตร์ 3 มิติ", "สร้างโมเดล 3D CAD ของ BPX 61 และ BPW 34 เพื่อวิเคราะห์การหักเหแสงผ่านกระจกครอบด้วย Ray-tracing"),
        ("3. วงจร TIA + ADC 24 บิต", "ออกแบบวงจรขยาย OPA380 คู่กับ ADS1220 และสมการชดเชยทางคณิตศาสตร์ 4 ขั้นตอน รองรับพิกัดถึง 4,000 W/m²"),
        ("4. ต้นแบบจริง + 4-Layer PCB", "บัดกรีและทดสอบบอร์ด Perfboard จริง ออกแบบเคส 3D CAD และแผ่นวงจรพิมพ์ 4 เลเยอร์ในงบประมาณไม่เกิน 3,000 บาท")
    ]
    for idx, (ctitle, cbody) in enumerate(cards_s5):
        left = 0.85 + idx * 2.98
        c = add_card(s5, left, 2.2, 2.75, 4.3, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.25)
        p = tf.paragraphs[0]
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = ctitle
        r.font.name = FONT_THAI
        r.font.size = Pt(22)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_BLACK

        pb = tf.add_paragraph()
        rb = pb.add_run()
        rb.text = cbody
        rb.font.name = FONT_THAI
        rb.font.size = Pt(18)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 6: คัดเลือกเซนเซอร์ (Polar Plot & Dimensional Drawing)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_bg(s6, C_WHITE)
    add_clean_header(s6, "ทฤษฎีและคุณลักษณะหัววัด", "การคัดเลือกเซนเซอร์: เปรียบเทียบโครงสร้างตัวถังและมุมรับแสง")

    p_pol_bpx = os.path.join(img_dir, "(Directional Polar Plot) ของ BPX 61 .png")
    p_pol_bpw = os.path.join(img_dir, "(Directional Polar Plot) ของ BPW 34.png")
    p_dim_bpx = os.path.join(img_dir, "dimention(BPX61).png")
    p_dim_bpw = os.path.join(img_dir, "dimention (BPW34).png")

    add_pic_safe(s6, p_pol_bpx, 0.85, 1.95, height=2.45)
    add_pic_safe(s6, p_dim_bpx, 3.85, 1.95, height=2.45)
    add_pic_safe(s6, p_pol_bpw, 7.05, 1.95, height=2.45)
    add_pic_safe(s6, p_dim_bpw, 9.95, 1.95, height=2.45)

    # Table comparing BPX 61 vs BPW 34 below
    rows_s6 = [
        ["พารามิเตอร์ทางเทคนิค", "ams-OSRAM BPX 61", "ams-OSRAM BPW 34 (รุ่นที่เลือกใช้เป็นหลัก)"],
        ["ขนาดชิปและพื้นที่รับแสงกัมมันต์", "2.65 × 2.65 mm (7.02 mm²)", "2.65 × 2.65 mm (7.02 mm² — เนื้อชิปเดียวกันเป๊ะ)"],
        ["ช่วงความยาวคลื่นและสัมประสิทธิ์อุณหภูมิ", "400 – 1,100 nm  •  TC_I = +0.18 %/K", "400 – 1,100 nm  •  TC_I = +0.18 %/K"],
        ["โครงสร้างตัวถังและมุมรับแสง", "กระป๋องโลหะ TO-39 ลึก 1.55–1.75 mm (ขอบบังแสงมุมเอียง)", "ตัวถังแบนราบ DIL ชิปอยู่ชิดผิวหน้า (รับแสงมุมกว้างตามกฎโคไซน์)"]
    ]
    tab6 = s6.shapes.add_table(len(rows_s6), 3, Inches(0.85), Inches(4.6), Inches(11.63), Inches(2.35)).table
    tab6.columns[0].width = Inches(3.23)
    tab6.columns[1].width = Inches(4.2)
    tab6.columns[2].width = Inches(4.2)
    for i, row in enumerate(rows_s6):
        for j, val in enumerate(row):
            cell = tab6.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_ROW_ALT if i % 2 == 0 else C_WHITE)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(18 if i == 0 else 17)
            r.font.bold = (i == 0 or (i == 3 and j in [1, 2]))
            if i == 0:
                r.font.color.rgb = C_WHITE
            elif i == 3 and j == 1:
                r.font.color.rgb = C_RED_TEXT
            elif i == 3 and j == 2:
                r.font.color.rgb = C_BLUE_TAG
            else:
                r.font.color.rgb = C_TEXT_BLACK

    # =========================================================================
    # SLIDE 7: 3D CAD Design สำหรับทั้ง BPX 61 และ BPW 34 (รวมรูปใหม่!)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_bg(s7, C_WHITE)
    add_clean_header(
        s7,
        "การสร้างแบบจำลอง 3 มิติ (3D CAD Modeling)",
        "การออกแบบโมเดล 3D CAD ของเซนเซอร์ BPX 61 (TO-39) และ BPW 34 (Flat DIL)",
        "จำลองโครงสร้างจริงทุกชิ้นส่วน (Active Die 2.65×2.65 mm², ตัวถัง, ขาโลหะ และกระจกหน้าต่าง) เพื่อใช้ใน Ray-tracing"
    )

    p_cad1 = os.path.join(img_dir, "CAD1.png")
    p_cad2 = os.path.join(img_dir, "CAD2.png")
    p_cad_bpw = os.path.join(img_dir, "cad_bpw34_model.png")

    add_pic_safe(s7, p_cad1, 0.85, 2.05, width=3.8)
    add_pic_safe(s7, p_cad2, 4.80, 2.05, width=3.8)
    add_pic_safe(s7, p_cad_bpw, 8.75, 2.05, width=3.7)

    captions_s7 = [
        (0.85, 3.8, "(ก) โมเดล 3D CAD เซนเซอร์ BPX 61 ในตัวถังโลหะ TO-39"),
        (4.80, 3.8, "(ข) มุมมองตัดขวางแสดงความลึกกระป๋อง 1.55–1.75 mm"),
        (8.75, 3.7, "(ค) โมเดล 3D CAD เซนเซอร์ BPW 34 ตัวถังใสแบนราบ DIL")
    ]
    for left, w, text in captions_s7:
        tb = s7.shapes.add_textbox(Inches(left), Inches(5.85), Inches(w), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.name = FONT_THAI
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 8: ผลการจำลองทางทัศนศาสตร์ BPX 61 vs BPW 34
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_bg(s8, C_WHITE)
    add_clean_header(
        s8,
        "การจำลองทางทัศนศาสตร์ 3 มิติ (Optical Ray-Tracing)",
        "ผลการจำลองการตอบสนองเชิงมุม: BPX 61 (ตัดแสงที่ 44°) เทียบกับ BPW 34",
        tag_color=C_ORANGE_TAG
    )

    p_sim_bpx = os.path.join(script_dir, "sim_bpx61_with_without_glass.png")
    p_sim_bpw = os.path.join(script_dir, "sim_bpw34_with_without_glass.png")
    add_pic_safe(s8, p_sim_bpx, 0.85, 1.95, width=5.6)
    add_pic_safe(s8, p_sim_bpw, 6.75, 1.95, width=5.7)

    c8_bot = add_card(s8, 0.85, 5.45, 11.63, 1.5, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf8 = c8_bot.text_frame
    tf8.word_wrap = True
    tf8.margin_left = tf8.margin_right = Inches(0.25)
    tf8.margin_top = Inches(0.15)
    p8_1 = tf8.paragraphs[0]
    r8_1 = p8_1.add_run()
    r8_1.text = "• การค้นพบจุดตัดแสงที่มุม 44° ของ BPX 61: "
    r8_1.font.name = FONT_THAI
    r8_1.font.size = Pt(19)
    r8_1.font.bold = True
    r8_1.font.color.rgb = C_RED_TEXT
    r8_1b = p8_1.add_run()
    r8_1b.text = "กระจกหน้าต่าง Schott 8250 (n = 1.487) หักเหลำแสงเอียงไปชนผนังกระป๋องโลหะ TO-39 เร็วกว่ากระป๋องเปล่า ทำให้สัญญาณเป็นศูนย์ตั้งแต่ 44°"
    r8_1b.font.name = FONT_THAI
    r8_1b.font.size = Pt(18)
    r8_1b.font.color.rgb = C_TEXT_BODY

    p8_2 = tf8.add_paragraph()
    r8_2 = p8_2.add_run()
    r8_2.text = "• ข้อได้เปรียบของ BPW 34 ใต้กระจกแซฟไฟร์แบน: "
    r8_2.font.name = FONT_THAI
    r8_2.font.size = Pt(19)
    r8_2.font.bold = True
    r8_2.font.color.rgb = C_BLUE_TAG
    r8_2b = p8_2.add_run()
    r8_2b.text = "ไร้ผนังกระป๋องบังแสง รับแสงได้ตลอด 0°–85° ตามกฎโคไซน์ (ที่มุม 44° ยังคงรับแสงได้ 0.710 เทียบกับ BPX 61 ที่มืดสนิท 0.000)"
    r8_2b.font.name = FONT_THAI
    r8_2b.font.size = Pt(18)
    r8_2b.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 9: หลักการวงจรขยาย TIA (OPA380 + BPW 34) และการรักษา Linearity ถึง 4,000 W/m²
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_bg(s9, C_WHITE)
    add_clean_header(
        s9,
        "การออกแบบวงจรแอนะล็อกฟรอนต์เอนด์ (TIA & Photodiode Physics)",
        "หลักการทำงานของ OPA380 TIA และการรักษาความเป็นเชิงเส้นที่แดดจัดถึง 4,000 W/m²"
    )

    p_tia_sch = os.path.join(img_dir, "tia_schematic_kicad.png")
    add_pic_safe(s9, p_tia_sch, 0.85, 1.9, width=5.5)

    c9_eq = add_card(s9, 0.85, 5.15, 5.5, 1.85, bg_rgb=C_DARK_BG, border_rgb=None)
    tf9_eq = c9_eq.text_frame
    tf9_eq.word_wrap = True
    tf9_eq.margin_left = tf9_eq.margin_right = Inches(0.22)
    tf9_eq.margin_top = Inches(0.15)
    p = tf9_eq.paragraphs[0]
    r = p.add_run()
    r.text = "สมการถ่ายโอน TIA & พิกัดความเชิงเส้น (Hamamatsu & Graeme):"
    r.font.name = FONT_THAI
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_ORANGE_TAG

    p2 = tf9_eq.add_paragraph()
    p2.space_before = Pt(4)
    r2 = p2.add_run()
    r2.text = "V_op = V_bias + (I_P × R1)   |   I_sat(linear) = (V_Bi + V_bias) / R_S(photo)"
    r2.font.name = FONT_HEAD
    r2.font.size = Pt(13.5)
    r2.font.bold = True
    r2.font.color.rgb = C_WHITE

    p3 = tf9_eq.add_paragraph()
    p3.space_before = Pt(4)
    r3 = p3.add_run()
    r3.text = "ภายใต้กระจกแซฟไฟร์ใส (T ≈ 85%): กระแสแสง I_P ≈ 2.60 µA ต่อ 1 W/m²"
    r3.font.name = FONT_THAI
    r3.font.size = Pt(17)
    r3.font.color.rgb = RGBColor(186, 230, 253)

    # Right side: Why Vbias = 0.418V + Table of R1 up to 4,000 W/m²
    c9_r1 = add_card(s9, 6.55, 1.9, 5.93, 2.35, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf9_r1 = c9_r1.text_frame
    tf9_r1.word_wrap = True
    tf9_r1.margin_left = tf9_r1.margin_right = Inches(0.2)
    tf9_r1.margin_top = Inches(0.12)
    p = tf9_r1.paragraphs[0]
    r = p.add_run()
    r.text = "ทำไมต้องคงค่า V_bias = 0.418 V (R3=100k, R4=10k) เมื่อไม่ใช้ PTFE?"
    r.font.name = FONT_THAI
    r.font.size = Pt(19)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG

    pts_s9 = [
        "1) ยกพื้นหนี Saturation Floor (60–100 mV) ของ OPA380 ในระบบไฟเลี้ยงเดี่ยว (TI SBOS291G)",
        "2) ต้านทานแรงดันตกคร่อมภายในเนื้อซิลิคอน (I_P × R_S,photo ≈ 3.8 mA × 50 Ω = 0.19 V) ไม่ให้รอยต่อ p-n พลิกเป็นไบแอสตรง (V_OC = 0.365 V) จึง Linear 100% แม้รับแดดจัดผ่านกระจกใส!",
        "3) ไม่ใช้ไฟไบแอสสูง 10 V เพราะกระแสรั่วมืด I_dark พุ่งสูงขึ้นถึง 24 เท่า (2,000 pA) โดยไม่ได้เพิ่มกระแสแสง"
    ]
    for pt_txt in pts_s9:
        p_pt = tf9_r1.add_paragraph()
        p_pt.space_before = Pt(2)
        r_pt = p_pt.add_run()
        r_pt.text = pt_txt
        r_pt.font.name = FONT_THAI
        r_pt.font.size = Pt(16.5)
        r_pt.font.color.rgb = C_TEXT_BLACK

    # Mini table of R1 scaling up to 4,000 W/m²
    rows_r1 = [
        ["ค่าตัวต้านทาน R1 (R_TIA)", "ค่า R4 (V_bias)", "พิกัดวัดแสงสูงสุด (G_max)", "ความละเอียด 24-bit ADC"],
        ["2,200 Ω (ค่าเริ่มต้นทดสอบ)", "10 kΩ (0.418 V)", "~714 W/m² (ในร่ม/ไฟฉาย)", "0.0003 W/m²"],
        ["1,100 Ω (2.2k || 2.2k)", "10 kΩ (0.418 V)", "~1,430 W/m² (แดดจัดปกติ)", "0.0006 W/m²"],
        ["680 Ω (แดดเขตร้อน)", "10 kΩ (0.418 V)", "~2,310 W/m² (ขอบเมฆรวมแสง)", "0.0010 W/m²"],
        ["390 Ω (รุ่นท็อป Class A)", "10k–15k (0.42–0.58V)", "4,000 W/m² (Flagship Spec)", "0.0017 W/m² (~1 mV/W·m⁻²)"]
    ]
    tab9 = s9.shapes.add_table(5, 4, Inches(6.55), Inches(4.4), Inches(5.93), Inches(2.6)).table
    w9 = [1.65, 1.35, 1.65, 1.28]
    for j, w in enumerate(w9):
        tab9.columns[j].width = Inches(w)
    for i, row in enumerate(rows_r1):
        for j, val in enumerate(row):
            cell = tab9.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_GREEN_BG if i == 4 else (C_ROW_ALT if i % 2 == 0 else C_WHITE))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(15.5 if i == 0 else 15)
            r.font.bold = (i == 0 or i == 4)
            r.font.color.rgb = C_WHITE if i == 0 else (C_GREEN_TEXT if i == 4 else C_TEXT_BLACK)

    # =========================================================================
    # SLIDE 10: สถาปัตยกรรมการอ่านสัญญาณและสมการชดเชยทางคณิตศาสตร์ 4 ขั้นตอน
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_bg(s10, C_WHITE)
    add_clean_header(
        s10,
        "สถาปัตยกรรมการอ่านสัญญาณและอัลกอริทึม",
        "ระบบสมการชดเชยทางคณิตศาสตร์ 4 ขั้นตอน (4-Stage Compensation Pipeline)"
    )

    c10_l = add_card(s10, 0.85, 1.95, 5.6, 4.9, bg_rgb=C_DARK_BG, border_rgb=None)
    tf10_l = c10_l.text_frame
    tf10_l.word_wrap = True
    tf10_l.margin_left = tf10_l.margin_right = Inches(0.28)
    tf10_l.margin_top = Inches(0.25)

    p = tf10_l.paragraphs[0]
    r = p.add_run()
    r.text = "เทคนิคผลต่างเทียมและการชดเชยความร้อนครบวงจร"
    r.font.name = FONT_THAI
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = C_ORANGE_TAG

    stages = [
        ("ขั้นที่ 1 (ชดเชยดริฟต์ของ REF5045 3 ppm/°C):", "V_REF(T) = 4.500 × [1 + 3×10⁻⁶ · (T_REF - 25)]"),
        ("ขั้นที่ 2 (หักลบ V_bias แบบผลต่างเทียม 100%):", "I_P(raw) = [ (V_op - V_bias) - ΔV_zero ] / R1"),
        ("ขั้นที่ 3 (ชดเชย Dark Current & TC_I +0.18%/K):", "I_P(25°C) = [I_P(raw) - I_dark(T)] / [1 + 0.0018·(T_NTC - 25)]"),
        ("ขั้นที่ 4 (แปลงเป็นความเข้มรังสีดวงอาทิตย์สุทธิ):", "G_comp [W/m²] = I_P(25°C) / S_eff   (S_eff ≈ 2.60 µA/W·m⁻²)")
    ]
    for st_title, st_eq in stages:
        pt = tf10_l.add_paragraph()
        pt.space_before = Pt(10)
        rt = pt.add_run()
        rt.text = st_title + "\n"
        rt.font.name = FONT_THAI
        rt.font.size = Pt(18)
        rt.font.bold = True
        rt.font.color.rgb = C_WHITE
        req = pt.add_run()
        req.text = st_eq
        req.font.name = FONT_HEAD
        req.font.size = Pt(13)
        req.font.bold = True
        req.font.color.rgb = RGBColor(56, 189, 248)

    # Right side: 3 Channel Mapping Cards (matching Slide 15 of user's PDF)
    ch_cards = [
        (
            "AIN0 & AIN1 (0x81 / 0x91) : สัญญาณแสง V_op และแรงดันฐาน V_bias",
            "บน Perfboard: AIN0 = V_op, AIN1 = V_bias (บน PCB สลับเป็น AIN0 = V_bias, AIN1 = V_op) หักลบออฟเซ็ตและสัญญาณรบกวนร่วมในซอฟต์แวร์ 100%"
        ),
        (
            "AIN2 (0xA1) : อุณหภูมิผิวเซนเซอร์ NTC Thermistor 10 kΩ (±1%)",
            "ดัดแนบชิดใต้ฐานโลหะของตัวเซนเซอร์แสง คำนวณอุณหภูมิผิวหัววัด T_diode เพื่อชดเชยทั้งกระแสรั่วมืด I_dark(T) และสัมประสิทธิ์อุณหภูมิ (+0.18 %/K)"
        ),
        (
            "AIN3 (0xB1) : อุณหภูมิชิปอ้างอิง REF5045 (ขา 3 TEMP)",
            "อ่านแรงดันขา TEMP (509 mV @ 25°C, 2.64 mV/°C) ติดตามเสถียรภาพความร้อนของแหล่งจ่ายอ้างอิง 4.500 V ร่วมกับเซนเซอร์ภายใน ADS1220"
        )
    ]
    for idx, (ch_t, ch_b) in enumerate(ch_cards):
        top = 1.95 + idx * 1.68
        cc = add_card(s10, 6.75, top, 5.73, 1.5, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tfc = cc.text_frame
        tfc.word_wrap = True
        tfc.margin_left = tfc.margin_right = Inches(0.22)
        tfc.margin_top = Inches(0.16)
        p = tfc.paragraphs[0]
        r = p.add_run()
        r.text = ch_t
        r.font.name = FONT_THAI
        r.font.size = Pt(19)
        r.font.bold = True
        r.font.color.rgb = C_BLUE_TAG if idx == 0 else C_TEXT_BLACK

        pb = tfc.add_paragraph()
        pb.space_before = Pt(4)
        rb = pb.add_run()
        rb.text = ch_b
        rb.font.name = FONT_THAI
        rb.font.size = Pt(17.5)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 11: การวิเคราะห์และแก้ปัญหาทางวิศวกรรมฮาร์ดแวร์จริง
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_bg(s11, C_WHITE)
    add_clean_header(
        s11,
        "การวิเคราะห์และแก้ปัญหาทางวิศวกรรม",
        "การดีบักภาคจ่ายไฟ LDO TPS7A4901 และเสถียรภาพแรงดันอ้างอิง REF5045 (4.500 V)",
        tag_color=C_ORANGE_TAG
    )

    # Red Card (Problems discovered)
    c11_red = add_card(s11, 0.85, 2.0, 5.6, 4.8, bg_rgb=C_RED_BG, border_rgb=C_RED_BORDER)
    tfr = c11_red.text_frame
    tfr.word_wrap = True
    tfr.margin_left = tfr.margin_right = Inches(0.28)
    tfr.margin_top = Inches(0.25)
    p = tfr.paragraphs[0]
    r = p.add_run()
    r.text = "ปัญหาทางวิศวกรรมที่ตรวจพบระหว่างสร้างต้นแบบ"
    r.font.name = FONT_THAI
    r.font.size = Pt(23)
    r.font.bold = True
    r.font.color.rgb = C_RED_TEXT

    probs = [
        ("1. ขา 9 (Thermal Pad) ของเรกูเลเตอร์ TPS7A4901 ลอย:", "เมื่อบัดกรีลงแผ่นแปลง MSOP-8 ทั่วไป ไฟออกตกเหลือ < 0.8 V เพราะไอซีตระกูลนี้บังคับให้แผ่นใต้ท้องชิป (ขา 9 EP) เป็นทางเดินกราวด์หลัก"),
        ("2. กระแสรั่วไหลขา 5 (Trim/NR) ของ REF5045:", "โหนดความต้านทานสูง (High-Z) เกิดกระแสรั่วบนเบรดบอร์ดทำให้ไฟอ้างอิงตกเหลือ 4.382 V"),
        ("3. ไฟเลี้ยง USB ดรอป & ค่า ESR ตัวเก็บประจุ:", "สาย USB แรงดันตกต่ำกว่า 4.70 V ทำให้ REF5045 หลุดสภาวะควบคุม (Dropout) และต้องการ C โหลดชนิด Low-ESR")
    ]
    for ph, pb in probs:
        pp = tfr.add_paragraph()
        pp.space_before = Pt(12)
        rh = pp.add_run()
        rh.text = ph + " "
        rh.font.name = FONT_THAI
        rh.font.size = Pt(19.5)
        rh.font.bold = True
        rh.font.color.rgb = C_RED_TEXT
        rb = pp.add_run()
        rb.text = pb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(18.5)
        rb.font.color.rgb = C_TEXT_BLACK

    # Green Card (Solutions implemented on Perfboard)
    c11_grn = add_card(s11, 6.85, 2.0, 5.63, 4.8, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tfg = c11_grn.text_frame
    tfg.word_wrap = True
    tfg.margin_left = tfg.margin_right = Inches(0.28)
    tfg.margin_top = Inches(0.25)
    p = tfg.paragraphs[0]
    r = p.add_run()
    r.text = "การแก้ไขสำเร็จจริงบนบอร์ด Perfboard และ PCB"
    r.font.name = FONT_THAI
    r.font.size = Pt(23)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_TEXT

    sols = [
        ("1. เชื่อมกราวด์ขา 9 (Exposed Pad) ของ TPS7A4901:", "ต่อสายกราวด์เข้ากับแผ่น Thermal Pad ใต้ท้องชิปโดยตรง ทำให้ได้ไฟเลี้ยงแอนะล็อกสะอาด +5.20 V นิ่งสนิททันที"),
        ("2. ปลดขา 5 (DNP Trim) และใส่ C เซรามิก Low-ESR 10 µF:", "ขจัดกระแสรั่วไหลและรักษาเสถียรภาพลูปควบคุมภายใน ทำให้ REF5045 จ่ายแรงดันอ้างอิงได้ตรงพิกัด 4.500 V"),
        ("3. สถาปัตยกรรมจ่ายไฟ 2 ชั้น (LM2596 -> TPS7A4901):", "สร้างส่วนต่างแรงดัน Headroom 700 mV เหนือ 4.500 V ไร้ผลกระทบจากการดึงกระแสส่ง Wi-Fi ของ ESP32-C6")
    ]
    for sh, sb in sols:
        ps = tfg.add_paragraph()
        ps.space_before = Pt(12)
        rh = ps.add_run()
        rh.text = sh + " "
        rh.font.name = FONT_THAI
        rh.font.size = Pt(19.5)
        rh.font.bold = True
        rh.font.color.rgb = C_GREEN_TEXT
        rb = ps.add_run()
        rb.text = sb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(18.5)
        rb.font.color.rgb = C_TEXT_BLACK

    # =========================================================================
    # SLIDE 12: บอร์ดต้นแบบ Perfboard จริง (3 มุมมองจริงที่เพิ่งอัปโหลด!)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_bg(s12, C_WHITE)
    add_clean_header(
        s12,
        "การสร้างต้นแบบฮาร์ดแวร์จริง (Perfboard Prototype 70×90 mm)",
        "ผลงานการบัดกรีประกอบบอร์ดต้นแบบ Perfboard แบบแยกด้านรับแสงและด้านวงจร",
        "ด้านบนติดตั้งเฉพาะหัววัดแสงและ NTC เพื่อให้เปิดมุมรับแสง 180° ไร้เงาบัง ส่วนด้านล่างรวมโมดูลวงจรครบทั้ง 6 ภาค"
    )

    pf_top = os.path.join(img_dir, "perfboard_top_view.jpg")
    pf_side = os.path.join(img_dir, "perfboard_side_ntc_view.jpg")
    pf_bot = os.path.join(img_dir, "perfboard_bottom_modules.jpg")

    add_pic_safe(s12, pf_top, 0.85, 2.0, width=3.75)
    add_pic_safe(s12, pf_side, 4.78, 2.0, width=3.75)
    add_pic_safe(s12, pf_bot, 8.71, 2.0, width=3.75)

    caps_s12 = [
        (0.85, "1. ด้านบนบอร์ด (Top Optical Side)", "ติดตั้งเฉพาะเซนเซอร์แสงกึ่งกลางบอร์ดและ NTC ไร้ชิ้นส่วนสูงบังแสงที่มุมเอียง"),
        (4.78, "2. มุมมองด้านข้างและการแนบ NTC", "ดัดหัววัด NTC 10 kΩ (สีเขียว) แนบสนิทกับฐานโลหะใต้ตัวถังเซนเซอร์เพื่อวัดอุณหภูมิจริง"),
        (8.71, "3. ด้านล่างบอร์ด (Bottom Electronics)", "ประกอบด้วย LM2596, TPS7A4901 (+5.2V), REF5045 (4.5V), OPA380, ADS1220 และ ESP32-C6")
    ]
    for left, ct, cb in caps_s12:
        card = add_card(s12, left, 5.05, 3.75, 1.85, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.14)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = ct
        r.font.name = FONT_THAI
        r.font.size = Pt(19)
        r.font.bold = True
        r.font.color.rgb = C_BLUE_TAG
        pb = tf.add_paragraph()
        pb.space_before = Pt(4)
        rb = pb.add_run()
        rb.text = cb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(17)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 13: ผลการอ่านสัญญาณจริง 24 บิตบนบอร์ดต้นแบบ Perfboard
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_bg(s13, C_WHITE)
    add_clean_header(
        s13,
        "ผลการทดลองบนแผงวงจรต้นแบบ (Perfboard Validation)",
        "ผลการอ่านสัญญาณจริง 24 บิตครบทั้ง 4 ช่องสัญญาณใน 4 สภาวะแสง"
    )

    rows_s13 = [
        ["สภาวะแสงทดสอบจริง", "AIN0: V_op", "AIN1: V_bias", "AIN2: NTC (Temp)", "AIN3: Temp REF", "แรงดันสุทธิ ΔV", "กระแส I_P / ความเข้มแสง"],
        ["1. ปิดทึบมืดสนิท (Dark)", "0.4185 V", "0.4180 V", "2.0942 V (28.48 °C)", "0.5940 V (32.20 °C)", "+0.50 mV", "0.227 µA (~0.09 W/m²)"],
        ["2. แสงสว่างปกติในห้อง", "0.4405 V", "0.4180 V", "2.0940 V (28.49 °C)", "0.5942 V (32.27 °C)", "+22.50 mV", "10.23 µA (~3.93 W/m²)"],
        ["3. โคมไฟส่องระยะใกล้", "1.5420 V", "0.4180 V", "2.0885 V (28.62 °C)", "0.5948 V (32.50 °C)", "+1,124.0 mV", "510.9 µA (~196.5 W/m²)"],
        ["4. ไฟฉายแรงสูงจ่อชิด (R1=2.2k)", "4.4998 V (เต็ม)", "0.4181 V", "2.0810 V (28.80 °C)", "0.5951 V (32.61 °C)", "+4,081.7 mV", "1,855 µA (อิ่มตัวที่ ~714 W/m²)"]
    ]
    tab13 = s13.shapes.add_table(5, 7, Inches(0.85), Inches(1.95), Inches(11.63), Inches(3.9)).table
    w13 = [2.25, 1.35, 1.3, 1.75, 1.75, 1.45, 1.78]
    for j, w in enumerate(w13):
        tab13.columns[j].width = Inches(w)

    for i, row in enumerate(rows_s13):
        for j, val in enumerate(row):
            cell = tab13.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_BLUE_BG if i == 4 else (C_ROW_ALT if i % 2 == 0 else C_WHITE))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(17.5 if i == 0 else 17)
            r.font.bold = (i == 0 or j == 6)
            r.font.color.rgb = C_WHITE if i == 0 else (C_BLUE_TAG if j == 6 else C_TEXT_BLACK)

    tb13_note = s13.shapes.add_textbox(Inches(0.85), Inches(6.05), Inches(11.63), Inches(0.9))
    p_n13 = tb13_note.text_frame.paragraphs[0]
    r_n13 = p_n13.add_run()
    r_n13.text = (
        "สรุปผลการทดลองจริง: แรงดันฐาน V_bias นิ่งสนิทที่ 0.4180 V ตลอดทุกสภาวะแสง ช่วยยก V_op พ้นพื้นอิ่มตัว 100 mV ของ OPA380 อย่างสมบูรณ์ "
        "และพิสูจน์ว่าเมื่อรับแสงความเข้มสูงผ่านกระจกใส การเปลี่ยน R1 ลงเป็น 680 Ω หรือ 390 Ω จะขยายพิกัดการวัดได้ถึง 2,310 – 4,000 W/m² ทันที"
    )
    r_n13.font.name = FONT_THAI
    r_n13.font.size = Pt(18.5)
    r_n13.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 14: ปัจจุบันทำอะไรไปแล้วบ้าง — เคส 3D CAD สำหรับ Perfboard + 4-Layer PCB
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_bg(s14, C_WHITE)
    add_clean_header(
        s14,
        "ความก้าวหน้าด้านการออกแบบทางกลและแผ่นวงจรพิมพ์ (Mechanical & PCB Engineering)",
        "การออกแบบเคส 3D CAD สำหรับ Perfboard และแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน",
        "พร้อมสำหรับการประกอบทดสอบกลางแจ้งและการผลิตแผ่นวงจรพิมพ์มาตรฐานโรงงาน (DRC Violations = 0)"
    )

    p_case = os.path.join(img_dir, "cad_perfboard_case_assy.png")
    p_ring = os.path.join(img_dir, "cad_sapphire_retaining_ring.png")
    p_pcb1 = os.path.join(img_dir, "pcb_3d_render_iso.png")
    p_pcb2 = os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png")

    add_pic_safe(s14, p_case, 0.85, 2.0, width=2.75)
    add_pic_safe(s14, p_ring, 3.78, 2.0, width=2.75)
    add_pic_safe(s14, p_pcb1, 6.71, 2.0, width=2.75)
    add_pic_safe(s14, p_pcb2, 9.64, 2.0, width=2.75)

    cards_s14 = [
        (0.85, "1. กล่องเคสสำหรับ Perfboard", "ออกแบบขนาดพอดีบอร์ด 70×90 mm พร้อมช่องหน้าต่างรับแสงทรงกลมด้านบนและพอร์ต Cable Gland กันน้ำด้านข้าง"),
        (3.78, "2. แหวนล็อกกระจก Sapphire", "ฝาแหวนกดล็อกกระจกแซฟไฟร์ (Retaining Bezel Ring) แบบมีร่องขันล็อก 6 จุด ถอดสลับทดสอบกระจกได้สะดวก"),
        (6.71, "3. 4-Layer PCB (รุ่น BPX 61)", "แผ่นวงจรพิมพ์ 4 เลเยอร์ รุ่น V1.1 L1_Analog สำหรับตัวถังโลหะ TO-39 พร้อมวงแหวน Guard Ring รอบขา -IN"),
        (9.64, "4. 4-Layer PCB (รุ่น BPW 34)", "แผ่นวงจรพิมพ์ 4 เลเยอร์ รุ่น V1.1 L1_Analog_BPW34 สำหรับตัวถังแบนราบ DIL เพื่อทดสอบเปรียบเทียบแบบควบคุมตัวแปร")
    ]
    for left, ct, cb in cards_s14:
        c = add_card(s14, left, 4.95, 2.75, 1.95, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.15)
        tf.margin_top = Inches(0.12)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = ct
        r.font.name = FONT_THAI
        r.font.size = Pt(18)
        r.font.bold = True
        r.font.color.rgb = C_BLUE_TAG
        pb = tf.add_paragraph()
        pb.space_before = Pt(3)
        rb = pb.add_run()
        rb.text = cb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(16)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 15: การออกแบบระเบียบวิธีการทดลองทางทัศนศาสตร์จริง (LabQuest 2 & Goniometer)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_bg(s15, C_WHITE)
    add_clean_header(
        s15,
        "ความร่วมมือกับภาควิชาฟิสิกส์ คณะวิทยาศาสตร์",
        "การออกแบบระเบียบวิธีการทดลองทางทัศนศาสตร์จริง"
    )

    p_lq = os.path.join(img_dir, "labquest2_vernier_spectrometer.png")
    p_gon = os.path.join(img_dir, "goniometer_blender_perspective.png")
    add_pic_safe(s15, p_lq, 0.85, 1.95, height=2.9)
    add_pic_safe(s15, p_gon, 6.25, 1.95, height=2.9)

    lab_cards = [
        ("1. วัดการส่องผ่านสเปกตรัม", "ใช้เครื่อง Vernier LabQuest 2 ช่วง 350–900 nm ทดสอบค่าร้อยละการส่องผ่าน (%T ≈ 85%) ของกระจกแซฟไฟร์จริง"),
        ("2. ทดสอบมุมบนแท่นหมุน", "ยึดหัววัดร่วมแกนหมุน Optical Rotation Stage ปรับมุมรับแสงตั้งแต่ 0° ถึง 85° เทียบระหว่างมีและไม่มีกระจก"),
        ("3. ลำแสงขนานเสมือน", "ตั้งหลอดทังสเตน-ฮาโลเจนไฟตรงห่าง 1.5–2.0 เมตร (มุมถ่างแสง < 0.08°) พร้อมท่อบังแสงสะท้อน Baffle Tube")
    ]
    for idx, (lt, lb) in enumerate(lab_cards):
        left = 0.85 + idx * 3.98
        c = add_card(s15, left, 5.05, 3.68, 1.85, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.14)
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = lt
        r.font.name = FONT_THAI
        r.font.size = Pt(20)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_BLACK
        pb = tf.add_paragraph()
        pb.space_before = Pt(4)
        rb = pb.add_run()
        rb.text = lb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(17.5)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 16: สถานะปัจจุบันและแผนการดำเนินงานสู่โครงงาน 2 (Updated Roadmap!)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_bg(s16, C_WHITE)
    add_clean_header(
        s16,
        "สรุปความก้าวหน้าและแผนงาน",
        "สถานะปัจจุบันและแผนการดำเนินงานสู่โครงงาน 2"
    )

    roadmap_cards = [
        (
            "สำเร็จแล้ว 100%",
            "จำลองแสง 3 มิติ & พิสูจน์ทฤษฎี",
            "สร้าง 3D CAD ทั้ง BPX 61 และ BPW 34 พิสูจน์จุดตัดแสง 44° และกำหนดสมการชดเชยทางคณิตศาสตร์ 4 ขั้นตอน",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "สำเร็จแล้ว 100%",
            "บัดกรี Perfboard & ออกแบบเคส",
            "บัดกรีวงจรลงบอร์ด Perfboard (70×90 mm) แก้ปัญหาขา 9 ของ TPS7A4901 อ่านค่า 24 บิตสำเร็จ และออกแบบเคส 3D CAD + 4-Layer PCB",
            C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "โครงงาน 2 (ระยะที่ 1)",
            "ทดสอบห้องปฏิบัติการฟิสิกส์",
            "นำหัววัดจริงเข้าวัดสเปกตรัมด้วย Vernier LabQuest 2 และทดสอบการตอบสนองเชิงมุม 0°–85° บนแท่นหมุน",
            C_AMBER_BG, C_ORANGE_TAG
        ),
        (
            "โครงงาน 2 (ระยะที่ 2)",
            "สอบเทียบกับเครื่องมาตรฐาน",
            "ประกอบลงกล่องเคสกันน้ำพร้อมกระจกแซฟไฟร์ และนำไปตั้งวัดแสงอาทิตย์จริงเทียบกับไพราโนมิเตอร์มาตรฐาน",
            C_CARD_GRAY, C_TEXT_BLACK
        )
    ]
    for idx, (tag, rtitle, rbody, bg_c, tc) in enumerate(roadmap_cards):
        left = 0.85 + idx * 2.98
        c = add_card(s16, left, 2.2, 2.75, 4.3, bg_rgb=bg_c, border_rgb=C_BORDER_GRAY)
        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = Inches(0.24)
        p_t = tf.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = tag
        r_t.font.name = FONT_THAI
        r_t.font.size = Pt(18)
        r_t.font.bold = True
        r_t.font.color.rgb = tc

        p_h = tf.add_paragraph()
        p_h.space_before = Pt(6)
        p_h.space_after = Pt(8)
        r_h = p_h.add_run()
        r_h.text = rtitle
        r_h.font.name = FONT_THAI
        r_h.font.size = Pt(22)
        r_h.font.bold = True
        r_h.font.color.rgb = C_TEXT_BLACK

        p_b = tf.add_paragraph()
        r_b = p_b.add_run()
        r_b.text = rbody
        r_b.font.name = FONT_THAI
        r_b.font.size = Pt(18)
        r_b.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # SLIDE 17: Closing Slide (Dark Navy Theme matching Slide 24 of user's PDF)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_bg(s17, C_DARK_BG)

    bar_end = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.36), Inches(2.6), Inches(0.6), Inches(0.06))
    bar_end.fill.solid()
    bar_end.fill.fore_color.rgb = RGBColor(249, 115, 22) # orange accent line
    bar_end.line.fill.background()

    tb_end = s17.shapes.add_textbox(Inches(1.0), Inches(2.85), Inches(11.33), Inches(3.0))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True

    pe1 = tf_end.paragraphs[0]
    pe1.alignment = PP_ALIGN.CENTER
    re1 = pe1.add_run()
    re1.text = "ขอบคุณครับ • รับฟังข้อเสนอแนะ"
    re1.font.name = FONT_THAI
    re1.font.size = Pt(46)
    re1.font.bold = True
    re1.font.color.rgb = C_WHITE

    pe2 = tf_end.add_paragraph()
    pe2.alignment = PP_ALIGN.CENTER
    pe2.space_before = Pt(12)
    re2 = pe2.add_run()
    re2.text = "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (Capstone Project 1)"
    re2.font.name = FONT_THAI
    re2.font.size = Pt(22)
    re2.font.color.rgb = RGBColor(148, 163, 184)

    pe3 = tf_end.add_paragraph()
    pe3.alignment = PP_ALIGN.CENTER
    pe3.space_before = Pt(24)
    re3 = pe3.add_run()
    re3.text = "เตชินท์ รัฐวงษา  •  ณัฐพัชร์ รัตนนวงศ์  •  พีรวิช เฉลิมเกียรติ"
    re3.font.name = FONT_THAI
    re3.font.size = Pt(19)
    re3.font.color.rgb = RGBColor(203, 213, 225)

    candidates = [
        "69_IoT_Pyranometer_Capstone1_Slides.pptx",
        "69_IoT_Pyranometer_Capstone1_Slides_Latest.pptx",
        "69_IoT_Pyranometer_Capstone1_Slides_Clean.pptx"
    ]
    saved_path = None
    for fname in candidates:
        out_path = os.path.join(script_dir, fname)
        try:
            prs.save(out_path)
            print(f"[OK] Clean Presentation saved to: {out_path}")
            saved_path = out_path
            break
        except PermissionError:
            print(f"[NOTE] '{fname}' is open in PowerPoint. Trying next...")
    return saved_path


if __name__ == "__main__":
    create_presentation()
