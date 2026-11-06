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

    # Font definitions
    FONT_HEAD = "Segoe UI"
    FONT_BODY = "Segoe UI"
    FONT_THAI = "TH SarabunPSK"

    # Color Palette
    C_NAVY_DARK = RGBColor(11, 19, 43)     # #0B132B
    C_NAVY_SLATE = RGBColor(15, 23, 42)    # #0F172A
    C_SLATE_CARD = RGBColor(30, 41, 59)    # #1E293B
    C_TEXT_LIGHT = RGBColor(248, 250, 252) # #F8FAFC
    C_TEXT_MUTED = RGBColor(148, 163, 184) # #94A3B8
    C_CYAN_ACCENT = RGBColor(14, 165, 233) # #0EA5E9
    C_CYAN_LIGHT = RGBColor(56, 189, 248)  # #38BDF8
    
    # Light theme colors
    C_BG_LIGHT = RGBColor(248, 250, 252)   # #F8FAFC
    C_TEXT_MAIN = RGBColor(15, 23, 42)     # #0F172A
    C_TEXT_SUB = RGBColor(71, 85, 105)     # #475569
    C_CARD_BG = RGBColor(255, 255, 255)    # #FFFFFF
    C_CARD_BORDER = RGBColor(226, 232, 240)# #E2E8F0
    C_BLUE_BRAND = RGBColor(2, 132, 199)   # #0284C7
    C_GREEN_OK = RGBColor(16, 185, 129)    # #10B981
    C_AMBER_WARN = RGBColor(217, 119, 6)   # #D97706
    C_PURPLE_IOT = RGBColor(124, 58, 237)  # #7C3AED

    TOTAL_SLIDES = 15

    # -------------------------------------------------------------
    # Helper Functions
    # -------------------------------------------------------------
    def set_slide_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background()
        return bg

    def add_header(slide, category_tag, title_text, subtitle_text=""):
        # Category Tag Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(2.6), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = RGBColor(224, 242, 254) # sky 100
        pill.line.color.rgb = C_BLUE_BRAND
        pill.line.width = Pt(1)
        tf_pill = pill.text_frame
        tf_pill.word_wrap = True
        tf_pill.margin_left = tf_pill.margin_right = tf_pill.margin_top = tf_pill.margin_bottom = 0
        p_pill = tf_pill.paragraphs[0]
        p_pill.alignment = PP_ALIGN.CENTER
        r_pill = p_pill.add_run()
        r_pill.text = category_tag.upper()
        r_pill.font.name = FONT_HEAD
        r_pill.font.size = Pt(10)
        r_pill.font.bold = True
        r_pill.font.color.rgb = C_BLUE_BRAND

        # Title & Subtitle text box
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        
        p_title = tf.paragraphs[0]
        p_title.space_after = Pt(2)
        r_title = p_title.add_run()
        r_title.text = title_text
        r_title.font.name = FONT_HEAD
        r_title.font.size = Pt(22)
        r_title.font.bold = True
        r_title.font.color.rgb = C_TEXT_MAIN

        if subtitle_text:
            p_sub = tf.add_paragraph()
            r_sub = p_sub.add_run()
            r_sub.text = subtitle_text
            r_sub.font.name = FONT_BODY
            r_sub.font.size = Pt(12)
            r_sub.font.color.rgb = C_TEXT_SUB

    def add_footer(slide, current_slide):
        # Line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(226, 232, 240)
        line.line.fill.background()

        # Footer Text
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(10.0), Inches(0.3))
        tf = tb.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = "Capstone Project 1 | High-Precision Solar Pyranometer | Dual Degree IoT & Industrial Physics, KMITL"
        r.font.name = FONT_BODY
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(148, 163, 184)

        # Page Number
        tb_num = slide.shapes.add_textbox(Inches(11.0), Inches(7.05), Inches(1.533), Inches(0.3))
        tf_num = tb_num.text_frame
        tf_num.margin_left = tf_num.margin_right = tf_num.margin_top = tf_num.margin_bottom = 0
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        r_num = p_num.add_run()
        r_num.text = f"{current_slide} / {TOTAL_SLIDES}"
        r_num.font.name = FONT_BODY
        r_num.font.size = Pt(9.5)
        r_num.font.bold = True
        r_num.font.color.rgb = C_BLUE_BRAND

    def add_card(slide, left, top, width, height, bg_rgb=C_CARD_BG, border_rgb=C_CARD_BORDER, border_pt=1):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        card.line.color.rgb = border_rgb
        card.line.width = Pt(border_pt)
        return card

    # =========================================================================
    # SLIDE 1: Title Slide (Dark Theme / Executive High Impact)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, C_NAVY_DARK)

    # Accent decorative glow bar
    bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.8), Inches(0.15), Inches(4.0))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_CYAN_ACCENT
    bar.line.fill.background()

    # Title box
    tb_t = slide1.shapes.add_textbox(Inches(1.2), Inches(0.7), Inches(11.3), Inches(3.6))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True

    p_badge = tf_t.paragraphs[0]
    p_badge.space_after = Pt(8)
    r_b = p_badge.add_run()
    r_b.text = "CAPSTONE PROJECT 1 PROPOSAL & PROGRESS REPORT (2026)"
    r_b.font.name = FONT_HEAD
    r_b.font.size = Pt(11)
    r_b.font.bold = True
    r_b.font.color.rgb = C_CYAN_LIGHT

    p_th = tf_t.add_paragraph()
    p_th.space_after = Pt(6)
    r_th = p_th.add_run()
    r_th.text = "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"
    r_th.font.name = FONT_HEAD
    r_th.font.size = Pt(28)
    r_th.font.bold = True
    r_th.font.color.rgb = C_TEXT_LIGHT

    p_en = tf_t.add_paragraph()
    p_en.space_after = Pt(12)
    r_en = p_en.add_run()
    r_en.text = "Development of High-Precision Solar Pyranometer"
    r_en.font.name = FONT_HEAD
    r_en.font.size = Pt(21)
    r_en.font.bold = True
    r_en.font.color.rgb = C_CYAN_ACCENT

    p_sub = tf_t.add_paragraph()
    r_sub = p_sub.add_run()
    r_sub.text = "Dual Degree: B.Eng (IoT System & Information Engineering) & B.Sc. (Industrial Physics)\nFaculty of Engineering & Faculty of Science, King Mongkut's Institute of Technology Ladkrabang"
    r_sub.font.name = FONT_BODY
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = C_TEXT_MUTED

    # Student Cards (3 Columns)
    st_lefts = [1.2, 5.0, 8.8]
    st_names = [
        ("นายเตชินท์ รัฐวงษา", "รหัสนักศึกษา 66010296", "MR. TECHIN RATTAWONGSA"),
        ("นายณัฐพัชร์ รัตนนวงศ์", "รหัสนักศึกษา 66010256", "MR. NATTAPAT RATANANWONG"),
        ("นายพีรวิช เฉลิมเกียรติ", "รหัสนักศึกษา 66010591", "MR. PEERAWICH CHALOEMKIAT")
    ]
    for idx, (th_n, sid, en_n) in enumerate(st_names):
        c_st = add_card(slide1, st_lefts[idx], 4.7, 3.5, 1.2, bg_rgb=C_SLATE_CARD, border_rgb=RGBColor(51, 65, 85))
        tf_sc = c_st.text_frame
        tf_sc.word_wrap = True
        tf_sc.margin_left = tf_sc.margin_right = tf_sc.margin_top = tf_sc.margin_bottom = Inches(0.12)
        p1 = tf_sc.paragraphs[0]
        r1 = p1.add_run()
        r1.text = th_n
        r1.font.name = FONT_HEAD
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_LIGHT
        
        p2 = tf_sc.add_paragraph()
        r2 = p2.add_run()
        r2.text = sid
        r2.font.name = FONT_BODY
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_CYAN_LIGHT

        p3 = tf_sc.add_paragraph()
        r3 = p3.add_run()
        r3.text = en_n
        r3.font.name = FONT_BODY
        r3.font.size = Pt(9)
        r3.font.color.rgb = C_TEXT_MUTED

    # Advisor Bar
    c_adv = add_card(slide1, 1.2, 6.05, 11.1, 0.95, bg_rgb=RGBColor(15, 23, 42), border_rgb=RGBColor(51, 65, 85))
    tf_adv = c_adv.text_frame
    tf_adv.word_wrap = True
    tf_adv.margin_left = tf_adv.margin_right = tf_adv.margin_top = tf_adv.margin_bottom = Inches(0.12)
    p_a1 = tf_adv.paragraphs[0]
    r_a1 = p_a1.add_run()
    r_a1.text = "อาจารย์ที่ปรึกษาโครงงาน (Project Advisors):  "
    r_a1.font.name = FONT_HEAD
    r_a1.font.size = Pt(11)
    r_a1.font.bold = True
    r_a1.font.color.rgb = C_CYAN_LIGHT

    r_a2 = p_a1.add_run()
    r_a2.text = "ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ "
    r_a2.font.name = FONT_HEAD
    r_a2.font.size = Pt(11)
    r_a2.font.bold = True
    r_a2.font.color.rgb = C_TEXT_LIGHT
    r_a2_sub = p_a1.add_run()
    r_a2_sub.text = "(ที่ปรึกษาหลัก - วศ.บ. ระบบไอโอทีและสารสนเทศ)    |    "
    r_a2_sub.font.size = Pt(9.5)
    r_a2_sub.font.color.rgb = C_TEXT_MUTED

    r_a3 = p_a1.add_run()
    r_a3.text = "ผศ.ดร.ประธาน บูรณศิริ "
    r_a3.font.name = FONT_HEAD
    r_a3.font.size = Pt(11)
    r_a3.font.bold = True
    r_a3.font.color.rgb = C_TEXT_LIGHT
    r_a3_sub = p_a1.add_run()
    r_a3_sub.text = "(ที่ปรึกษาร่วม - วท.บ. ฟิสิกส์อุตสาหกรรม, ผู้เชี่ยวชาญทัศนศาสตร์)"
    r_a3_sub.font.size = Pt(9.5)
    r_a3_sub.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 2: Background & Problem Statement
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, C_BG_LIGHT)
    add_header(slide2, "Project Motivation", "ความเป็นมา ความสำคัญ และปัญหาของระบบวัดแสงอาทิตย์", 
               "ทำไมต้องวิจัยและพัฒนาระบบวัดรังสีแสงอาทิตย์ความแม่นยำสูง (Solar Pyranometer) ขึ้นเอง?")
    add_footer(slide2, 2)

    # 3 Cards Layout
    w_card = 3.65
    h_card = 5.0
    c_lefts = [0.8, 4.84, 8.88]

    # Card 1: ความสำคัญ
    c1 = add_card(slide2, c_lefts[0], 1.75, w_card, h_card)
    tf1 = c1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_right = tf1.margin_top = tf1.margin_bottom = Inches(0.2)
    p = tf1.paragraphs[0]
    r = p.add_run()
    r.text = "ความสำคัญในอุตสาหกรรม\nพลังงานแสงอาทิตย์"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    pts1 = [
        ("การประเมิน Performance Ratio (PR):", "จำเป็นต้องใช้ค่าความเข้มแสงอาทิตย์ (Irradiance: W/m²) ที่แม่นยำเพื่อชี้วัดประสิทธิภาพโรงไฟฟ้า PV"),
        ("มาตรฐานสากล ISO 9060:2018:", "แบ่งชั้นคุณภาพเป็น Class A (Spectrally Flat), Class B และ Class C ซึ่งควบคุม Directional & Spectral Error อย่างเข้มงวด"),
        ("ความต้องการในระบบ Smart Grid:", "ต้องมีอัตราการตอบสนองเร็ว (Fast Response Time) เพื่อตรวจจับเมฆบดบังแสงชั่วขณะ (Cloud Transients)")
    ]
    for h, b in pts1:
        p_h = tf1.add_paragraph()
        p_h.space_before = Pt(8)
        r_h = p_h.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(11)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_h.add_run()
        r_b.text = b
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = C_TEXT_SUB

    # Card 2: ข้อจำกัดอุปกรณ์ในตลาด
    c2 = add_card(slide2, c_lefts[1], 1.75, w_card, h_card, bg_rgb=RGBColor(254, 242, 242), border_rgb=RGBColor(254, 202, 202))
    tf2 = c2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = Inches(0.2)
    p = tf2.paragraphs[0]
    r = p.add_run()
    r.text = "ข้อจำกัดของอุปกรณ์\nในท้องตลาดปัจจุบัน"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(220, 38, 38) # red

    pts2 = [
        ("Thermopile Pyranometer (Class A):", "ราคาสูงมาก (100,000 - 350,000+ บาท), ตอบสนองช้า (5-30 วินาที), มีปัญหา Thermal Zero Offset A/B เมื่อท้องฟ้าโปร่ง"),
        ("Silicon Pyranometer ทั่วไป (< 10,000 บ.):", "เกิด Cosine Error รุนแรงที่มุมชัน, สเปกตรัมแคบ และมี Thermal Drift สูงตามอุณหภูมิแวดล้อม (+0.18%/K)"),
        ("ช่องว่างทางเทคโนโลยี (The Gap):", "ยังไม่มีอุปกรณ์ที่รวมความเร็วของโฟโตไดโอด เข้ากับความแม่นยำระดับ Class A ในราคาประหยัด")
    ]
    for h, b in pts2:
        p_h = tf2.add_paragraph()
        p_h.space_before = Pt(8)
        r_h = p_h.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(11)
        r_h.font.color.rgb = RGBColor(185, 28, 28)
        r_b = p_h.add_run()
        r_b.text = b
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = C_TEXT_SUB

    # Card 3: วิสัยทัศน์โครงการ
    c3 = add_card(slide2, c_lefts[2], 1.75, w_card, h_card, bg_rgb=RGBColor(240, 253, 244), border_rgb=RGBColor(187, 247, 208))
    tf3 = c3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = tf3.margin_right = tf3.margin_top = tf3.margin_bottom = Inches(0.2)
    p = tf3.paragraphs[0]
    r = p.add_run()
    r.text = "เป้าหมายและแนวทางนวัตกรรม\nของโครงงานนี้"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_OK

    pts3 = [
        ("การบูรณาการหลักสูตรสองปริญญา:", "ผสานองค์ความรู้ด้าน Optical Physics (การสะท้อน, การหักเห, กฎโคไซน์) เข้ากับ IoT & 24-bit Embedded Engineering"),
        ("การชดเชยอุณหภูมิหลายจุด (Multi-Point):", "ตรวจวัดอุณหภูมิผิวเซนเซอร์และตัวแปลงสัญญาณเพื่อชดเชยค่าความไวทางความร้อนในแบบเรียลไทม์"),
        ("ต้นแบบต้นทุนต่ำแต่เที่ยงตรงสูง:", "ใช้อุปกรณ์มาตรฐานที่มีจำหน่ายทั่วไป สร้างระบบความแม่นยำสูงในงบประมาณวัสดุต้นแบบไม่เกิน 10,000 บาท")
    ]
    for h, b in pts3:
        p_h = tf3.add_paragraph()
        p_h.space_before = Pt(8)
        r_h = p_h.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(11)
        r_h.font.color.rgb = RGBColor(21, 128, 61)
        r_b = p_h.add_run()
        r_b.text = b
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 3: Scope & Phased Roadmap (Capstone 1 vs Capstone 2)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, C_BG_LIGHT)
    add_header(slide3, "Project Scope & Roadmap", "ขอบเขตการดำเนินงานโครงงาน 1 และแผนพัฒนาโครงงาน 2",
               "การแบ่งเป้าหมายการดำเนินงานตามกรอบระยะเวลาของหลักสูตรอย่างเป็นรูปธรรม")
    add_footer(slide3, 3)

    # 2 Big Cards: Capstone 1 (Done) vs Capstone 2 (Next)
    w_half = 5.7
    h_body = 5.0

    # Left: Capstone 1
    c_p1 = add_card(slide3, 0.8, 1.75, w_half, h_body, bg_rgb=RGBColor(240, 249, 255), border_rgb=C_BLUE_BRAND, border_pt=2)
    tf_p1 = c_p1.text_frame
    tf_p1.word_wrap = True
    tf_p1.margin_left = tf_p1.margin_right = tf_p1.margin_top = tf_p1.margin_bottom = Inches(0.25)
    
    p = tf_p1.paragraphs[0]
    r = p.add_run()
    r.text = "Capstone Project 1 (ผลงานที่สำเร็จในปัจจุบัน)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    items_p1 = [
        ("การจำลองทางทัศนศาสตร์ (Optical Simulation):", "สร้างโมเดล 3D CAD และจำลอง Ray-tracing เปรียบเทียบ BPW 34 vs BPX 61 จนค้นพบปัญหาการตัดแสงของกระจกหน้าต่างตัวถัง TO-39"),
        ("การออกแบบและต่อวงจรทดสอบ AFE บน Breadboard:", "สร้างวงจร 24-bit ADC (ADS1220) ร่วมกับ REF5045 (4.500 V) และ OPA380 TIA"),
        ("การแก้ปัญหา High-Z Leakage และ USB Dropout:", "แก้กระแสรั่วขา Trim ด้วยการ DNP และปรับมาใช้สวิตชิ่ง 5.2V เพื่อก้าวข้าม Dropout"),
        ("ระบบประมวลผลและ IoT Web Dashboard:", "พัฒนาเฟิร์มแวร์ ESP32-C6 สื่อสาร SPI 1MHz, ตรวจจับ Watchdog DRDY และแสดงผลสดผ่าน Web Server"),
        ("การจัดทำข้อเสนอการทดลองทางทัศนศาสตร์ (Lab Proposal):", "ร่างระเบียบวิธีทดสอบจริงด้วยเครื่อง Vernier LabQuest 2 และแท่นหมุน Goniometer")
    ]
    for h, b in items_p1:
        p_i = tf_p1.add_paragraph()
        p_i.space_before = Pt(8)
        r_h = p_i.add_run()
        r_h.text = f"✔ {h}\n"
        r_h.font.bold = True
        r_h.font.size = Pt(11)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = f"   {b}"
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = C_TEXT_SUB

    # Right: Capstone 2
    c_p2 = add_card(slide3, 6.83, 1.75, w_half, h_body, bg_rgb=RGBColor(245, 243, 255), border_rgb=C_PURPLE_IOT, border_pt=2)
    tf_p2 = c_p2.text_frame
    tf_p2.word_wrap = True
    tf_p2.margin_left = tf_p2.margin_right = tf_p2.margin_top = tf_p2.margin_bottom = Inches(0.25)
    
    p = tf_p2.paragraphs[0]
    r = p.add_run()
    r.text = "Capstone Project 2 (แผนงานพัฒนาต่อเนื่อง)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(15)
    r.font.bold = True
    r.font.color.rgb = C_PURPLE_IOT

    items_p2 = [
        ("การประกอบฮาร์ดแวร์ลง Perfboard / Custom PCB:", "ติดตั้งตัวเก็บประจุเซรามิก Low-ESR 10 µF X7R และออกแบบแผ่นวงจรพิมพ์แยกกราวด์แอนะล็อก/ดิจิทัลอย่างสมบูรณ์"),
        ("การทดสอบจริงในห้องปฏิบัติการฟิสิกส์:", "นำชิ้นงานเข้าทดสอบสเปกตรัมและแท่นหมุนร่วมกับอาจารย์ ดร.ประธาน บูรณศิริ ในห้องแล็บทัศนศาสตร์ สจล."),
        ("การติดตั้ง Ultra-Low-Noise LDO Post-Regulator:", "เพิ่มชิป LDO สัญญาณรบกวนต่ำพิเศษ (TPS7A/LP5907) กรอง Switching Ripple สู่เกณฑ์ Noise Floor ระดับ Class A"),
        ("การปรับปรุงตัวถัง Cosine Diffuser:", "วิจัยและทดสอบโดมกระจายแสง (Diffuser Dome) เช่น ควอตซ์หรือเทฟลอน เพื่อแก้ปัญหา Fresnel Drop ที่มุมชัน"),
        ("การสอบเทียบกลางแจ้งเทียบกับมาตรฐาน ISO 9060:", "นำเครื่องวัดเข้าเทียบเคียงกับ Pyranometer อ้างอิงกลางแจ้งจริงตลอด 30 วัน")
    ]
    for h, b in items_p2:
        p_i = tf_p2.add_paragraph()
        p_i.space_before = Pt(8)
        r_h = p_i.add_run()
        r_h.text = f"➔ {h}\n"
        r_h.font.bold = True
        r_h.font.size = Pt(11)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = f"   {b}"
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 4: Overall System Architecture (Block Diagram)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, C_BG_LIGHT)
    add_header(slide4, "System Architecture", "สถาปัตยกรรมระบบวัดและผังวงจรรวม 4 ภาคการทำงาน",
               "โครงสร้างการเชื่อมต่อระหว่างภาคหัววัดแสง, วงจรแอนะล็อก 24 บิต, ไมโครคอนโทรลเลอร์ และภาคจ่ายไฟ 5.2V")
    add_footer(slide4, 4)

    # Embed High-Resolution System Block Diagram on Left
    sys_diag_path = os.path.join(img_dir, "system_block_diagram.png")
    if os.path.exists(sys_diag_path):
        slide4.shapes.add_picture(sys_diag_path, Inches(0.8), Inches(1.85), width=Inches(8.1), height=Inches(4.9))

    # Right: Summary Card of 4 Subsystems
    c_s4 = add_card(slide4, 9.1, 1.85, 3.433, 4.9, bg_rgb=C_CARD_BG, border_rgb=C_BLUE_BRAND, border_pt=1.5)
    tf_s4 = c_s4.text_frame
    tf_s4.word_wrap = True
    tf_s4.margin_left = tf_s4.margin_right = tf_s4.margin_top = tf_s4.margin_bottom = Inches(0.18)
    
    p = tf_s4.paragraphs[0]
    r = p.add_run()
    r.text = "สรุป 4 ภาคการทำงานหลัก"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    subs = [
        ("1. Optical Sensing Head", "BPW 34 + กระจกแซฟไฟร์สังเคราะห์ + OPA380 TIA (R_TIA = 2.2 kΩ) แปลงกระแสแสงเป็นแรงดัน V_diode & V_bias"),
        ("2. 24-bit Analog Front-End", "TI ADS1220 24-bit ΔΣ ADC (20 SPS, Sinc³ Filter) + TI REF5045 Precision Ref (4.500 V, 3 ppm/°C)"),
        ("3. Processing & IoT Platform", "ESP32-C6 Super Mini (RISC-V) สื่อสาร SPI 1 MHz, DRDY Watchdog 150 ms, Web Server (Port 80) & REST API"),
        ("4. Power Architecture", "5.2 V Switching Power Supply ก้าวข้าม Dropout ของ REF5045 พร้อมแผนเพิ่ม Low-Noise LDO ในโครงงาน 2")
    ]
    for h, b in subs:
        p_sub = tf_s4.add_paragraph()
        p_sub.space_before = Pt(6)
        r_h = p_sub.add_run()
        r_h.text = f"{h}:\n"
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_sub.add_run()
        r_b.text = b
        r_b.font.size = Pt(8.8)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 5: Optical Sensor Selection (BPW 34 vs BPX 61)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, C_BG_LIGHT)
    add_header(slide5, "Optical Hardware", "การคัดเลือกหัววัดแสงซิลิคอนและกระจกแซฟไฟร์สังเคราะห์",
               "การวิเคราะห์เปรียบเทียบมิติทางกายภาพ พฤติกรรมเชิงมุม และคุณสมบัติการส่องผ่านทางแสง")
    add_footer(slide5, 5)

    # 3 Cards: BPX 61 vs BPW 34 vs Sapphire
    w3 = 3.65
    h3 = 5.0
    
    # Card 1: BPX 61
    c_bpx = add_card(slide5, c_lefts[0], 1.75, w3, h3)
    tf = c_bpx.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.15)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "OSRAM BPX 61 (TO-39 Metal Can)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    # Embed Polar Plot BPX 61
    p_bpx_img = os.path.join(img_dir, "(Directional Polar Plot) ของ BPX 61 .png")
    if os.path.exists(p_bpx_img):
        slide5.shapes.add_picture(p_bpx_img, Inches(0.95), Inches(2.25), width=Inches(3.35))

    tb_bpx = slide5.shapes.add_textbox(Inches(0.95), Inches(4.7), Inches(3.35), Inches(1.9))
    tf_b = tb_bpx.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    pts_bpx = [
        "ชิปขนาด 7.02 mm² ในตัวถังโลหะผนึกแน่น (Hermetic)",
        "ชิปจมลึกในหลุมกระป๋องถึง 2.2 mm",
        "มีกระจก Schott 8250 บัดกรีติดปากหลุม",
        "มุมรับแสงครึ่งค่า ±55° (ตามสเปก)"
    ]
    for pt in pts_bpx:
        p_pt = tf_b.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"• {pt}"
        r_pt.font.size = Pt(9.5)
        r_pt.font.color.rgb = C_TEXT_SUB

    # Card 2: BPW 34
    c_bpw = add_card(slide5, c_lefts[1], 1.75, w3, h3, bg_rgb=RGBColor(240, 253, 244), border_rgb=C_GREEN_OK)
    tf = c_bpw.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.15)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "OSRAM BPW 34 (Flat Mold) [เซนเซอร์หลัก]"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_OK

    p_bpw_img = os.path.join(img_dir, "(Directional Polar Plot) ของ BPW 34.png")
    if os.path.exists(p_bpw_img):
        slide5.shapes.add_picture(p_bpw_img, Inches(4.99), Inches(2.25), width=Inches(3.35))

    tb_bpw = slide5.shapes.add_textbox(Inches(4.99), Inches(4.7), Inches(3.35), Inches(1.9))
    tf_w = tb_bpw.text_frame
    tf_w.word_wrap = True
    tf_w.margin_left = tf_w.margin_right = tf_w.margin_top = tf_w.margin_bottom = 0
    pts_bpw = [
        "ชิปขนาด 7.02 mm² เท่ากัน แต่เป็นตัวถังแบนราบ",
        "ไม่มีขอบหลุมกระป๋องโลหะมาบังแสงในมุมกว้าง",
        "มุมรับแสงกว้างพิเศษ ±65° ตามกฎโคไซน์",
        "ตอบสนองสอดคล้องกับมาตรฐาน ISO 9060 ดีเยี่ยม"
    ]
    for pt in pts_bpw:
        p_pt = tf_w.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"• {pt}"
        r_pt.font.size = Pt(9.5)
        r_pt.font.color.rgb = C_TEXT_SUB

    # Card 3: Sapphire Window
    c_sap = add_card(slide5, c_lefts[2], 1.75, w3, h3)
    tf = c_sap.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0.15)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "กระจกแซฟไฟร์สังเคราะห์ (Al₂O₃)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_AMBER_WARN

    p_sap_img = os.path.join(img_dir, "sapphire_transmission_spectrum.png")
    if os.path.exists(p_sap_img):
        slide5.shapes.add_picture(p_sap_img, Inches(9.03), Inches(2.25), width=Inches(3.35))

    tb_sap = slide5.shapes.add_textbox(Inches(9.03), Inches(4.7), Inches(3.35), Inches(1.9))
    tf_s = tb_sap.text_frame
    tf_s.word_wrap = True
    tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
    pts_sap = [
        "ความแข็งระดับ 9 Mohs ทนทานต่อการขีดข่วนกลางแจ้ง",
        "การส่องผ่านราบเรียบ 85–88% ตั้งแต่ UV ถึง NIR (200-1100 nm)",
        "ไม่เกิดการดูดกลืนแสงเฉพาะย่านเหมือนแก้วโซดาไลม์",
        "ทำหน้าที่เป็นโดมปกป้องพื้นผิวเซนเซอร์ในระยะยาว"
    ]
    for pt in pts_sap:
        p_pt = tf_s.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"• {pt}"
        r_pt.font.size = Pt(9.5)
        r_pt.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 6: Optical Ray-tracing Discovery: BPX 61 Can Window Cut-off
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, C_BG_LIGHT)
    add_header(slide6, "Major Optical Discovery", "การค้นพบความผิดเพี้ยนเชิงมุมของกระจกตัวถัง TO-39 (BPX 61)",
               "การหักเหของแสงผ่านกระจกหน้าต่างส่งผลให้เกิดการตัดแสง (Cut-off) ร่นลงมาอย่างมีนัยสำคัญ")
    add_footer(slide6, 6)

    # Left: Simulation Result Image
    p_sim_bpx = os.path.join(img_dir, "sim_bpx61_with_without_glass.png")
    if os.path.exists(p_sim_bpx):
        slide6.shapes.add_picture(p_sim_bpx, Inches(0.8), Inches(1.75), width=Inches(6.2))

    # Right: Analysis & Findings Card
    c_find = add_card(slide6, 7.2, 1.75, 5.333, 5.0, bg_rgb=RGBColor(254, 242, 242), border_rgb=RGBColor(239, 68, 68), border_pt=2)
    tf_f = c_find.text_frame
    tf_f.word_wrap = True
    tf_f.margin_left = tf_f.margin_right = tf_f.margin_top = tf_f.margin_bottom = Inches(0.25)
    
    p = tf_f.paragraphs[0]
    r = p.add_run()
    r.text = "ข้อค้นพบสำคัญทางทัศนศาสตร์ (Critical Finding)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(185, 28, 28)

    pts_f = [
        ("การจำลองแบบไม่มีกระจก (Without Glass):", "ลำแสงเดินทางเป็นเส้นตรงในอากาศ แสงเริ่มถูกขอบกระป๋องโลหะบังที่มุม 62° ซึ่งสอดคล้องตรงกับสเปกชีตของผู้ผลิต OSRAM"),
        ("ปรากฏการณ์เมื่อมีกระจก Schott 8250:", "แสงที่ตกกระทบผิวแก้วที่มุมเฉียงจะเกิดการหักเห (Refraction) เบนเข้าหาแนวฉากตามกฎของสเนลล์ (Snell's Law) ทำให้ลำแสงที่ผ่านแก้วเบนลงต่ำกว่าปกติ"),
        ("ผลกระทบทางกายภาพ:", "ลำแสงที่เบนต่ำพุ่งชนเข้ากับขอบด้านในของกระป๋องโลหะ TO-39 เร็วขึ้นอย่างมาก ส่งผลให้จุดตัดแสง (Cut-off) ร่นจาก 62° ลงมาเหลือเพียง 44°!"),
        ("ผลลัพธ์เชิงวิศวกรรม:", "ที่มุม 44° ขึ้นไป เซนเซอร์มืดสนิท (0 W/m²) ทำให้เกิด Directional Error สูงถึง -100% ซึ่งไม่สามารถนำมาใช้เป็น Pyranometer ได้เลย!"),
        ("ข้อสรุปการออกแบบ:", "คณะผู้จัดทำจึงตัดสินใจเลือกใช้โฟโตไดโอดตัวถังแบนราบ BPW 34 เป็นเซนเซอร์หลักเพื่อขจัดปัญหาขอบกระป๋องบังแสงอย่างสิ้นเชิง")
    ]
    for h, b in pts_f:
        p_i = tf_f.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10.5)
        r_h.font.color.rgb = RGBColor(153, 27, 27)
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 7: Optical Simulation of BPW 34 & ISO 9060
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, C_BG_LIGHT)
    add_header(slide7, "Optical Simulation", "ผลการจำลองหัววัด BPW 34 และการประเมินมาตรฐาน ISO 9060",
               "การวิเคราะห์การตอบสนองเชิงมุมของตัวถังแบนราบ และความจำเป็นในการทดสอบจริงในห้องปฏิบัติการ")
    add_footer(slide7, 7)

    # Left: Simulation Result Image
    p_sim_bpw = os.path.join(img_dir, "sim_bpw34_with_without_glass.png")
    if os.path.exists(p_sim_bpw):
        slide7.shapes.add_picture(p_sim_bpw, Inches(0.8), Inches(1.75), width=Inches(6.2))

    # Right: Analysis Card
    c_bpw_a = add_card(slide7, 7.2, 1.75, 5.333, 5.0, bg_rgb=RGBColor(240, 253, 244), border_rgb=C_GREEN_OK, border_pt=1.5)
    tf_ba = c_bpw_a.text_frame
    tf_ba.word_wrap = True
    tf_ba.margin_left = tf_ba.margin_right = tf_ba.margin_top = tf_ba.margin_bottom = Inches(0.25)

    p = tf_ba.paragraphs[0]
    r = p.add_run()
    r.text = "การประเมินผลตามเกณฑ์ ISO 9060"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = RGBColor(21, 128, 61)

    pts_ba = [
        ("BPW 34 แบบไม่มีกระจก (Without Glass):", "เส้นการตอบสนองเชิงมุมสอดคล้องกับกฎโคไซน์อุดมคติสมบูรณ์แบบ ผลการประเมิน Directional Error อยู่ในกรอบ ±10 W/m² ตลอดช่วงมุม 0 ถึง 80° ผ่านเกณฑ์มาตรฐาน ISO 9060 Class A อย่างสมบูรณ์"),
        ("BPW 34 เมื่อติดตั้งกระจกแซฟไฟร์แบน:", "ในช่วงมุม 0° ถึง 50° การตอบสนองยังคงดีเยี่ยม แต่ที่มุมชันเกิน 55° ขึ้นไป การสะท้อนที่ผิวแก้วแซฟไฟร์ (Fresnel Reflection) เพิ่มสูงขึ้นอย่างรวดเร็ว ทำให้ค่าในแบบจำลองตกลงแตะระดับ -80 W/m²"),
        ("ข้อจำกัดของแบบจำลองคณิตศาสตร์:", "โปรแกรมจำลองทางแสงมีข้อจำกัดในการคำนวณการสะท้อนหลายชั้น (Multiple Internal Reflections) และการกระเจิงผิวจริง"),
        ("เหตุผลความจำเป็นที่ต้องทดลองจริง:", "ความจริงทางวิศวกรรมจึงไม่สามารถสรุปได้ด้วยการจำลองคอมพิวเตอร์เพียงอย่างเดียว คณะผู้จัดทำจึงต้องออกแบบข้อเสนอการทดลองจริงในห้องปฏิบัติการเพื่อพิสูจน์ค่าพารามิเตอร์จริง")
    ]
    for h, b in pts_ba:
        p_i = tf_ba.add_paragraph()
        p_i.space_before = Pt(8)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10.5)
        r_h.font.color.rgb = RGBColor(21, 128, 61)
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 8: Proposed Optical Lab Validation (Dr. Prathan Collaboration)
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, C_BG_LIGHT)
    add_header(slide8, "Lab Validation Proposal", "ข้อเสนอระเบียบวิธีการทดลองทางทัศนศาสตร์ในห้องปฏิบัติการ",
               "การวางแผนการทดลองร่วมกับอาจารย์ ดร.ประธาน บูรณศิริ ในห้องปฏิบัติการภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล.")
    add_footer(slide8, 8)

    # 2 Subsections: Test 1 (Spectrometer) & Test 2 (Goniometer)
    w_t = 5.7
    h_t = 5.0

    # Test 1: Vernier LabQuest 2
    c_t1 = add_card(slide8, 0.8, 1.75, w_t, h_t)
    tf_t1 = c_t1.text_frame
    tf_t1.word_wrap = True
    tf_t1.margin_left = tf_t1.margin_right = tf_t1.margin_top = tf_t1.margin_bottom = Inches(0.2)
    p = tf_t1.paragraphs[0]
    r = p.add_run()
    r.text = "ชุดการทดลองที่ 1: ตรวจวัดสเปกตรัมการส่องผ่าน\n(Spectral Transmission Test with Vernier LabQuest 2)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    p_lq = os.path.join(img_dir, "labquest2_vernier_spectrometer.png")
    if os.path.exists(p_lq):
        slide8.shapes.add_picture(p_lq, Inches(1.0), Inches(2.45), width=Inches(5.3))

    tb_t1_txt = slide8.shapes.add_textbox(Inches(1.0), Inches(5.15), Inches(5.3), Inches(1.5))
    tf_t1t = tb_t1_txt.text_frame
    tf_t1t.word_wrap = True
    tf_t1t.margin_left = tf_t1t.margin_right = tf_t1t.margin_top = tf_t1t.margin_bottom = 0
    pts_t1 = [
        "ใช้เครื่องสเปกโตรมิเตอร์ Vernier LabQuest 2 ที่มีในห้องปฏิบัติการฟิสิกส์ สจล.",
        "ตรวจวัดสเปกตรัมช่วงความยาวคลื่น 350.1 – 899.7 nm ครอบคลุมย่านตอบสนองหลัก",
        "วัดเส้นฐานอ้างอิง I_0(λ) (100% Baseline) และวัดการส่องผ่านชิ้นงานแซฟไฟร์จริง I(λ)",
        "คำนวณ %T(λ) = [I(λ)/I_0(λ)] × 100% เพื่อยืนยันว่าการส่องผ่านสม่ำเสมอ 85-88% จริง"
    ]
    for pt in pts_t1:
        p_pt = tf_t1t.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"• {pt}"
        r_pt.font.size = Pt(9)
        r_pt.font.color.rgb = C_TEXT_SUB

    # Test 2: Goniometer Stage
    c_t2 = add_card(slide8, 6.83, 1.75, w_t, h_t)
    tf_t2 = c_t2.text_frame
    tf_t2.word_wrap = True
    tf_t2.margin_left = tf_t2.margin_right = tf_t2.margin_top = tf_t2.margin_bottom = Inches(0.2)
    p = tf_t2.paragraphs[0]
    r = p.add_run()
    r.text = "ชุดการทดลองที่ 2: วัดการตอบสนองเชิงมุมบนแท่นหมุน\n(Directional Cosine Response Test on Goniometer Stage)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = C_PURPLE_IOT

    p_gon = os.path.join(img_dir, "goniometer_blender_perspective.png")
    if os.path.exists(p_gon):
        slide8.shapes.add_picture(p_gon, Inches(7.03), Inches(2.45), width=Inches(5.3))

    tb_t2_txt = slide8.shapes.add_textbox(Inches(7.03), Inches(5.15), Inches(5.3), Inches(1.5))
    tf_t2t = tb_t2_txt.text_frame
    tf_t2t.word_wrap = True
    tf_t2t.margin_left = tf_t2t.margin_right = tf_t2t.margin_top = tf_t2t.margin_bottom = 0
    pts_t2 = [
        "จำลองแบบ 3D Setup ด้วย Blender: ท่อ Collimator + แท่นหมุนองศา (Rotary Stage)",
        "หมุนทดสอบตั้งแต่ 0° ถึง 85° ละเอียดทุก 5° เพื่อพล็อต Polar Cosine Response Curve",
        "ข้อเสนอชุดกำเนิดแสงแสวงเครื่อง (Practical Setup): ใช้หลอดทังสเตน-ฮาโลเจน (DC Bench Power) ร่วมกับท่อ Collimator สีดำและม่าน Iris บีบลำแสงขนาน",
        "ทดสอบเปรียบเทียบ BPW 34 ทั้งกรณีมีและไม่มีกระจก เพื่อนำข้อมูลจริงไปออกแบบ Diffuser"
    ]
    for pt in pts_t2:
        p_pt = tf_t2t.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"• {pt}"
        r_pt.font.size = Pt(9)
        r_pt.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 9: 24-bit Analog Front-End (ADS1220 + REF5045)
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide9, C_BG_LIGHT)
    add_header(slide9, "24-bit Hardware AFE", "การออกแบบวงจรขยายและแปลงสัญญาณความละเอียดสูง 24 บิต",
               "โครงสร้างไอซี ADS1220, แรงดันอ้างอิงความเที่ยงตรงสูง REF5045 และเทคนิค Pseudo-Differential")
    add_footer(slide9, 9)

    # Left: ADS1220 Block Diagram
    p_ads = os.path.join(img_dir, "ADS1220_diagram.png")
    if os.path.exists(p_ads):
        slide9.shapes.add_picture(p_ads, Inches(0.8), Inches(1.75), width=Inches(5.2))

    # Right: Circuit Features Card
    c_afe = add_card(slide9, 6.2, 1.75, 6.333, 5.0)
    tf_af = c_afe.text_frame
    tf_af.word_wrap = True
    tf_af.margin_left = tf_af.margin_right = tf_af.margin_top = tf_af.margin_bottom = Inches(0.2)
    p = tf_af.paragraphs[0]
    r = p.add_run()
    r.text = "คุณสมบัติทางวิศวกรรมของวงจร AFE 24 บิต"
    r.font.name = FONT_HEAD
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    pts_af = [
        ("TI ADS1220 (24-bit Delta-Sigma ADC):", "ความละเอียดการแปลงสัญญาณสูงถึง 24 บิต (16,777,216 ระดับ) สามารถตรวจวัดกระแสแสงโฟโตเคอร์เรนต์ได้ละเอียดถึงระดับ sub-microampere"),
        ("Single-Ended Input MUX vs AVSS (GND):", "ตั้งค่าอ่าน 4 ช่องสัญญาณเทียบกราวด์ (Gain=1, PGA Bypassed) รองรับช่วงแรงดันอินพุตเต็มพิกัด 0 ถึง 4.500 V"),
        ("วงจรกรองดิจิทัล Sinc³ Filter (20 SPS Normal Mode):", "ตัดทอนสัญญาณรบกวนความถี่คู่ควบ 50 Hz/60 Hz จากสายส่งไฟฟ้าได้อย่างสมบูรณ์แบบ"),
        ("แรงดันอ้างอิงภายนอก TI REF5045AID (4.500 V):", "ค่าความคลาดเคลื่อนเริ่มต้นต่ำเพียง ±0.05% และมีอัตราการเลื่อนลอยทางอุณหภูมิต่ำพิเศษ 3 ppm/°C ต่อเข้าขั้ว REFP0/REFN0"),
        ("การประมวลผลผลต่างเทียม (Software Pseudo-Differential):", "อ่านแรงดัน AIN2 (V_diode) และ AIN3 (V_bias) แล้วคำนวณหักลบในไมโครคอนโทรลเลอร์:\nV_diode_actual = V_AIN2 - V_AIN3  |  I_photo = V_diode_actual / 2,200 Ω\nช่วยตัดทอน DC Offset และ Common-Mode Noise ออกได้อย่างสมบูรณ์")
    ]
    for h, b in pts_af:
        p_i = tf_af.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10.5)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 10: Multi-Point Thermal Monitoring & Compensation
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide10, C_BG_LIGHT)
    add_header(slide10, "Thermal Architecture", "สถาปัตยกรรมตรวจวัดอุณหภูมิ 3 จุดเพื่อการชดเชยความร้อน",
               "การแก้ปัญหาค่าสัมประสิทธิ์อุณหภูมิของโฟโตไดโอด (TC_I ≈ +0.18%/K) ด้วยการมอนิเตอร์ความร้อนหลายจุด")
    add_footer(slide10, 10)

    # 3 Cards for 3 Thermal Sensors
    w_th = 3.65
    h_th = 4.2
    
    # Sensor 1: NTC
    c_t1 = add_card(slide10, c_lefts[0], 1.75, w_th, h_th, bg_rgb=RGBColor(240, 249, 255), border_rgb=C_BLUE_BRAND)
    tf1 = c_t1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_right = tf1.margin_top = tf1.margin_bottom = Inches(0.2)
    p = tf1.paragraphs[0]
    r = p.add_run()
    r.text = "1. อุณหภูมิผิวหัววัดแสง\n(Surface NTC Thermistor)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    pts_ntc = [
        ("ช่องสัญญาณ ADS1220:", "AIN1 (0x91) Single-Ended"),
        ("เบอร์อุปกรณ์:", "Thinking TTC3A103F39HAEY"),
        ("คุณสมบัติทางเทคนิค:", "10 kΩ ±1% ที่ 25°C, ค่าสัมประสิทธิ์ B₂₅/₈₅ = 3950 K"),
        ("หน้าที่การทำงาน:", "ติดตั้งแนบชิดกับตัวถังโฟโตไดโอด วัดอุณหภูมิผิวจริง (T_surface) เพื่อนำไปคำนวณชดเชยค่ากระแสโฟโตเคอร์เรนต์ในสูตรทางคณิตศาสตร์")
    ]
    for h, b in pts_ntc:
        p_i = tf1.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # Sensor 2: REF5045 Temp Pin 3
    c_t2 = add_card(slide10, c_lefts[1], 1.75, w_th, h_th, bg_rgb=RGBColor(254, 243, 199), border_rgb=C_AMBER_WARN)
    tf2 = c_t2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = tf2.margin_right = tf2.margin_top = tf2.margin_bottom = Inches(0.2)
    p = tf2.paragraphs[0]
    r = p.add_run()
    r.text = "2. อุณหภูมิไอซีแรงดันอ้างอิง\n(REF5045 Pin 3 TEMP)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_AMBER_WARN

    pts_ref = [
        ("ช่องสัญญาณ ADS1220:", "AIN0 (0x81) Single-Ended"),
        ("เบอร์อุปกรณ์:", "TI REF5045AID (Pin 3 TEMP)"),
        ("คุณสมบัติทางเทคนิค:", "V_TEMP = 509 mV + 2.64 mV/°C ความไวเชิงเส้นสูง"),
        ("หน้าที่การทำงาน:", "ตรวจวัดอุณหภูมิของตัวถังไอซีสร้างแรงดันอ้างอิง แปลงค่าด้วยสูตร:\nT_REF = (V_AIN0 - 0.509) / 0.00264\nเพื่อตรวจสอบเสถียรภาพความร้อนของระบบ")
    ]
    for h, b in pts_ref:
        p_i = tf2.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # Sensor 3: Internal TS=1
    c_t3 = add_card(slide10, c_lefts[2], 1.75, w_th, h_th, bg_rgb=RGBColor(240, 253, 244), border_rgb=C_GREEN_OK)
    tf3 = c_t3.text_frame
    tf3.word_wrap = True
    tf3.margin_left = tf3.margin_right = tf3.margin_top = tf3.margin_bottom = Inches(0.2)
    p = tf3.paragraphs[0]
    r = p.add_run()
    r.text = "3. อุณหภูมิภายในตัวชิป ADC\n(ADS1220 Internal Junction)"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_OK

    pts_ads = [
        ("โหมดการทำงาน:", "Register 1 บิต TS = 1 (0x02)"),
        ("เซนเซอร์ภายใน:", "Internal Temperature Diode Junction"),
        ("คุณสมบัติทางเทคนิค:", "ความละเอียด 14 บิต (0.03125 °C/LSB)"),
        ("หน้าที่การทำงาน:", "ตรวจวัดอุณหภูมิรอยต่อซิลิคอนภายในตัวชิปแปลงสัญญาณ ADC โดยตรง แปลงค่าด้วยสูตร:\nT_ADS = RawValue × 0.03125 °C\nเฝ้าระวัง Thermal Drift ภายในระบบ")
    ]
    for h, b in pts_ads:
        p_i = tf3.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = b
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # Bottom Formula Banner
    c_ban = add_card(slide10, 0.8, 6.1, 11.733, 0.75, bg_rgb=RGBColor(248, 250, 252), border_rgb=RGBColor(203, 213, 225))
    tf_ban = c_ban.text_frame
    tf_ban.margin_left = tf_ban.margin_right = tf_ban.margin_top = tf_ban.margin_bottom = 0
    p_b = tf_ban.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = "สมการชดเชยกระแสแสงในเฟิร์มแวร์:   I_corrected = I_photo / [ 1 + TC_I × (T_surface - 25°C) ]    โดยที่ TC_I = +0.18%/K"
    r_b.font.name = FONT_HEAD
    r_b.font.size = Pt(11)
    r_b.font.bold = True
    r_b.font.color.rgb = C_BLUE_BRAND

    # =========================================================================
    # SLIDE 11: Hardware Prototyping & Breadboard Debugging Discoveries
    # =========================================================================
    slide11 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide11, C_BG_LIGHT)
    add_header(slide11, "Prototyping & Debugging", "การสร้างต้นแบบจริงบน Breadboard และการแก้ปัญหาเชิงวิศวกรรม",
               "การค้นพบทางวิศวกรรมที่สำคัญ 2 ประการจากการทดลองประกอบวงจรจริงในห้องปฏิบัติการ")
    add_footer(slide11, 11)

    # Left: Photo of Real Prototype
    p_bb = os.path.join(img_dir, "power_switching_prototype.jpg")
    if not os.path.exists(p_bb):
        p_bb = os.path.join(img_dir, "breadboard_real_photo.jpg")
    if os.path.exists(p_bb):
        slide11.shapes.add_picture(p_bb, Inches(0.8), Inches(1.75), width=Inches(5.0))

    # Right: 2 Discoveries Card
    c_disc = add_card(slide11, 6.0, 1.75, 6.533, 5.0)
    tf_d = c_disc.text_frame
    tf_d.word_wrap = True
    tf_d.margin_left = tf_d.margin_right = tf_d.margin_top = tf_d.margin_bottom = Inches(0.2)
    p = tf_d.paragraphs[0]
    r = p.add_run()
    r.text = "ประเด็นปัญหาและการแก้ไขเชิงวิศวกรรมบน Breadboard"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    pts_disc = [
        ("1. ปัญหา High-Z Node Leakage บนขา 5 (Trim/NR):", 
         "ในการต่อวงจรช่วงแรก แรงดันตกเหลือ 4.382 V (หายไป 118 mV) ตรวจสอบพบว่าขา 5 เป็นโหนดความต้านทานสูงมากและเกิดกระแสไฟฟ้ารั่วบนรางเบรดบอร์ด -> แก้ไขโดยการเว้นว่าง (DNP) การเชื่อมต่อขา 5 ทันที"),
        ("2. ปัญหา USB Dropout Voltage จากสายคอมพิวเตอร์:", 
         "เมื่อจ่ายไฟผ่าน USB 5V ของคอมพิวเตอร์ แรงดันตกเหลือต่ำกว่า 4.70 V ทำให้ REF5045 ตกสภาวะ Dropout (ต้องการ V_IN ≥ V_REF + 200 mV) เอาต์พุตแกว่งตามโหลด PC"),
        ("3. การแก้ปัญหาด้วย Switching Power Supply 5.2 V:", 
         "เปลี่ยนมาใช้แหล่งจ่ายไฟตรงแบบสวิตชิ่งปรับแรงดันที่ 5.20 V สร้างส่วนต่างแรงดันตกคร่อมสูงถึง 700 mV (> 200 mV) ทำให้ REF5045 ควบคุมแรงดันอ้างอิงนิ่งสนิทที่ 4.491 V (คลาดเคลื่อนเพียง -0.2% ในเกณฑ์ DMM)"),
        ("4. แผนงานการประกอบ Perfboard และชิป Low-Noise LDO:", 
         "ได้วิเคราะห์พบว่าความเหนี่ยวนำแฝงบนรางเบรดบอร์ดลดทอน Phase Margin จึงได้สั่งซื้อตัวเก็บประจุเซรามิก Low-ESR 10 µF X7R เพื่อนำมาบัดกรีลง Perfboard ในคราวเดียว และวางแผนต่อยอดด้วย LDO (TPS7A) ใน Capstone 2")
    ]
    for h, b in pts_disc:
        p_i = tf_d.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h}\n"
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = f"   {b}"
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 12: Custom 4-Layer PCB Design (V1.1 L1 by Tachin)
    # =========================================================================
    slide12 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide12, C_BG_LIGHT)
    add_header(slide12, "Hardware Engineering", "การออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom 4-Layer Sensor Head PCB)",
               "รุ่น V1.1 L1 by Tachin — ออกแบบตามหลัก Low-Noise Precision Analog บน KiCad 10 และผ่านการตรวจ DRC 100%")
    add_footer(slide12, 12)

    # Left: Two 3D Renders (Top View & Isometric View)
    p_pcb_top = os.path.join(img_dir, "pcb_3d_render_top.png")
    p_pcb_iso = os.path.join(img_dir, "pcb_3d_render_iso.png")
    if os.path.exists(p_pcb_top) and os.path.exists(p_pcb_iso):
        slide12.shapes.add_picture(p_pcb_top, Inches(0.8), Inches(1.75), width=Inches(2.7))
        slide12.shapes.add_picture(p_pcb_iso, Inches(3.6), Inches(1.75), width=Inches(2.7))
    elif os.path.exists(p_pcb_iso):
        slide12.shapes.add_picture(p_pcb_iso, Inches(0.8), Inches(1.75), width=Inches(5.5))

    # Right: PCB Architecture Card
    c_pcb = add_card(slide12, 6.5, 1.75, 6.033, 5.0)
    tf_pcb = c_pcb.text_frame
    tf_pcb.word_wrap = True
    tf_pcb.margin_left = tf_pcb.margin_right = tf_pcb.margin_top = tf_pcb.margin_bottom = Inches(0.2)
    p = tf_pcb.paragraphs[0]
    r = p.add_run()
    r.text = "สถาปัตยกรรมการออกแบบแผ่นวงจรพิมพ์ V1.1 L1"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_BRAND

    pts_pcb = [
        ("โครงสร้างทรงกลมกะทัดรัด (Circular Sensor Head):", 
         "ออกแบบแผ่นวงจรรูปทรงกลมเพื่อติดตั้งลงในกระบอกสูบตัวเรือนเซนเซอร์วัดรังสีอาทิตย์ได้อย่างสมบูรณ์แบบ ลดขนาดสายเชื่อมต่อภายนอก"),
        ("ตำแหน่งเซนเซอร์และ RT1 สมดุลกึ่งกลาง (Co-located Sensing):", 
         "จัดวาง BPX 61 ตรงกึ่งกลางบอร์ด พร้อมวาง RT1 (Thermistor 103AT-2) ประกบชิดผิวเซนเซอร์ เพื่อตรวจวัดอุณหภูมิจริงและทำ Temperature Compensation อย่างแม่นยำ"),
        ("ยกระดับสู่อุปกรณ์ความเที่ยงตรงสูง (Precision IC Upgrades):", 
         "ปรับใช้ REF5045AIDGKR ในตัวถัง VSSOP-8 (MSOP-8) ขนาดเล็กพิเศษ ร่วมกับ OPA380 TIA และชิป 24-bit ADC (ADS1220 TSSOP-16) ลดขนาด Loop และสัญญาณรบกวน"),
        ("สถาปัตยกรรม 4 เลเยอร์ (Continuous Ground Plane):", 
         "จัดสรร In1.Cu เป็น Solid Ground Plane ไร้รอยตัด ช่วยขจัดปัญหา Ground Bounce, IR Drop (~5-10 mV) และ Stray Inductance ที่เคยพบจากการต่อสาย Point-to-point"),
        ("พร้อมส่งผลิตโรงงาน (100% DRC Zero Violations):", 
         "ผ่านการตรวจสอบ Design Rules Check (DRC) บน KiCad 10.0 ครบถ้วน ไร้ข้อผิดพลาด (0 Violations / 0 Unconnected items) พร้อมส่งผลิต (PCBWay/JLCPCB) ใน Capstone 2")
    ]
    for h, b in pts_pcb:
        p_i = tf_pcb.add_paragraph()
        p_i.space_before = Pt(5)
        r_h = p_i.add_run()
        r_h.text = f"• {h}\n"
        r_h.font.bold = True
        r_h.font.size = Pt(9.8)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = f"   {b}"
        r_b.font.size = Pt(9.2)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 13: Embedded Firmware & IoT Platform
    # =========================================================================
    slide13 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide13, C_BG_LIGHT)
    add_header(slide13, "Firmware & IoT Platform", "เฟิร์มแวร์ไมโครคอนโทรลเลอร์ ESP32-C6 และระบบ IoT Web Server",
               "การควบคุมบัสสื่อสาร SPI, การเฝ้าระวัง Watchdog และการรายงานข้อมูลสดแบบไร้สายผ่านเครือข่าย Wi-Fi")
    add_footer(slide13, 13)

    # Left: Serial Monitor Output Image
    p_ser = os.path.join(img_dir, "Read value from ardunio.png")
    if os.path.exists(p_ser):
        slide12.shapes.add_picture(p_ser, Inches(0.8), Inches(1.75), width=Inches(5.5))

    # Right: Firmware Highlights Card
    c_fw = add_card(slide12, 6.5, 1.75, 6.033, 5.0)
    tf_fw = c_fw.text_frame
    tf_fw.word_wrap = True
    tf_fw.margin_left = tf_fw.margin_right = tf_fw.margin_top = tf_fw.margin_bottom = Inches(0.2)
    p = tf_fw.paragraphs[0]
    r = p.add_run()
    r.text = "โครงสร้างเฟิร์มแวร์และการทำงานของระบบ IoT"
    r.font.name = FONT_HEAD
    r.font.size = Pt(13.5)
    r.font.bold = True
    r.font.color.rgb = C_PURPLE_IOT

    pts_fw = [
        ("ESP32-C6 Super Mini (RISC-V 32-bit Architecture):", "ไมโครคอนโทรลเลอร์สมรรถนะสูง รองรับ Wi-Fi 6 (2.4 GHz) และการประมวลผลสัญญาณดิจิทัลแบบประหยัดพลังงาน"),
        ("Hardware SPI Master (1 MHz, Mode 1):", "กำหนดขาเชื่อมต่อชัดเจน: SCLK (GPIO 6), MISO (GPIO 2), MOSI (GPIO 4), CS (GPIO 14) ส่งคำสั่งและอ่านผลแปลง 24 บิตได้อย่างแม่นยำ"),
        ("ระบบเฝ้าระวังและกู้คืนอัตโนมัติ (DRDY Watchdog Protection):", "ขา DRDY (GPIO 0) ตรวจจับความพร้อมของการแปลงสัญญาณ หากเกิด Timeout เกิน 150 ms ระบบจะส่งคำสั่ง CMD_RESET (0x06) เพื่อกู้คืน ADC ทันที"),
        ("Embedded Web Server (Port 80):", "ให้บริการหน้าจอ Real-time Web Dashboard แบบ Responsive ดึงข้อมูลสดผ่าน AJAX/Fetch API ทุก 1 วินาที โดยไม่ต้องรีเฟรชหน้าเว็บ"),
        ("RESTful JSON API (/data):", "เปิดช่องทางบริการข้อมูล JSON เพื่อให้ระบบคลาวด์หรือคอมพิวเตอร์ภายนอกดึงข้อมูลไปบันทึกและวิเคราะห์ทางสถิติต่อได้ทันที")
    ]
    for h, b in pts_fw:
        p_i = tf_fw.add_paragraph()
        p_i.space_before = Pt(6)
        r_h = p_i.add_run()
        r_h.text = f"• {h}\n"
        r_h.font.bold = True
        r_h.font.size = Pt(10)
        r_h.font.color.rgb = C_TEXT_MAIN
        r_b = p_i.add_run()
        r_b.text = f"   {b}"
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 14: Live Experimental Results on Breadboard
    # =========================================================================
    slide14 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide14, C_BG_LIGHT)
    add_header(slide14, "Experimental Results", "ผลการทดสอบการอ่านค่าจริงบน Breadboard ใน 4 สภาวะแสง",
               "การพิสูจน์การตอบสนองเชิงไดนามิกและการตัดทอนออฟเซ็ตด้วยเทคนิค Pseudo-Differential")
    add_footer(slide14, 14)

    # Big Clean Table for Results
    rows = 5
    cols = 7
    left = Inches(0.8)
    top = Inches(1.85)
    width = Inches(11.733)
    height = Inches(3.2)

    table_shape = slide14.shapes.add_table(rows, cols, left, top, width, height)
    t = table_shape.table

    # Column widths
    t.columns[0].width = Inches(2.7)
    t.columns[1].width = Inches(1.5)
    t.columns[2].width = Inches(1.5)
    t.columns[3].width = Inches(1.5)
    t.columns[4].width = Inches(1.5)
    t.columns[5].width = Inches(1.5)
    t.columns[6].width = Inches(1.533)

    headers = ["สภาวะแสงทดสอบ", "AIN0: Temp REF", "AIN1: NTC", "AIN2: V_diode", "AIN3: V_bias", "Net Diode", "I_photo คำนวณ"]
    for j, h in enumerate(headers):
        cell = t.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(241, 245, 249)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h
        r.font.name = FONT_HEAD
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_MAIN

    data_rows = [
        ["1. ปิดไฟมืดสนิท (Dark / Offset)", "0.6008 V (34.77°C)", "1.0538 V (34.75°C)", "0.00282 V", "0.00280 V", "+0.020 mV", "0.009 µA"],
        ["2. เปิดแสงสว่างปกติในห้อง", "0.6010 V (34.85°C)", "1.0538 V (34.75°C)", "0.02480 V", "0.00278 V", "+22.020 mV", "10.009 µA"],
        ["3. เปิดแสง LED สีส้มสว่างสุด", "0.6019 V (35.19°C)", "1.0531 V (34.92°C)", "0.05780 V", "0.00278 V", "+55.020 mV", "25.009 µA"],
        ["4. เปิดแสงส้ม + ขาวพร้อมกัน", "0.6022 V (35.30°C)", "1.0526 V (35.12°C)", "0.13482 V", "0.00276 V", "+132.060 mV", "60.027 µA"]
    ]

    for i, row in enumerate(data_rows):
        for j, val in enumerate(row):
            cell = t.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(255, 255, 255) if i % 2 == 0 else RGBColor(248, 250, 252)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_BODY
            r.font.size = Pt(10)
            if j == 5:
                r.font.bold = True
                r.font.color.rgb = C_BLUE_BRAND
            elif j == 6:
                r.font.bold = True
                r.font.color.rgb = C_GREEN_OK
            else:
                r.font.color.rgb = C_TEXT_MAIN

    # Summary Takeaway Banner Below Table
    c_res = add_card(slide14, 0.8, 5.3, 11.733, 1.4, bg_rgb=RGBColor(240, 253, 244), border_rgb=C_GREEN_OK)
    tf_r = c_res.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = tf_r.margin_bottom = Inches(0.15)
    p = tf_r.paragraphs[0]
    r = p.add_run()
    r.text = "บทวิเคราะห์ผลลัพธ์เชิงวิศวกรรมจากการทดลองจริง:"
    r.font.name = FONT_HEAD
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(21, 128, 61)

    pts_res = [
        "ยืนยันการตอบสนองเชิงเส้น: เมื่อเพิ่มความเข้มแสง กระแสโฟโตเคอร์เรนต์ I_photo เพิ่มขึ้นอย่างต่อเนื่องจาก 0.009 µA ในความมืด สู่ 60.027 µA ที่แสงผสมสว่างสุด",
        "ความมีประสิทธิภาพของ Pseudo-Differential: แรงดันออฟเซ็ตฐาน V_bias (ประมาณ 2.76–2.80 mV) ถูกหักลบออกอย่างแม่นยำ ทำให้สัญญาณ Net Diode เหลือเพียงกระแสแสงบริสุทธิ์",
        "เสถียรภาพอุณหภูมิ: อุณหภูมิ REF5045 และ NTC สอดคล้องกันอยู่ในช่วง 34.7°C ถึง 35.3°C ยืนยันความพร้อมของฮาร์ดแวร์สำหรับการขึ้นรูปแผงวงจรจริง"
    ]
    for pt in pts_res:
        p_pt = tf_r.add_paragraph()
        r_pt = p_pt.add_run()
        r_pt.text = f"✔ {pt}"
        r_pt.font.size = Pt(9.5)
        r_pt.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 15: Capstone 2 Roadmap & Implementation Milestones
    # =========================================================================
    slide15 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide15, C_BG_LIGHT)
    add_header(slide15, "Future Roadmap", "แผนการดำเนินงานระยะถัดไปในโครงงาน 2 (Capstone Project 2)",
               "กรอบระยะเวลาและเป้าหมายการยกระดับความเที่ยงตรงสู่มาตรฐานสากล ISO 9060 Class A")
    add_footer(slide15, 15)

    # 4 Milestone Horizontal Cards
    w_m = 2.75
    h_m = 4.8
    m_lefts = [0.8, 3.8, 6.8, 9.8]

    milestones = [
        ("Milestone 1 (พ.ย. - ธ.ค. 2569)", "Hardware Fabrication & LDO Post-Regulation", C_BLUE_BRAND, [
            "บัดกรีประกอบวงจรลงบน Perfboard",
            "ติดตั้งตัวเก็บประจุ 10 µF X7R Low-ESR",
            "สั่งผลิตแผ่นวงจรพิมพ์ Custom 4-Layer PCB",
            "เพิ่ม Ultra-Low-Noise LDO (TPS7A)",
            "ทดสอบระดับ Noise Floor < 4 µV_RMS"
        ]),
        ("Milestone 2 (ม.ค. - ก.พ. 2570)", "Optics Lab Empirical Validation", C_PURPLE_IOT, [
            "เข้าทดสอบจริงในห้องแล็บภาควิชาฟิสิกส์",
            "วัดสเปกตรัมด้วย Vernier LabQuest 2",
            "ทดสอบการตอบสนองบนแท่นหมุน Goniometer",
            "เปรียบเทียบผลจริงกับผลซิมของ BPW 34",
            "ขอรับคำปรึกษาจาก ผศ.ดร.ประธาน บูรณศิริ"
        ]),
        ("Milestone 3 (มี.ค. 2570)", "Diffuser Dome & Housing Optimization", C_AMBER_WARN, [
            "วิจัยและทดสอบ Cosine Diffuser Dome",
            "ทดลองวัสดุ Quartz และ Optical Teflon",
            "ปรับปรุงเพื่อแก้ปัญหา Fresnel Drop มุมชัน",
            "สร้างตัวถัง Weatherproof กันน้ำกันฝุ่น",
            "ติดตั้งระบบระบายความร้อนป้องกันความชื้น"
        ]),
        ("Milestone 4 (เม.ย. - พ.ค. 2570)", "Field Testing & ISO 9060 Calibration", C_GREEN_OK, [
            "ติดตั้งทดสอบภาคสนามกลางแจ้ง 30 วัน",
            "สอบเทียบเทียบกับ Pyranometer อ้างอิง",
            "ประเมินค่า Uncertainty ทางมาตรวิทยา",
            "พัฒนาแดชบอร์ด IoT Cloud แบบสมบูรณ์",
            "สรุปผลการวิจัยและเล่มวิทยานิพนธ์ฉบับสมบูรณ์"
        ])
    ]

    for idx, (m_title, m_sub, m_col, m_pts) in enumerate(milestones):
        card = add_card(slide15, m_lefts[idx], 1.85, w_m, h_m, border_rgb=m_col, border_pt=1.5)
        tf_m = card.text_frame
        tf_m.word_wrap = True
        tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = Inches(0.15)
        
        p1 = tf_m.paragraphs[0]
        r1 = p1.add_run()
        r1.text = m_title
        r1.font.name = FONT_HEAD
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = m_col

        p2 = tf_m.add_paragraph()
        p2.space_after = Pt(8)
        r2 = p2.add_run()
        r2.text = m_sub
        r2.font.name = FONT_HEAD
        r2.font.size = Pt(9.5)
        r2.font.bold = True
        r2.font.color.rgb = C_TEXT_MAIN

        for pt in m_pts:
            p_pt = tf_m.add_paragraph()
            p_pt.space_before = Pt(4)
            r_pt = p_pt.add_run()
            r_pt.text = f"• {pt}"
            r_pt.font.size = Pt(8.8)
            r_pt.font.color.rgb = C_TEXT_SUB

    # =========================================================================
    # SLIDE 16: Conclusion & Discussion / Q&A (Dark Executive Theme)
    # =========================================================================
    slide16 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide16, C_NAVY_DARK)

    # Top Section: Title & 3 Accomplishment Bullets
    tb_end = slide16.shapes.add_textbox(Inches(1.2), Inches(0.65), Inches(10.933), Inches(3.3))
    tf_e = tb_end.text_frame
    tf_e.word_wrap = True
    tf_e.margin_left = tf_e.margin_right = tf_e.margin_top = tf_e.margin_bottom = 0

    p = tf_e.paragraphs[0]
    p.space_after = Pt(8)
    r = p.add_run()
    r.text = "สรุปผลสัมฤทธิ์โครงงาน 1 และขอคำชี้แนะจากอาจารย์ที่ปรึกษา"
    r.font.name = FONT_HEAD
    r.font.size = Pt(24)
    r.font.bold = True
    r.font.color.rgb = C_CYAN_LIGHT

    bullets = [
        ("ความสำเร็จด้านฮาร์ดแวร์และซอฟต์แวร์:", "พิสูจน์การทำงานจริงของวงจร AFE 24 บิต บน Breadboard ร่วมกับ REF5045 (4.500 V) และ ESP32-C6 Web Server รายงานข้อมูลได้สำเร็จสมบูรณ์"),
        ("การออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์ (V1.1 L1):", "ออกแบบ Custom 4-Layer PCB สมบูรณ์แบบด้วย KiCad 10.0 ผ่านการตรวจสอบ DRC 100% (Zero Violations) พร้อมส่งผลิตสำหรับ Capstone 2"),
        ("การค้นพบสำคัญทางทัศนศาสตร์:", "ชี้ชัดว่าตัวถัง TO-39 (BPX 61) มีการตัดแสงรุนแรงที่ 44° และยืนยันว่า BPW 34 ให้การตอบสนองที่สอดคล้องกับมาตรฐาน ISO 9060 Class A มากกว่า"),
        ("ความพร้อมในการทดสอบจริง:", "จัดทำข้อเสนอการทดลองทางทัศนศาสตร์ (Proposed Lab Methodology) ไว้อย่างชัดเจน ทั้งแผนวัดสเปกตรัมและแผนวัดมุมด้วยแท่นหมุน")
    ]
    for h, b in bullets:
        p_b = tf_e.add_paragraph()
        p_b.space_before = Pt(6)
        r_h = p_b.add_run()
        r_h.text = f"✔ {h} "
        r_h.font.bold = True
        r_h.font.size = Pt(11.5)
        r_h.font.color.rgb = C_TEXT_LIGHT
        r_b = p_b.add_run()
        r_b.text = b
        r_b.font.size = Pt(10.5)
        r_b.font.color.rgb = C_TEXT_MUTED

    # Middle Section: Questions for Dr. Prathan Banner Card
    c_q = add_card(slide16, 1.2, 4.05, 10.933, 2.15, bg_rgb=C_SLATE_CARD, border_rgb=C_CYAN_ACCENT, border_pt=1.5)
    tf_q = c_q.text_frame
    tf_q.word_wrap = True
    tf_q.margin_left = tf_q.margin_right = tf_q.margin_top = tf_q.margin_bottom = Inches(0.18)
    
    pq = tf_q.paragraphs[0]
    rq = pq.add_run()
    rq.text = "ประเด็นขอคำแนะนำและปรึกษาอาจารย์ ดร.ประธาน บูรณศิริ (Discussion Points):"
    rq.font.name = FONT_HEAD
    rq.font.size = Pt(12.5)
    rq.font.bold = True
    rq.font.color.rgb = C_CYAN_LIGHT

    q_items = [
        "1. ความเป็นไปได้และช่วงเวลาที่เหมาะสมในการนำอุปกรณ์เข้าทดสอบกับชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์ สจล.",
        "2. คำแนะนำเกี่ยวกับข้อเสนอการจัดชุดแสงขนานแสวงเครื่อง (ท่อ Collimator + หลอดทังสเตน-ฮาโลเจน) บนแท่นหมุน Goniometer",
        "3. แนวทางการออกแบบโดมกระจายแสง (Cosine Diffuser) เพื่อยกระดับความเที่ยงตรงสู่มาตรฐานสากล ISO 9060 Class A ในโครงงาน 2"
    ]
    for q in q_items:
        p_qi = tf_q.add_paragraph()
        p_qi.space_before = Pt(4)
        r_qi = p_qi.add_run()
        r_qi.text = q
        r_qi.font.size = Pt(10)
        r_qi.font.color.rgb = C_TEXT_LIGHT

    # Bottom Section: Thank you / Q&A Box
    tb_ty = slide16.shapes.add_textbox(Inches(1.2), Inches(6.35), Inches(10.933), Inches(0.6))
    tf_ty = tb_ty.text_frame
    tf_ty.margin_left = tf_ty.margin_right = tf_ty.margin_top = tf_ty.margin_bottom = 0
    p_ty = tf_ty.paragraphs[0]
    p_ty.alignment = PP_ALIGN.CENTER
    r_ty = p_ty.add_run()
    r_ty.text = "THANK YOU  |  Q & A"
    r_ty.font.name = FONT_HEAD
    r_ty.font.size = Pt(18)
    r_ty.font.bold = True
    r_ty.font.color.rgb = C_CYAN_LIGHT

    out_file = os.path.join(script_dir, "69_IoT_Pyranometer_Capstone1_Slides.pptx")
    prs.save(out_file)
    print(f"[OK] PowerPoint Presentation successfully created at: {out_file}")
    return out_file

if __name__ == "__main__":
    create_presentation()
