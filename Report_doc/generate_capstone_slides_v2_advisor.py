import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def create_presentation_v2():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    script_dir = os.path.dirname(os.path.abspath(__file__))
    img_dir = os.path.join(script_dir, "image")
    ext_dir = os.path.join(img_dir, "v2_extracted")
    iso_dir = os.path.join(img_dir, "iso9060_crops")
    web_dir = os.path.join(img_dir, "web_mockup")

    FONT_THAI = "TH SarabunPSK"
    FONT_HEAD = "Segoe UI"

    # Clean Minimalist Palette
    C_WHITE = RGBColor(255, 255, 255)        # #FFFFFF
    C_DARK_BG = RGBColor(17, 24, 39)         # #111827
    C_BLUE_TAG = RGBColor(2, 132, 199)       # #0284C7
    C_ORANGE_TAG = RGBColor(217, 119, 6)     # #D97706
    C_TEXT_BLACK = RGBColor(15, 23, 42)      # #0F172A
    C_TEXT_BODY = RGBColor(51, 65, 85)       # #334155
    C_TEXT_MUTED = RGBColor(100, 116, 139)   # #64748B
    C_CARD_GRAY = RGBColor(248, 250, 252)    # #F8FAFC
    C_BORDER_GRAY = RGBColor(226, 232, 240)  # #E2E8F0
    C_NAVY_HEADER = RGBColor(15, 23, 42)     # #0F172A
    C_ROW_ALT = RGBColor(241, 245, 249)      # #F1F5F9
    C_RED_BG = RGBColor(254, 242, 242)       # #FEF2F2
    C_RED_BORDER = RGBColor(254, 202, 202)   # #FECACA
    C_RED_TEXT = RGBColor(153, 27, 27)       # #991B1B
    C_GREEN_BG = RGBColor(240, 253, 244)     # #F0FDF4
    C_GREEN_BORDER = RGBColor(187, 247, 208) # #BBF7D0
    C_GREEN_TEXT = RGBColor(22, 101, 52)     # #166534
    C_BLUE_BG = RGBColor(240, 249, 255)      # #F0F9FF
    C_AMBER_BG = RGBColor(254, 252, 232)     # #FEFCE8
    C_AMBER_BORDER = RGBColor(253, 230, 138) # #FDE68A

    slide_counter = [0]

    def new_slide(bg_color=C_WHITE, show_num=True):
        slide_counter[0] += 1
        s = prs.slides.add_slide(blank_layout)
        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = bg_color
        bg.line.fill.background()
        if show_num:
            tb_num = s.shapes.add_textbox(Inches(12.1), Inches(0.35), Inches(0.85), Inches(0.45))
            tf_n = tb_num.text_frame
            p_n = tf_n.paragraphs[0]
            p_n.alignment = PP_ALIGN.RIGHT
            r_n = p_n.add_run()
            r_n.text = str(slide_counter[0])
            r_n.font.name = FONT_HEAD
            r_n.font.size = Pt(16)
            r_n.font.bold = True
            r_n.font.color.rgb = C_TEXT_MUTED
        return s

    def add_clean_header(slide, tag_text, title_text, subtitle_text="", tag_color=C_BLUE_TAG):
        tb = slide.shapes.add_textbox(Inches(0.85), Inches(0.40), Inches(11.3), Inches(1.25))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_tag = tf.paragraphs[0]
        p_tag.space_after = Pt(1)
        r_tag = p_tag.add_run()
        r_tag.text = tag_text
        r_tag.font.name = FONT_THAI
        r_tag.font.size = Pt(18.5)
        r_tag.font.bold = True
        r_tag.font.color.rgb = tag_color

        p_title = tf.add_paragraph()
        p_title.space_after = Pt(2)
        r_title = p_title.add_run()
        r_title.text = title_text
        r_title.font.name = FONT_THAI
        r_title.font.size = Pt(30)
        r_title.font.bold = True
        r_title.font.color.rgb = C_TEXT_BLACK

        if subtitle_text:
            p_sub = tf.add_paragraph()
            r_sub = p_sub.add_run()
            r_sub.text = subtitle_text
            r_sub.font.name = FONT_THAI
            r_sub.font.size = Pt(17.5)
            r_sub.font.color.rgb = C_TEXT_BODY

    def add_card(slide, left, top, width, height, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY, border_pt=1):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        card.adjustments[0] = 0.04
        card.fill.solid()
        card.fill.fore_color.rgb = bg_rgb
        if border_rgb is not None and border_pt > 0:
            card.line.color.rgb = border_rgb
            card.line.width = Pt(border_pt)
        else:
            card.line.fill.background()
        # Ensure text frame inside card always anchors to TOP and never overflows vertically
        tf = card.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = Inches(0.20)
        tf.margin_right = Inches(0.20)
        tf.margin_top = Inches(0.14)
        tf.margin_bottom = Inches(0.10)
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
    # บทที่ 1: ที่มาและความสำคัญ วัตถุประสงค์ และภาพรวมระบบ (Slides 1 - 5)
    # =========================================================================

    # SLIDE 1: หน้าปก (เพิ่มสาขา + รหัสโปรเจกต์ + เลขสไลด์ ตามคอมเมนต์อาจารย์)
    s1 = new_slide(C_WHITE, show_num=True)
    tb1_dept = s1.shapes.add_textbox(Inches(0.85), Inches(0.40), Inches(11.4), Inches(0.45))
    p_dept = tb1_dept.text_frame.paragraphs[0]
    p_dept.alignment = PP_ALIGN.CENTER
    r_dept = p_dept.add_run()
    r_dept.text = "ภาควิชาวิศวกรรมไอโอทีและสารสนเทศ & ภาควิชาฟิสิกส์ (หลักสูตรสองปริญญา PhysIoT)  •  รหัสโครงงาน: 69_IoT_Project1"
    r_dept.font.name = FONT_THAI
    r_dept.font.size = Pt(18.5)
    r_dept.font.bold = True
    r_dept.font.color.rgb = C_ORANGE_TAG

    tb1 = s1.shapes.add_textbox(Inches(0.85), Inches(0.88), Inches(11.6), Inches(1.35))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "Development of High-Precision Solar Pyranometer"
    r1.font.name = FONT_HEAD
    r1.font.size = Pt(32)
    r1.font.bold = True
    r1.font.color.rgb = C_BLUE_TAG

    p1_sub = tf1.add_paragraph()
    p1_sub.alignment = PP_ALIGN.CENTER
    p1_sub.space_before = Pt(4)
    r1_sub = p1_sub.add_run()
    r1_sub.text = "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (Capstone Project 1)"
    r1_sub.font.name = FONT_THAI
    r1_sub.font.size = Pt(25)
    r1_sub.font.bold = True
    r1_sub.font.color.rgb = C_TEXT_BLACK

    # Evenly spaced hero images without touching
    add_pic_safe(s1, os.path.join(img_dir, "perfboard_top_view.jpg"), 0.90, 2.35, height=3.10)
    add_pic_safe(s1, os.path.join(img_dir, "cad_perfboard_case_assy.png"), 5.35, 2.35, height=3.10)
    add_pic_safe(s1, os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png"), 8.95, 2.35, height=3.10)

    tb1_f = s1.shapes.add_textbox(Inches(0.85), Inches(5.85), Inches(11.6), Inches(1.2))
    tf1_f = tb1_f.text_frame
    tf1_f.word_wrap = True
    pf1 = tf1_f.paragraphs[0]
    pf1.alignment = PP_ALIGN.CENTER
    rf1 = pf1.add_run()
    rf1.text = "นายเตชินท์ รัฐวงษา (66010296)  •  นายณัฐพัชร์ รัตนนวงศ์ (66010256)  •  นายพีรวิช เฉลิมเกียรติ (66010591)"
    rf1.font.name = FONT_THAI
    rf1.font.size = Pt(21)
    rf1.font.bold = True
    rf1.font.color.rgb = C_TEXT_BLACK

    pf2 = tf1_f.add_paragraph()
    pf2.alignment = PP_ALIGN.CENTER
    rf2 = pf2.add_run()
    rf2.text = "อาจารย์ที่ปรึกษา: ผศ.ดร.ธนวิชญ์ อนุวงศ์พินิจ (วิศวกรรมไอโอทีและสารสนเทศ) และ ผศ.ดร.ประธาน บูรณศิริ (ฟิสิกส์อุตสาหกรรม)"
    rf2.font.name = FONT_THAI
    rf2.font.size = Pt(18)
    rf2.font.color.rgb = C_TEXT_MUTED

    # -------------------------------------------------------------------------
    # SLIDE 2: บทที่ 1 ที่มาและความสำคัญ (1) - บทบาทของ Pyranometer ในอุตสาหกรรมโซลาร์เซลล์
    # -------------------------------------------------------------------------
    s2 = new_slide(C_WHITE)
    add_clean_header(
        s2,
        "บทที่ 1 : ที่มาและความสำคัญของปัญหา (1/2)",
        "บทบาทของ Pyranometer ในการประเมินสมรรถนะโรงไฟฟ้าโซลาร์เซลล์"
    )
    add_pic_safe(s2, os.path.join(ext_dir, "p02_img1_450x300.png"), 1.0, 1.80, width=4.8)

    tb2_r = s2.shapes.add_textbox(Inches(6.15), Inches(1.80), Inches(6.3), Inches(3.1))
    tf2_r = tb2_r.text_frame
    tf2_r.word_wrap = True
    tf2_r.vertical_anchor = MSO_ANCHOR.TOP
    p2_1 = tf2_r.paragraphs[0]
    r2_1 = p2_1.add_run()
    r2_1.text = "ไพราโนมิเตอร์ (Solar Pyranometer) คืออะไร?"
    r2_1.font.name = FONT_THAI
    r2_1.font.size = Pt(23)
    r2_1.font.bold = True
    r2_1.font.color.rgb = C_BLUE_TAG

    bullets_s2 = [
        "เป็นเครื่องมือวัดความเข้มรังสีอาทิตย์รวมบนระนาบแนวราบ (Global Horizontal Irradiance: GHI) ในหน่วยวัตต์ต่อตารางเมตร (W/m²)",
        "เป็นตัวหารสำคัญในสมการ Performance Ratio (PR) เพื่อแยกแยะว่าวันที่ผลิตไฟฟ้าได้ลดลง เกิดจากเมฆบัง หรือเกิดจากแผงสกปรก/ระบบขัดข้อง (เช่น PR ตกจาก 82% เหลือ 68%)",
        "ใช้เป็นหลักฐานอ้างอิงตามสัญญาซื้อขายไฟฟ้าและตรวจรับงาน (EPC & PPA)"
    ]
    for btxt in bullets_s2:
        pb = tf2_r.add_paragraph()
        pb.space_before = Pt(6)
        rb = pb.add_run()
        rb.text = "• " + btxt
        rb.font.name = FONT_THAI
        rb.font.size = Pt(19.5)
        rb.font.color.rgb = C_TEXT_BODY

    c2_bot = add_card(s2, 1.0, 5.20, 11.33, 1.75, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf2_b = c2_bot.text_frame
    p2_b1 = tf2_b.paragraphs[0]
    p2_b1.alignment = PP_ALIGN.CENTER
    r2_b1 = p2_b1.add_run()
    r2_b1.text = "Global Horizontal Irradiance (GHI)  &  Performance Ratio (PR)"
    r2_b1.font.name = FONT_HEAD
    r2_b1.font.size = Pt(18)
    r2_b1.font.bold = True
    r2_b1.font.color.rgb = C_BLUE_TAG

    p2_b2 = tf2_b.add_paragraph()
    p2_b2.alignment = PP_ALIGN.CENTER
    p2_b2.space_before = Pt(4)
    r2_b2 = p2_b2.add_run()
    r2_b2.text = "PR  =  E_actual  /  ( G_GHI  ×  A_PV  ×  η_STC )"
    r2_b2.font.name = FONT_HEAD
    r2_b2.font.size = Pt(21)
    r2_b2.font.bold = True
    r2_b2.font.color.rgb = C_TEXT_BLACK

    p2_b3 = tf2_b.add_paragraph()
    p2_b3.alignment = PP_ALIGN.CENTER
    p2_b3.space_before = Pt(4)
    r2_b3 = p2_b3.add_run()
    r2_b3.text = "ทำหน้าที่เป็น “มาตรฐานอ้างอิงในการตรวจวัดความเข้มรังสีอาทิตย์จริง (W/m²)” เพื่อประเมินประสิทธิภาพการผลิตไฟฟ้าอย่างแม่นยำ"
    r2_b3.font.name = FONT_THAI
    r2_b3.font.size = Pt(18.5)
    r2_b3.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 3: บทที่ 1 ที่มาและความสำคัญ (2) - ข้อจำกัดของเครื่องมือวัดในท้องตลาดและทำไมต้องสร้างเอง
    # -------------------------------------------------------------------------
    s3 = new_slide(C_WHITE)
    add_clean_header(
        s3,
        "บทที่ 1 : ที่มาและความสำคัญของปัญหา (2/2)",
        "ข้อจำกัดของเครื่องมือวัดในท้องตลาด และเหตุผลที่เราต้องพัฒนาขึ้นเอง"
    )
    add_pic_safe(s3, os.path.join(ext_dir, "p03_img1_470x470.png"), 1.6, 1.70, height=2.60)
    add_pic_safe(s3, os.path.join(ext_dir, "p03_img0_550x400.png"), 7.5, 1.80, height=2.40)

    c3_l = add_card(s3, 0.95, 4.45, 5.4, 1.45, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf3_l = c3_l.text_frame
    p = tf3_l.paragraphs[0]
    r = p.add_run()
    r.text = "1. แบบความร้อน (Thermopile มาตรฐาน)"
    r.font.name = FONT_THAI
    r.font.size = Pt(21)
    r.font.bold = True
    r.font.color.rgb = C_TEXT_BLACK
    for txt in ["• ราคาสูงมาก (100,000 – 150,000 บาท/เครื่อง)", "• ตอบสนองช้า (5 – 30 วินาที) เมื่อเมฆเคลื่อนผ่านเร็ว"]:
        p2 = tf3_l.add_paragraph()
        r2 = p2.add_run()
        r2.text = txt
        r2.font.name = FONT_THAI
        r2.font.size = Pt(18.5)
        r2.font.color.rgb = C_TEXT_BODY

    c3_r = add_card(s3, 6.95, 4.45, 5.4, 1.45, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf3_r = c3_r.text_frame
    p = tf3_r.paragraphs[0]
    r = p.add_run()
    r.text = "2. แบบสารกึ่งตัวนำ (Silicon Photodiode สำเร็จรูป)"
    r.font.name = FONT_THAI
    r.font.size = Pt(21)
    r.font.bold = True
    r.font.color.rgb = C_TEXT_BLACK
    for txt in ["• ตอบสนองไวระดับเสี้ยววินาที (< 0.01 วินาที)", "• แต่ราคานำเข้ายังสูง (15,000 – 30,000+ บาท/เครื่อง)"]:
        p2 = tf3_r.add_paragraph()
        r2 = p2.add_run()
        r2.text = txt
        r2.font.name = FONT_THAI
        r2.font.size = Pt(18.5)
        r2.font.color.rgb = C_TEXT_BODY

    c3_bot = add_card(s3, 0.95, 6.04, 11.4, 1.06, bg_rgb=C_RED_BG, border_rgb=C_RED_BORDER)
    tf3_b = c3_bot.text_frame
    pb3 = tf3_b.paragraphs[0]
    pb3.alignment = PP_ALIGN.CENTER
    rb3 = pb3.add_run()
    rb3.text = (
        "ช่องว่างสำคัญ: เฉพาะ “หัววัด (Probe)” ในตลาดมีราคาสูง 15,000–150,000 บาท และยังต้องซื้อ Data Logger แยกอีกหลายหมื่นบาท\n"
        "➔ โครงงานนี้จึงพัฒนา “หัววัด 24 บิต” ในงบ < 3,000 บาท ส่งค่า Raw ได้ทั้ง Wi-Fi / RS-485 เข้าสู่ Data Logger และ Web ที่พัฒนาขึ้นเอง"
    )
    rb3.font.name = FONT_THAI
    rb3.font.size = Pt(17.5)
    rb3.font.bold = True
    rb3.font.color.rgb = C_RED_TEXT

    # -------------------------------------------------------------------------
    # SLIDE 4: บทที่ 1 แหล่งอ้างอิงมาตรฐานสากล ISO 9060:2018 (1/2) — หลักฐานจากเอกสารฉบับจริง
    # -------------------------------------------------------------------------
    s4_iso1 = new_slide(C_WHITE)
    add_clean_header(
        s4_iso1,
        "บทที่ 1 : แหล่งอ้างอิงมาตรฐานสากล ISO 9060:2018 (1/2)",
        "หลักฐานอ้างอิงจากเอกสารมาตรฐานสากล ISO 9060:2018 ฉบับจริง",
        "ยืนยันการรับรองหัววัดชนิดโฟโตไดโอดซิลิคอน (Si-pyranometer) และการชดเชยความคลาดเคลื่อนด้วยหน่วยประมวลผลภายในเครื่อง"
    )
    add_pic_safe(s4_iso1, os.path.join(iso_dir, "iso_cover.png"), 0.75, 1.92, height=5.00)
    add_pic_safe(s4_iso1, os.path.join(iso_dir, "iso_p10_types_4_2.png"), 4.35, 1.92, width=4.85)
    add_pic_safe(s4_iso1, os.path.join(iso_dir, "iso_p14_fast_and_note.png"), 4.35, 3.55, width=4.85)

    c4_iso_note = add_card(s4_iso1, 9.40, 1.92, 3.20, 5.05, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG, border_pt=1.5)
    tf4_in = c4_iso_note.text_frame
    p = tf4_in.paragraphs[0]
    r = p.add_run()
    r.text = "สาระสำคัญจากมาตรฐาน:"
    r.font.name = FONT_THAI
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG

    iso_pts = [
        "1) หัวข้อ 4.2 Types (หน้า 4):\nรับรองไพราโนมิเตอร์ชนิดสารกึ่งตัวนำซิลิคอน (Photoelectric / Si-pyranometer) ควบคู่กับแบบ Thermopile อย่างเป็นทางการ",
        "2) หัวข้อ 4.3.3 (หน้า 8):\nกำหนดหมวดหมู่ Fast Response สำหรับหัววัดที่ตอบสนองเร็วกว่า < 0.5 วินาที (โฟโตไดโอดของเราทำได้ < 0.01 วินาที)",
        "3) กรอบ NOTE (หน้า 8):\nอนุญาตให้นำสัญญาณที่ผ่านสมการหักลบออฟเซ็ตและชดเชยความคลาดเคลื่อนในหน่วยประมวลผล / Data Logger ของระบบ มานับเกรดมาตรฐานได้โดยตรง"
    ]
    for ipt in iso_pts:
        pi = tf4_in.add_paragraph()
        pi.space_before = Pt(6)
        ri = pi.add_run()
        ri.text = ipt
        ri.font.name = FONT_THAI
        ri.font.size = Pt(16.5)
        ri.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 5: บทที่ 1 แหล่งอ้างอิงมาตรฐานสากล ISO 9060:2018 (2/2) — ตาราง Table 1 ฉบับจริง
    # -------------------------------------------------------------------------
    s4_iso2 = new_slide(C_WHITE)
    add_clean_header(
        s4_iso2,
        "บทที่ 1 : แหล่งอ้างอิงมาตรฐานสากล ISO 9060:2018 (2/2)",
        "เกณฑ์การจัดระดับไพราโนมิเตอร์จากตารางที่ 1 (Table 1) ในเอกสาร ISO 9060:2018 หน้า 6–7"
    )
    add_pic_safe(s4_iso2, os.path.join(iso_dir, "iso_p12_table1_part1.png"), 0.85, 1.68, height=5.38)
    add_pic_safe(s4_iso2, os.path.join(iso_dir, "iso_p13_table1_part2.png"), 6.85, 1.68, height=3.55)

    c4_iso_t = add_card(s4_iso2, 6.65, 5.28, 5.85, 1.80, bg_rgb=C_AMBER_BG, border_rgb=C_ORANGE_TAG)
    tf4_it = c4_iso_t.text_frame
    p = tf4_it.paragraphs[0]
    r = p.add_run()
    r.text = "พารามิเตอร์หลักที่โครงงานนี้นำมาเป็นเกณฑ์ออกแบบ (Table 1):"
    r.font.name = FONT_THAI
    r.font.size = Pt(17)
    r.font.bold = True
    r.font.color.rgb = C_ORANGE_TAG
    p2 = tf4_it.add_paragraph()
    p2.space_before = Pt(2)
    r2 = p2.add_run()
    r2.text = (
        "• ข้อ a (Response time): Class A < 10 s / Fast response < 0.5 s\n"
        "• ข้อ c2 (Nonlinearity): Class A ±0.5%, Class B ±1%, Class C ±3%\n"
        "• ข้อ c3 (Directional error): Class A ±10, Class B ±20, Class C ±30 W/m²\n"
        "• ข้อ c5 & c7 (Temp. & Processing): Class A ±1% และ ±2 W/m²"
    )
    r2.font.name = FONT_THAI
    r2.font.size = Pt(15.5)
    r2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 6: บทที่ 1 วัตถุประสงค์และเป้าหมายสมรรถนะตามมาตรฐาน ISO 9060:2018 (อัปเดตตรง Table 1 100%)
    # -------------------------------------------------------------------------
    s4 = new_slide(C_WHITE)
    add_clean_header(
        s4,
        "บทที่ 1 : วัตถุประสงค์และเป้าหมายของโครงงาน",
        "ตารางสรุปเป้าหมายสมรรถนะของโครงงานเทียบเกณฑ์ ISO 9060:2018 (Table 1)"
    )
    rows_s4 = [
        ["พารามิเตอร์สมรรถนะ (ISO 9060 Table 1)", "Class A (ระดับสูงสุด)", "Class B", "Class C", "แนวทางของโครงงานนี้ (งบ < 3,000 บาท)"],
        ["เวลาการตอบสนอง 95% (ข้อ a: Response Time)", "< 10 s (< 0.5 s Fast)", "< 20 s", "< 30 s", "< 0.01 s (โฟโตไดโอดตอบสนองระดับ Fast Response)"],
        ["ความไม่เป็นเชิงเส้น (ข้อ c2: Non-linearity)", "< ±0.5%", "< ±1.0%", "< ±3.0%", "< ±0.5% (คง V_bias = 0.418 V ป้องกันอิ่มตัว)"],
        ["ความคลาดเคลื่อนเชิงมุม (ข้อ c3: Directional Error)", "< ±10 W/m²", "< ±20 W/m²", "< ±30 W/m²", "ชดเชยด้วยพิกัดติดตั้ง/เวลา (มุมดวงอาทิตย์) + ผลแล็บฟิสิกส์"],
        ["ความคลาดเคลื่อนจากอุณหภูมิ (ข้อ c5: Temp. Response)", "< ±1%", "< ±2%", "< ±4%", "< ±1% (ชดเชยด้วย NTC + สมการใน Data Logger)"],
        ["ความคลาดเคลื่อนภาคประมวลผล (ข้อ c7: Processing Error)", "< ±2 W/m²", "< ±5 W/m²", "< ±10 W/m²", "< ±0.01 W/m² (ADS1220 24-bit ADC + REF5045)"],
        ["พิกัดวัดแสงสูงสุด & ช่วงคลื่น (WMO-No. 8 / Si)", "4,000 W/m² (300–2800nm)", "2,000 W/m²", "1,500 W/m² (400–1100nm)", "ปรับได้ถึง 4,000 W/m² (R1 = 390 Ohm, 400–1100nm)"]
    ]
    tab4 = s4.shapes.add_table(len(rows_s4), 5, Inches(0.85), Inches(1.72), Inches(11.63), Inches(4.20)).table
    col_w4 = [3.35, 1.95, 1.45, 1.75, 3.13]
    for j, w in enumerate(col_w4):
        tab4.columns[j].width = Inches(w)

    for i, row in enumerate(rows_s4):
        for j, val in enumerate(row):
            cell = tab4.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if i == 0:
                cell.fill.fore_color.rgb = C_NAVY_HEADER
            elif j == 4:
                cell.fill.fore_color.rgb = C_BLUE_BG
            elif (i in [1, 2, 4, 5, 6] and j == 1) or (i == 6 and j == 3):
                cell.fill.fore_color.rgb = C_AMBER_BG
            elif i % 2 == 0:
                cell.fill.fore_color.rgb = C_ROW_ALT
            else:
                cell.fill.fore_color.rgb = C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(17 if i == 0 else 16)
            r.font.bold = (i == 0 or j == 4)
            r.font.color.rgb = C_WHITE if i == 0 else (C_BLUE_TAG if j == 4 else C_TEXT_BLACK)

    c4_ref = add_card(s4, 0.85, 6.22, 11.63, 0.85, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf4_r = c4_ref.text_frame
    p4_r = tf4_r.paragraphs[0]
    r4_r = p4_r.add_run()
    r4_r.text = (
        "แหล่งอ้างอิงมาตรฐาน (Standard Reference): ISO 9060:2018(E) Table 1 (Pages 6–8) “Solar energy — Specification and classification of instruments for measuring "
        "hemispherical solar and direct solar radiation” และเอกสาร WMO-No. 8 (World Meteorological Organization Guide, Chapter 7)"
    )
    r4_r.font.name = FONT_THAI
    r4_r.font.size = Pt(16.5)
    r4_r.font.bold = True
    r4_r.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 5: บทที่ 1 ภาพรวมของระบบ (System Overview Architecture - พร้อมรูปประกอบ & สถานะเหมือนของเพื่อน)
    # -------------------------------------------------------------------------
    s5 = new_slide(C_WHITE)
    add_clean_header(
        s5,
        "บทที่ 1 : ภาพรวมของระบบ (System Overview)",
        "สถาปัตยกรรมภาพรวมของระบบ Smart IoT Pyranometer",
        "บูรณาการการออกแบบทางทัศนศาสตร์ วงจรแอนะล็อกความละเอียดสูง 24 บิต และระบบประมวลผลขึ้นสู่ Web Dashboard"
    )
    overview_blocks = [
        (
            "1. ส่วนหัววัดทางทัศนศาสตร์\n(Optical Sensor Head)",
            os.path.join(img_dir, "cad_perfboard_case_assy.png"),
            [
                "• กระจกแซฟไฟร์ (Al2O3, 9 Mohs)",
                "• โฟโตไดโอดตัวถังแบน BPW 34",
                "• เทอร์มิสเตอร์ NTC 10 kOhm",
                "• เคส 3D CAD ปิดผนึกกันความชื้น"
            ],
            "ทำเสร็จแล้ว • โครงงาน 1",
            C_BLUE_BG, C_BLUE_TAG, C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "2. วงจรแอนะล็อก 24 บิต\n(24-Bit Analog Front-End)",
            os.path.join(img_dir, "perfboard_bottom_modules.jpg"),
            [
                "• วงจร OPA380 TIA (V_bias 0.418V)",
                "• ปรับ R1 รองรับพิกัด 4,000 W/m²",
                "• ชิป ADS1220 24-Bit ADC (4 ช่อง)",
                "• ไฟเลี้ยง TPS7A4901 & REF5045"
            ],
            "ทำเสร็จแล้ว • โครงงาน 1",
            C_GREEN_BG, C_GREEN_TEXT, C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "3. ภาคส่งข้อมูลค่าดิบ\n(Wi-Fi & RS-485 Modbus)",
            os.path.join(web_dir, "slide06_pic03.png"),
            [
                "• ESP32-C6 อ่าน ADC 24 บิต (SPI)",
                "• ส่งค่า Raw ครบทั้ง 4 ช่องสัญญาณ",
                "• ส่งไร้สายผ่าน Wi-Fi (MQTT)",
                "• หรือส่งผ่านสาย RS-485 (Modbus)"
            ],
            "อ่านค่าได้แล้ว • เชื่อมต่อในโครงงาน 2",
            C_AMBER_BG, C_ORANGE_TAG, C_AMBER_BG, C_ORANGE_TAG
        ),
        (
            "4. Data Logger & หน้าเว็บ\n(Logger & Web Dashboard)",
            os.path.join(web_dir, "slide15_pic01.jpg"),
            [
                "• บันทึกค่า Raw & เวลา (พิกัดติดตั้ง)",
                "• ชดเชยมุมด้วยผลแล็บฟิสิกส์ (0°–85°)",
                "• ชดเชยอุณหภูมิ NTC & ดึงไฟ AC",
                "• แสดงผลบน Web & คำนวณค่า PR"
            ],
            "มี Mockup แล้ว • ทำจริงในโครงงาน 2",
            C_CARD_GRAY, C_TEXT_BLACK, C_AMBER_BG, C_ORANGE_TAG
        )
    ]
    for idx, (btitle, bimg, bitems, bstatus, bg_c, tc, s_bg, s_tc) in enumerate(overview_blocks):
        left = 0.85 + idx * 2.98
        c = add_card(s5, left, 1.92, 2.75, 4.48, bg_rgb=bg_c, border_rgb=tc, border_pt=1.5)
        tf = c.text_frame
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = btitle
        r.font.name = FONT_THAI
        r.font.size = Pt(18.5)
        r.font.bold = True
        r.font.color.rgb = tc

        # Image frame inside the card
        img_box = add_card(s5, left + 0.18, 2.68, 2.39, 1.48, bg_rgb=C_WHITE, border_rgb=C_BORDER_GRAY)
        if idx == 2:
            add_pic_safe(s5, bimg, left + 0.90, 2.92, height=1.00)
        elif idx == 3:
            add_pic_safe(s5, bimg, left + 0.24, 2.76, width=2.27)
        else:
            add_pic_safe(s5, bimg, left + 0.42, 2.73, height=1.38)

        tb_list = s5.shapes.add_textbox(Inches(left + 0.12), Inches(4.22), Inches(2.52), Inches(2.10))
        tf_l = tb_list.text_frame
        tf_l.word_wrap = True
        for i_idx, item in enumerate(bitems):
            pi = tf_l.paragraphs[0] if i_idx == 0 else tf_l.add_paragraph()
            pi.space_before = Pt(2)
            ri = pi.add_run()
            ri.text = item
            ri.font.name = FONT_THAI
            ri.font.size = Pt(16.0)
            ri.font.color.rgb = C_TEXT_BLACK

        # Status badge at bottom
        c_stat = add_card(s5, left, 6.50, 2.75, 0.48, bg_rgb=s_bg, border_rgb=s_tc, border_pt=1.2)
        tf_s = c_stat.text_frame
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf_s.margin_top = Inches(0.04)
        tf_s.margin_bottom = Inches(0.04)
        tf_s.margin_left = Inches(0.05)
        tf_s.margin_right = Inches(0.05)
        ps = tf_s.paragraphs[0]
        ps.alignment = PP_ALIGN.CENTER
        rs = ps.add_run()
        rs.text = bstatus
        rs.font.name = FONT_THAI
        rs.font.size = Pt(15.5)
        rs.font.bold = True
        rs.font.color.rgb = s_tc

        if idx < 3:
            arr = s5.shapes.add_textbox(Inches(left + 2.74), Inches(4.05), Inches(0.25), Inches(0.5))
            pa = arr.text_frame.paragraphs[0]
            pa.alignment = PP_ALIGN.CENTER
            ra = pa.add_run()
            ra.text = "➔"
            ra.font.name = FONT_HEAD
            ra.font.size = Pt(18)
            ra.font.bold = True
            ra.font.color.rgb = C_BLUE_TAG

    # =========================================================================
    # บทที่ 2: ทฤษฎีและความท้าทายทางวิศวกรรม 3 ด้านในการสร้างหัววัดซิลิคอน (Slides 6 - 8)
    # =========================================================================

    # SLIDE 6: ความท้าทายที่ 1 (ด้านแสง/มุมตกกระทบ - Cosine Error)
    s6 = new_slide(C_WHITE)
    add_clean_header(
        s6,
        "บทที่ 2 : ทฤษฎีและความท้าทายในการพัฒนาหัววัดซิลิคอน (1/3)",
        "ความท้าทายที่ 1 : ด้านทัศนศาสตร์และมุมตกกระทบตามกฎโคไซน์ (Cosine Response)",
        "ในการนำเซนเซอร์ซิลิคอนเปล่ามาสร้างเครื่องวัดเอง ต้องออกแบบไม่ให้ขอบตัวถังบังแสงเมื่อดวงอาทิตย์ทำมุมเอียง (0° – 85°)"
    )
    add_pic_safe(s6, os.path.join(ext_dir, "p12_img0_504x426.png"), 1.05, 1.95, height=4.95)
    add_pic_safe(s6, os.path.join(ext_dir, "p12_img1_822x796.png"), 6.45, 1.95, height=4.95)

    # SLIDE 7: ความท้าทายที่ 2 (ด้านสัญญาณไฟฟ้า - Current to Voltage Conversion)
    s7 = new_slide(C_WHITE)
    add_clean_header(
        s7,
        "บทที่ 2 : ทฤษฎีและความท้าทายในการพัฒนาหัววัดซิลิคอน (2/3)",
        "ความท้าทายที่ 2 : ด้านการแปลงกระแสแสงเป็นแรงดันเชิงเส้น (Current-to-Voltage)",
        "การต่อตัวต้านทานวัดแรงดันโดยตรงจะทำให้ไดโอดอิ่มตัวเมื่อแรงดันเข้าใกล้ V_OC (365 mV) จึงต้องใช้วงจร OPA380 TIA"
    )
    add_pic_safe(s7, os.path.join(ext_dir, "p13_img0_516x547.png"), 1.15, 2.05, height=4.55)
    # Clean vector red 'X' badge instead of black-background bitmap
    x_shape = s7.shapes.add_shape(MSO_SHAPE.MATH_MULTIPLY, Inches(3.20), Inches(2.35), Inches(1.05), Inches(1.05))
    x_shape.fill.solid()
    x_shape.fill.fore_color.rgb = RGBColor(220, 38, 38)
    x_shape.line.color.rgb = RGBColor(153, 27, 27)
    x_shape.line.width = Pt(1.5)
    add_pic_safe(s7, os.path.join(ext_dir, "p13_img1_879x466.png"), 5.75, 2.15, width=6.7)

    # SLIDE 8: ความท้าทายที่ 3 (ด้านอุณหภูมิ - Thermal Drift ของซิลิคอน)
    s8 = new_slide(C_WHITE)
    add_clean_header(
        s8,
        "บทที่ 2 : ทฤษฎีและความท้าทายในการพัฒนาหัววัดซิลิคอน (3/3)",
        "ความท้าทายที่ 3 : ด้านผลกระทบจากอุณหภูมิต่อสารกึ่งตัวนำซิลิคอน (Thermal Drift)",
        "เมื่อซิลิคอนร้อนขึ้น ช่องว่างพลังงาน (Bandgap) จะแคบลง ทำให้กระแสลัดวงจรเพิ่มขึ้นตามสัมประสิทธิ์ TC_I = +0.18 %/K"
    )
    add_pic_safe(s8, os.path.join(ext_dir, "p14_img1_706x72.png"), 1.0, 1.90, width=3.8)
    add_pic_safe(s8, os.path.join(ext_dir, "p14_img0_227x310.png"), 1.25, 2.45, height=2.65)
    add_pic_safe(s8, os.path.join(ext_dir, "p16_img0_703x261.png"), 5.2, 2.0, width=7.2)

    c8_bot = add_card(s8, 0.95, 5.30, 11.45, 1.65, bg_rgb=C_AMBER_BG, border_rgb=C_ORANGE_TAG, border_pt=1.5)
    tf8 = c8_bot.text_frame
    p8_1 = tf8.paragraphs[0]
    p8_1.alignment = PP_ALIGN.CENTER
    r8_1 = p8_1.add_run()
    r8_1.text = "ผลกระทบทางคณิตศาสตร์จากสเปกชีต BPW 34 (TC_I = +0.18 %/K หรือ +0.18 %/°C):"
    r8_1.font.name = FONT_THAI
    r8_1.font.size = Pt(20.5)
    r8_1.font.bold = True
    r8_1.font.color.rgb = C_ORANGE_TAG

    p8_2 = tf8.add_paragraph()
    p8_2.alignment = PP_ALIGN.CENTER
    p8_2.space_before = Pt(4)
    r8_2 = p8_2.add_run()
    r8_2.text = (
        "ที่ 25°C กระแสอยู่ที่ 1.00 เท่า แต่เมื่อตากแดดจนหัววัดร้อนขึ้นเป็น 65°C (อุณหภูมิเพิ่ม +40°C) กระแสจะพุ่งขึ้นเป็น 1.072 เท่า (เพี้ยน +7.2%!)\n"
        "ตัวอย่างเช่น แดดจริง 1,000 W/m² เครื่องจะอ่านหลอกว่าแดดแรง 1,072 W/m² ➔ จึงต้องติด NTC วัดอุณหภูมิจริงเพื่อหารลดทอนค่าที่เฟ้อออก"
    )
    r8_2.font.name = FONT_THAI
    r8_2.font.size = Pt(19.5)
    r8_2.font.bold = True
    r8_2.font.color.rgb = C_TEXT_BLACK

    # =========================================================================
    # บทที่ 3: การออกแบบและวิธีดำเนินงาน (Slides 9 - 18)
    # =========================================================================

    # SLIDE 9: ภาพรวมขั้นตอนการออกแบบและวิธีดำเนินงาน 6 ขั้นตอนหลัก (แยกชัดเจน ไม่ยัดรวมกัน)
    s9 = new_slide(C_WHITE)
    add_clean_header(
        s9,
        "บทที่ 3 : การออกแบบและวิธีดำเนินงาน (Methodology Overview)",
        "ภาพรวมขั้นตอนการออกแบบและพัฒนาในโครงงาน (6 ขั้นตอนหลัก)",
        "แยกลำดับงานตั้งแต่การจำลองแสง การสร้างฮาร์ดแวร์ การเก็บกราฟจริงในแล็บฟิสิกส์ ไปจนถึงการชดเชยค่าใน Data Logger"
    )
    method_steps = [
        (
            "ขั้นตอนที่ 1 : คัดเลือกหัววัด & จำลองแสง 3 มิติ",
            "เปรียบเทียบ BPX 61 กับ BPW 34 ด้วยโมเดล 3D CAD (FreeCAD) และจำลอง Ray-Tracing ด้วยเอ็นจิน Mitsuba 3 เพื่อวิเคราะห์มุมรับแสง",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "ขั้นตอนที่ 2 : ออกแบบเคสปิดผนึก & กระจกแซฟไฟร์",
            "วิเคราะห์การเสื่อมสภาพของเรซินและขาโลหะเงิน แล้วออกแบบเคสปิดผนึกพร้อมฝาครอบกระจกแซฟไฟร์สังเคราะห์ (Al2O3 ความแข็ง 9 Mohs)",
            C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "ขั้นตอนที่ 3 : ออกแบบวงจร TIA 24 บิต & ภาคจ่ายไฟ",
            "ออกแบบวงจร OPA380 TIA (V_bias = 0.418 V) รองรับพิกัด 4,000 W/m² คู่กับ ADS1220 24-bit ADC, NTC และไฟอ้างอิง TPS7A4901 / REF5045",
            C_AMBER_BG, C_ORANGE_TAG
        ),
        (
            "ขั้นตอนที่ 4 : สร้างบอร์ดต้นแบบจริง & 4-Layer PCB",
            "บัดกรีและดีบักบอร์ด Perfboard จริง อ่านค่า 24 บิตครบ 4 ช่องสัญญาณ และออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนานพร้อมกล่องเคส 3D",
            C_CARD_GRAY, C_TEXT_BLACK
        ),
        (
            "ขั้นตอนที่ 5 : เก็บกราฟ Cosine Error จริงในแล็บฟิสิกส์",
            "วัดสเปกตรัมการส่องผ่านของกระจกแซฟไฟร์ และหมุนวัดค่าที่หัววัดอ่านได้จริงตามมุมตกกระทบ 0°–85° บนแท่นหมุน Optical Rotation Stage",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "ขั้นตอนที่ 6 : พัฒนา Data Logger & ชดเชยค่าบน Web",
            "ส่งค่า Raw ทั้ง 4 ช่อง (Wi-Fi / RS-485) เข้า Data Logger เพื่อนำกราฟแล็บฟิสิกส์มาชดเชยร่วมกับพิกัดโซลาร์เซลล์ เวลาจริง และอุณหภูมิ",
            C_GREEN_BG, C_GREEN_TEXT
        )
    ]
    for idx, (st_title, st_desc, bg_c, tc) in enumerate(method_steps):
        col = idx % 3
        row = idx // 3
        left = 0.85 + col * 3.94
        top = 1.92 + row * 2.55
        c = add_card(s9, left, top, 3.75, 2.32, bg_rgb=bg_c, border_rgb=tc, border_pt=1.5)
        tf = c.text_frame
        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = st_title
        r0.font.name = FONT_THAI
        r0.font.size = Pt(19.5)
        r0.font.bold = True
        r0.font.color.rgb = tc

        p2 = tf.add_paragraph()
        p2.space_before = Pt(6)
        r2 = p2.add_run()
        r2.text = st_desc
        r2.font.name = FONT_THAI
        r2.font.size = Pt(17.5)
        r2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 10: ขั้นตอนที่ 1.1 การคัดเลือกเซนเซอร์ (+ ตอบคำถามอาจารย์: ทำไมเลือกตอนแรก & ทำไมถึงเปลี่ยน)
    # -------------------------------------------------------------------------
    s10 = new_slide(C_WHITE)
    add_clean_header(
        s10,
        "บทที่ 3 : ขั้นตอนที่ 1 — การคัดเลือกเซนเซอร์รับแสง",
        "เปรียบเทียบโครงสร้างตัวถังและเหตุผลในการเปลี่ยนจาก BPX 61 เป็น BPW 34"
    )
    add_pic_safe(s10, os.path.join(img_dir, "(Directional Polar Plot) ของ BPX 61 .png"), 0.85, 1.70, height=2.10)
    add_pic_safe(s10, os.path.join(img_dir, "dimention(BPX61).png"), 3.85, 1.70, height=2.10)
    add_pic_safe(s10, os.path.join(img_dir, "(Directional Polar Plot) ของ BPW 34.png"), 7.05, 1.70, height=2.10)
    add_pic_safe(s10, os.path.join(img_dir, "dimention (BPW34).png"), 9.95, 1.70, height=2.10)

    rows_s10 = [
        ["พารามิเตอร์ทางเทคนิค", "ams-OSRAM BPX 61 (รุ่นที่ศึกษาในระยะเริ่มต้น)", "ams-OSRAM BPW 34 (รุ่นที่เลือกใช้เป็นหลัก)"],
        ["ขนาดชิปและพื้นที่รับแสงกัมมันต์", "2.65 × 2.65 mm (7.02 mm²)", "2.65 × 2.65 mm (7.02 mm² — ชิปขนาดเดียวกัน)"],
        ["ช่วงความยาวคลื่น & สัมประสิทธิ์อุณหภูมิ", "400 – 1,100 nm  •  TC_I = +0.18 %/K", "400 – 1,100 nm  •  TC_I = +0.18 %/K"],
        ["โครงสร้างตัวถังและมุมรับแสง", "กระป๋องโลหะ TO-39 ลึก 1.55 mm (ขอบบังแสงมุมเอียง)", "ตัวถังแบนราบ DIL ชิปอยู่ชิดผิวหน้า (รับแสงมุมกว้างตามกฎโคไซน์)"]
    ]
    tab10 = s10.shapes.add_table(len(rows_s10), 3, Inches(0.85), Inches(3.92), Inches(11.63), Inches(1.90)).table
    tab10.columns[0].width = Inches(3.23)
    tab10.columns[1].width = Inches(4.2)
    tab10.columns[2].width = Inches(4.2)
    for i, row in enumerate(rows_s10):
        for j, val in enumerate(row):
            cell = tab10.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_ROW_ALT if i % 2 == 0 else C_WHITE)
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(17.5 if i == 0 else 16.5)
            r.font.bold = (i == 0 or (i == 3 and j in [1, 2]))
            r.font.color.rgb = C_WHITE if i == 0 else (C_RED_TEXT if (i == 3 and j == 1) else (C_BLUE_TAG if (i == 3 and j == 2) else C_TEXT_BLACK))

    c10_why1 = add_card(s10, 0.85, 6.02, 5.70, 1.10, bg_rgb=C_RED_BG, border_rgb=C_RED_BORDER)
    tf10_w1 = c10_why1.text_frame
    p = tf10_w1.paragraphs[0]
    r = p.add_run()
    r.text = "ทำไมตอนแรกถึงเลือก BPX 61? : "
    r.font.name = FONT_THAI
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_RED_TEXT
    r2 = p.add_run()
    r2.text = "เพราะบรรจุในตัวถังโลหะปิดผนึก (Hermetic Metal Can TO-39) พร้อมหน้าต่างกระจกในตัว กันความชื้นได้ดีมากสำหรับการติดตั้งกลางแจ้ง"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(17.5)
    r2.font.color.rgb = C_TEXT_BLACK

    c10_why2 = add_card(s10, 6.78, 6.02, 5.70, 1.10, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tf10_w2 = c10_why2.text_frame
    p = tf10_w2.paragraphs[0]
    r = p.add_run()
    r.text = "ทำไมจึงเปลี่ยนมาใช้ BPW 34? : "
    r.font.name = FONT_THAI
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_TEXT
    r2 = p.add_run()
    r2.text = "เพราะแผ่นชิปของ BPX 61 จมลึกจากขอบกระป๋อง 1.55 mm ทำให้บังแสงมุมเอียง จึงเปลี่ยนเป็น BPW 34 ที่ชิปชิดผิวหน้าแล้วใส่เคสกระจกแซฟไฟร์แทน"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(17.5)
    r2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 11: ขั้นตอนที่ 1.2 การสร้างแบบจำลอง 3D CAD ด้วยโปรแกรม FreeCAD
    # -------------------------------------------------------------------------
    s11 = new_slide(C_WHITE)
    add_clean_header(
        s11,
        "บทที่ 3 : ขั้นตอนที่ 1 — การสร้างแบบจำลอง 3 มิติ (3D CAD Modeling)",
        "การออกแบบโมเดล 3D CAD ของเซนเซอร์ BPX 61 (TO-39) และ BPW 34 (Flat DIL)",
        "สร้างด้วยซอฟต์แวร์ FreeCAD ตามสัดส่วนจริงทุกชิ้นส่วน (Active Die 2.65×2.65 mm², ตัวถัง, ขาโลหะ และกระจก) เพื่อส่งเข้าจำลอง Ray-Tracing"
    )
    add_pic_safe(s11, os.path.join(img_dir, "CAD1.png"), 0.85, 2.00, width=3.70)
    add_pic_safe(s11, os.path.join(img_dir, "CAD2.png"), 4.82, 2.00, width=3.70)
    add_pic_safe(s11, os.path.join(img_dir, "cad_bpw34_model.png"), 8.80, 2.00, width=3.65)

    captions_s11 = [
        (0.85, 3.70, "(ก) โมเดล 3D CAD ของ BPX 61 (TO-39)"),
        (4.82, 3.70, "(ข) ภาพตัดขวางความลึกกระป๋อง 1.55–1.75 mm"),
        (8.80, 3.65, "(ค) โมเดล 3D CAD ของ BPW 34 (Flat DIL)")
    ]
    for left, w, text in captions_s11:
        tb = s11.shapes.add_textbox(Inches(left), Inches(5.90), Inches(w), Inches(0.8))
        tb.text_frame.word_wrap = True
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.name = FONT_THAI
        r.font.size = Pt(18.5)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 12: ขั้นตอนที่ 2.1 ทำไมไม่ใช้ BPW 34 เปล่าๆ ตากแดด? (เหตุผลที่ 1: เรซิน Epoxy เสื่อมสภาพ)
    # -------------------------------------------------------------------------
    s12 = new_slide(C_WHITE)
    add_clean_header(
        s12,
        "บทที่ 3 : ขั้นตอนที่ 2 — การวิเคราะห์ความทนทานและการปกป้องหัววัด (1/2)",
        "ทำไมจึงไม่ใช้ BPW 34 เปล่าๆ ตากแดดโดยตรง? (ด้านวัสดุพอลิเมอร์เคลือบผิว)",
        "แม้ BPW 34 ตัวเปล่าจะให้กราฟมุมรับแสงใกล้เคียง Ideal Cosine แต่ตัวถังเรซิน Epoxy ไม่สามารถทนสภาพแวดล้อมกลางแจ้งได้"
    )
    add_pic_safe(s12, os.path.join(ext_dir, "p22_img0_250x498.png"), 1.35, 1.95, height=4.95)

    c12_r1 = add_card(s12, 4.45, 2.00, 8.0, 2.25, bg_rgb=C_RED_BG, border_rgb=C_RED_BORDER)
    tf12_1 = c12_r1.text_frame
    p = tf12_1.paragraphs[0]
    r = p.add_run()
    r.text = "1. การขูดขีดจากฝุ่นทรายควอตซ์ (Mechanical Abrasion)"
    r.font.name = FONT_THAI
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = C_RED_TEXT
    p2 = tf12_1.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = (
        "ผิวพลาสติก Epoxy ของ BPW 34 มีความแข็งต่ำมาก (เพียง 2–3 Mohs) ในขณะที่ฝุ่นทรายในอากาศคือผลึกควอตซ์ (SiO2 ความแข็ง 7 Mohs) "
        "เมื่อลมพัดฝุ่นทรายมาเสียดสี ผิวพลาสติกจะถูกขูดจนกลายเป็นฝ้าขาวและสูญเสียสารเคลือบกัน UV"
    )
    r2.font.name = FONT_THAI
    r2.font.size = Pt(19.5)
    r2.font.color.rgb = C_TEXT_BLACK

    c12_r2 = add_card(s12, 4.45, 4.50, 8.0, 2.40, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf12_2 = c12_r2.text_frame
    p = tf12_2.paragraphs[0]
    r = p.add_run()
    r.text = "2. ปฏิกิริยาออกซิเดชันด้วยแสง (Photo-oxidation of Polymers)"
    r.font.name = FONT_THAI
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = C_TEXT_BLACK
    p2 = tf12_2.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = (
        "การทำงานร่วมกันของรังสี UV ในแสงแดดและแก๊สออกซิเจนในอากาศ ทำให้สายโซ่พอลิเมอร์แตกหัก (Chain Scission) "
        "ส่งผลให้เรซินใสเปลี่ยนเป็นสีเหลือง เปราะแตกร้าว (ภาพกล้องจุลทรรศน์ SEM ด้านซ้าย) และทำให้ค่าการส่องผ่านแสงเพี้ยนไปอย่างถาวร"
    )
    r2.font.name = FONT_THAI
    r2.font.size = Pt(19.5)
    r2.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 13: ขั้นตอนที่ 2.2 การเกิดสนิมที่ขาเงิน & การป้องกันด้วยกล่องเคสปิดผนึก + กระจกแซฟไฟร์
    # -------------------------------------------------------------------------
    s13 = new_slide(C_WHITE)
    add_clean_header(
        s13,
        "บทที่ 3 : ขั้นตอนที่ 2 — การวิเคราะห์ความทนทานและการปกป้องหัววัด (2/2)",
        "การป้องกันการกัดกร่อนของขาโลหะเงินด้วยเคสปิดผนึกและกระจกแซฟไฟร์"
    )
    c13_top = add_card(s13, 0.85, 1.75, 6.95, 2.25, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf13 = c13_top.text_frame
    p = tf13.paragraphs[0]
    r = p.add_run()
    r.text = "จุดอ่อนที่ผู้ผลิตระบุ: ขาโลหะมีส่วนผสมของเงิน (Silver Leads)"
    r.font.name = FONT_THAI
    r.font.size = Pt(20.5)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG
    p2 = tf13.add_paragraph()
    p2.space_before = Pt(4)
    r2 = p2.add_run()
    r2.text = (
        "• ขาโลหะของ BPW 34 มีส่วนผสมของเงินซึ่งไวต่อความชื้นและแก๊สซัลเฟอร์ในอากาศ หากสัมผัสความชื้นจะเกิดสนิมและสร้างกระแสรั่วไหล (Leakage Current) กวนสัญญาณระดับไมโครแอมป์ทันที\n"
        "• ทางออกของเรา: ออกแบบ “กล่องเคสปิดผนึก + ฝาครอบกระจกแซฟไฟร์สังเคราะห์ (Al2O3 ความแข็ง 9 Mohs)” ตัดวงจรความชื้น ฝุ่นทราย และออกซิเจนอย่างสมบูรณ์"
    )
    r2.font.name = FONT_THAI
    r2.font.size = Pt(17.5)
    r2.font.color.rgb = C_TEXT_BLACK

    add_pic_safe(s13, os.path.join(ext_dir, "p23_img1_816x480.png"), 0.95, 4.12, height=3.05)
    add_pic_safe(s13, os.path.join(img_dir, "cad_sapphire_retaining_ring.png"), 8.20, 1.75, width=3.8)
    add_pic_safe(s13, os.path.join(img_dir, "cad_bpw34_model.png"), 8.20, 4.30, height=2.85)

    # -------------------------------------------------------------------------
    # SLIDE 14: ขั้นตอนที่ 3.1 หลักการทำงานของวงจร OPA380 TIA และ Vbias = 0.418 V
    # -------------------------------------------------------------------------
    s14 = new_slide(C_WHITE)
    add_clean_header(
        s14,
        "บทที่ 3 : ขั้นตอนที่ 3 — การออกแบบวงจรแอนะล็อกฟรอนต์เอนด์ (1/4)",
        "หลักการทำงานของวงจร OPA380 TIA และการรักษาความเป็นเชิงเส้นที่แดดจัด"
    )
    add_pic_safe(s14, os.path.join(img_dir, "tia_schematic_kicad.png"), 0.85, 1.95, width=5.65)

    c14_r = add_card(s14, 6.75, 1.85, 5.73, 5.25, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf14 = c14_r.text_frame
    p = tf14.paragraphs[0]
    r = p.add_run()
    r.text = "ทำไมต้องกำหนดแรงดัน V_bias = 0.418 V (R3=100k, R4=10k)?"
    r.font.name = FONT_THAI
    r.font.size = Pt(21)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG

    pts14 = [
        (
            "1. หนีจุดอิ่มตัวล่างของ Op-Amp (Output Saturation Floor):",
            "ในระบบไฟเลี้ยงเดี่ยว (Single Supply) เอาต์พุตของ OPA380 ไม่สามารถลงถึง 0.000 V ได้จริง (อิ่มตัวที่ช่วง 60 – 100 mV) การยกฐานด้วย V_bias = 0.418 V ช่วยให้วัดแสงตั้งแต่มืดสนิทได้โดยสัญญาณไม่จมหาย"
        ),
        (
            "2. ป้องกันไดโอดอิ่มตัวเมื่อเจอแดดจัดผ่านกระจกใส:",
            "ด้วยหลักการ Virtual Short ขา -IN จะถูกล็อกไว้ที่ 0.418 V ตลอดเวลา ช่วยดึง Reverse Bias สู้กับแรงดันตกคร่อมความต้านทานภายในเนื้อซิลิคอน (R_S) ไม่ให้รอยต่อ p-n พลิกกลับเป็น Forward Bias (V_OC = 0.365 V) กราฟจึงเป็นเส้นตรง (Linear) 100%"
        ),
        (
            "สมการถ่ายโอนแรงดันเอาต์พุต:",
            "V_op = V_bias + (I_P × R1)  ➔  คำนวณกระแสแสงสุทธิจาก (V_op - V_bias) / R1"
        )
    ]
    for ph, pb in pts14:
        p_item = tf14.add_paragraph()
        p_item.space_before = Pt(8)
        rh = p_item.add_run()
        rh.text = ph + "\n"
        rh.font.name = FONT_THAI
        rh.font.size = Pt(18.5)
        rh.font.bold = True
        rh.font.color.rgb = C_TEXT_BLACK
        rb = p_item.add_run()
        rb.text = pb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(17.5)
        rb.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 15: ขั้นตอนที่ 3.2 ตารางการเลือกค่าตัวต้านทาน R1 รองรับพิกัด 4,000 W/m² (1 หน้าโล่งๆ)
    # -------------------------------------------------------------------------
    s15 = new_slide(C_WHITE)
    add_clean_header(
        s15,
        "บทที่ 3 : ขั้นตอนที่ 3 — การออกแบบวงจรแอนะล็อกฟรอนต์เอนด์ (2/4)",
        "การกำหนดค่าตัวต้านทานป้อนกลับ R1 (R_TIA) เพื่อรองรับพิกัดสูงสุด 4,000 W/m²",
        "สามารถปรับเลือกค่า R1 ตามสภาพแวดล้อมการวัด โดยคงความละเอียดระดับไมโครโวลต์ด้วย ADC 24 บิต"
    )
    rows_s15 = [
        ["ค่าตัวต้านทาน R1 (R_TIA)", "ค่า R4 (V_bias)", "พิกัดวัดแสงสูงสุด (G_max)", "ความละเอียด 24-bit ADC"],
        ["2,200 Ohm (ค่าเริ่มต้นทดสอบบน Perfboard)", "10 kOhm (0.418 V)", "~714 W/m² (ทดสอบในร่ม / ไฟฉาย)", "0.0003 W/m²"],
        ["1,100 Ohm (2.2 kOhm ขนาน 2 ตัว)", "10 kOhm (0.418 V)", "~1,430 W/m² (แดดจัดกลางแจ้งปกติ)", "0.0006 W/m²"],
        ["680 Ohm (พิกัดแดดเขตร้อน)", "10 kOhm (0.418 V)", "~2,310 W/m² (สภาวะขอบเมฆรวมแสง)", "0.0010 W/m²"],
        ["390 Ohm (รุ่นท็อป ISO 9060 Class A)", "10k – 15k (0.42–0.58 V)", "4,000 W/m² (Flagship Class A Spec)", "0.0017 W/m²"]
    ]
    tab15 = s15.shapes.add_table(len(rows_s15), 4, Inches(0.85), Inches(2.05), Inches(11.63), Inches(4.75)).table
    w15 = [3.33, 2.50, 3.30, 2.50]
    for j, w in enumerate(w15):
        tab15.columns[j].width = Inches(w)
    for i, row in enumerate(rows_s15):
        for j, val in enumerate(row):
            cell = tab15.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_GREEN_BG if i == 4 else (C_ROW_ALT if i % 2 == 0 else C_WHITE))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(20 if i == 0 else 19.5)
            r.font.bold = (i == 0 or i == 4)
            r.font.color.rgb = C_WHITE if i == 0 else (C_GREEN_TEXT if i == 4 else C_TEXT_BLACK)

    # -------------------------------------------------------------------------
    # SLIDE 16: ขั้นตอนที่ 3.3 วงจรอ่านค่า 24-Bit ADC (ADS1220) ครบ 4 ช่องสัญญาณ
    # -------------------------------------------------------------------------
    s16 = new_slide(C_WHITE)
    add_clean_header(
        s16,
        "บทที่ 3 : ขั้นตอนที่ 3 — การออกแบบวงจรแอนะล็อกฟรอนต์เอนด์ (3/4)",
        "การออกแบบวงจรอ่านค่า 24-Bit Sigma-Delta ADC (ADS1220) ทั้ง 4 ช่องสัญญาณ"
    )
    add_pic_safe(s16, os.path.join(img_dir, "ads1220_schematic_kicad.png"), 0.85, 1.85, width=5.6)
    add_pic_safe(s16, os.path.join(img_dir, "ADS1220_diagram.png"), 0.85, 4.70, width=3.5)

    c16_r = add_card(s16, 6.75, 1.85, 5.73, 5.20, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf16 = c16_r.text_frame
    chs = [
        (
            "AIN0 & AIN1 : สัญญาณแสง V_op และแรงดันฐาน V_bias",
            "อ่านค่าทั้งสองช่องมาลบกันในซอฟต์แวร์ (V_op - V_bias) เพื่อหักลบออฟเซ็ตและสัญญาณรบกวนร่วม (Common-mode Noise) ออก 100%"
        ),
        (
            "AIN2 : อุณหภูมิชิปอ้างอิง REF5045 (ขา 3 TEMP)",
            "อ่านแรงดันขา TEMP (509 mV @ 0°C หรือ 575 mV @ 25°C, 2.64 mV/°C) เพื่อติดตามเสถียรภาพความร้อนของแหล่งจ่ายอ้างอิง 4.500 V ร่วมกับเซนเซอร์ภายใน ADS1220"
        ),
        (
            "AIN3 : อุณหภูมิผิวเซนเซอร์ NTC Thermistor 10 kOhm (±1%)",
            "ดัดแนบชิดใต้ฐานของตัวเซนเซอร์แสง คำนวณอุณหภูมิหัววัด T_diode (B25/85 = 3975 K) ด้วยสมการ Beta เพื่อชดเชยทั้งกระแสมืด I_dark(T) และสัมประสิทธิ์อุณหภูมิ (+0.18 %/K)"
        )
    ]
    for idx, (ct, cb) in enumerate(chs):
        p = tf16.paragraphs[0] if idx == 0 else tf16.add_paragraph()
        if idx > 0:
            p.space_before = Pt(10)
        r = p.add_run()
        r.text = ct + "\n"
        r.font.name = FONT_THAI
        r.font.size = Pt(20)
        r.font.bold = True
        r.font.color.rgb = C_BLUE_TAG if idx == 0 else C_TEXT_BLACK
        rb = p.add_run()
        rb.text = cb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(18)
        rb.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 17: ขั้นตอนที่ 3.4 การออกแบบภาคจ่ายไฟสัญญาณรบกวนต่ำและแรงดันอ้างอิงความแม่นยำสูง
    # -------------------------------------------------------------------------
    s17 = new_slide(C_WHITE)
    add_clean_header(
        s17,
        "บทที่ 3 : ขั้นตอนที่ 3 — การออกแบบวงจรแอนะล็อกฟรอนต์เอนด์ (4/4)",
        "การออกแบบภาคจ่ายไฟสัญญาณรบกวนต่ำ (TPS7A4901) และแรงดันอ้างอิง (REF5045)",
        "แยกภาคจ่ายไฟแอนะล็อก +5.20 V ออกจากภาคดิจิทัล และสร้างแรงดันอ้างอิงมาตรฐาน 4.500 V (ดริฟต์ต่ำเพียง 3 ppm/°C)"
    )
    add_pic_safe(s17, os.path.join(ext_dir, "p27_img0_793x350.png"), 0.85, 2.10, width=5.75)
    add_pic_safe(s17, os.path.join(ext_dir, "p27_img1_577x318.png"), 6.80, 2.10, width=5.65)

    c17_l = add_card(s17, 0.85, 5.35, 5.75, 1.55, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf17_l = c17_l.text_frame
    p = tf17_l.paragraphs[0]
    r = p.add_run()
    r.text = "1. วงจรแรงดันอ้างอิง REF5045AIDGKR (+4.500 V)"
    r.font.name = FONT_THAI
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG
    p2 = tf17_l.add_paragraph()
    r2 = p2.add_run()
    r2.text = "จ่ายแรงดันอ้างอิงความแม่นยำสูง (±0.05%) ให้กับทั้งขั้ว REFP0 ของ ADS1220, วงจรแบ่งแรงดัน V_bias และวงจรแบ่งแรงดัน NTC (Ratiometric)"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(18)
    r2.font.color.rgb = C_TEXT_BLACK

    c17_r = add_card(s17, 6.80, 5.35, 5.65, 1.55, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tf17_r = c17_r.text_frame
    p = tf17_r.paragraphs[0]
    r = p.add_run()
    r.text = "2. วงจรเรกูเลเตอร์สัญญาณรบกวนต่ำพิเศษ TPS7A4901 (+5.20 V)"
    r.font.name = FONT_THAI
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_TEXT
    p2 = tf17_r.add_paragraph()
    r2 = p2.add_run()
    r2.text = "กรองสัญญาณรบกวนจากภาคสวิตชิ่ง (PSRR สูงถึง 72 dB) และรักษาระยะเผื่อแรงดัน (Headroom 700 mV) เหนือระดับ 4.500 V ให้วงจรแอนะล็อกนิ่งสนิท"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(18)
    r2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 18: ขั้นตอนที่ 4 การออกแบบเคส 3D CAD สำหรับ Perfboard และแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน
    # -------------------------------------------------------------------------
    s18 = new_slide(C_WHITE)
    add_clean_header(
        s18,
        "บทที่ 3 : ขั้นตอนที่ 4 — การออกแบบทางกลและแผ่นวงจรพิมพ์ (Mechanical & PCB Design)",
        "การออกแบบเคส 3D CAD สำหรับ Perfboard และแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน",
        "พร้อมสำหรับการประกอบทดสอบกลางแจ้งและการผลิตแผ่นวงจรพิมพ์มาตรฐานโรงงาน (DRC Violations = 0)"
    )
    add_pic_safe(s18, os.path.join(img_dir, "cad_perfboard_case_assy.png"), 0.85, 2.0, width=2.75)
    add_pic_safe(s18, os.path.join(img_dir, "cad_sapphire_retaining_ring.png"), 3.78, 2.0, width=2.75)
    add_pic_safe(s18, os.path.join(img_dir, "pcb_3d_render_iso.png"), 6.71, 2.0, width=2.75)
    add_pic_safe(s18, os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png"), 9.64, 2.0, width=2.75)

    cards_s18 = [
        (0.85, "1. กล่องเคสสำหรับ Perfboard", "ออกแบบขนาดพอดีบอร์ด 70×90 mm พร้อมช่องหน้าต่างรับแสงทรงกลมด้านบนและพอร์ต Cable Gland กันน้ำด้านข้าง"),
        (3.78, "2. แหวนล็อกกระจก Sapphire", "ฝาแหวนกดล็อกกระจกแซฟไฟร์ (Retaining Bezel Ring) แบบมีร่องขันล็อก 6 จุด ถอดสลับทดสอบกระจกได้สะดวก"),
        (6.71, "3. 4-Layer PCB (รุ่น BPX 61)", "แผ่นวงจรพิมพ์ 4 เลเยอร์ รุ่น V1.1 L1_Analog สำหรับตัวถังโลหะ TO-39 พร้อมวงแหวน Guard Ring รอบขา -IN"),
        (9.64, "4. 4-Layer PCB (รุ่น BPW 34)", "แผ่นวงจรพิมพ์ 4 เลเยอร์ รุ่น V1.2 L1_Analog_BPW34 สำหรับตัวถังแบนราบ DIL เพื่อทดสอบเปรียบเทียบแบบควบคุมตัวแปร")
    ]
    for left, ct, cb in cards_s18:
        c = add_card(s18, left, 4.95, 2.75, 1.95, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = c.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = ct
        r.font.name = FONT_THAI
        r.font.size = Pt(18.5)
        r.font.bold = True
        r.font.color.rgb = C_BLUE_TAG
        pb = tf.add_paragraph()
        pb.space_before = Pt(3)
        rb = pb.add_run()
        rb.text = cb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(16.5)
        rb.font.color.rgb = C_TEXT_BODY

    # =========================================================================
    # บทที่ 4: ผลการทดลองและอภิปรายผล — Results & Discussion (Slides 19 - 23)
    # =========================================================================

    # SLIDE 19: ผลการจำลองการตอบสนองเชิงมุม BPX 61 (+ ตอบคำถามอาจารย์: ใช้โปรแกรมอะไรยังไง)
    s19 = new_slide(C_WHITE)
    add_clean_header(
        s19,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (1/7)",
        "ผลการจำลองการตอบสนองเชิงมุม 3 มิติของเซนเซอร์ BPX 61 (Optical Ray-Tracing)",
        tag_color=C_ORANGE_TAG
    )
    add_pic_safe(s19, os.path.join(ext_dir, "p19_img0_1024x558.png"), 0.85, 1.62, width=6.80)
    add_pic_safe(s19, os.path.join(img_dir, "CAD1.png"), 8.05, 1.62, width=2.10)
    add_pic_safe(s19, os.path.join(img_dir, "CAD2.png"), 10.30, 1.62, width=2.10)
    add_pic_safe(s19, os.path.join(ext_dir, "p19_img1_522x381.png"), 8.35, 2.95, height=2.35)

    c19_bot = add_card(s19, 0.85, 5.42, 11.60, 1.68, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf19 = c19_bot.text_frame
    p19_1 = tf19.paragraphs[0]
    r19_1 = p19_1.add_run()
    r19_1.text = "ซอฟต์แวร์และระเบียบวิธีจำลอง (Simulation Software & Method): "
    r19_1.font.name = FONT_THAI
    r19_1.font.size = Pt(18)
    r19_1.font.bold = True
    r19_1.font.color.rgb = C_BLUE_TAG
    r19_1b = p19_1.add_run()
    r19_1b.text = (
        "ส่งออกโมเดล 3 มิติ (.OBJ) จากโปรแกรม FreeCAD เข้าสู่เอ็นจิน Mitsuba 3 (Monte Carlo Physically-Based Optical Ray-Tracing ในภาษา Python) "
        "โดยกำหนดวัสดุตัวถังเป็นตัวนำสะท้อนแสงและกระจกหน้าต่างเป็น Dielectric BSDF ตามสมการเฟรสเนล (Schott 8250 Borosilicate, n = 1.487)"
    )
    r19_1b.font.name = FONT_THAI
    r19_1b.font.size = Pt(17)
    r19_1b.font.color.rgb = C_TEXT_BLACK

    p19_2 = tf19.add_paragraph()
    p19_2.space_before = Pt(3)
    r19_2 = p19_2.add_run()
    r19_2.text = "การอภิปรายผล (Discussion): "
    r19_2.font.name = FONT_THAI
    r19_2.font.size = Pt(18)
    r19_2.font.bold = True
    r19_2.font.color.rgb = C_RED_TEXT
    r19_2b = p19_2.add_run()
    r19_2b.text = "เมื่อไม่มีกระจก ขอบกระป๋องโลหะจะตัดแสงที่มุม ~62° (ตรงตามสเปกชีต) แต่เมื่อรวมการหักเหและสะท้อนของกระจกหน้าต่าง มุมตัดแสงจะหดสั้นลงเหลือเพียง ~44° เท่านั้น"
    r19_2b.font.name = FONT_THAI
    r19_2b.font.size = Pt(17)
    r19_2b.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 20: ผลการจำลองเปรียบเทียบกับ BPW 34 และการอภิปรายผล (Discussion)
    # -------------------------------------------------------------------------
    s20 = new_slide(C_WHITE)
    add_clean_header(
        s20,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (2/7)",
        "ผลการจำลองการตอบสนองเชิงมุมของ BPW 34 และการอภิปรายผล (Discussion)",
        tag_color=C_ORANGE_TAG
    )
    add_pic_safe(s20, os.path.join(ext_dir, "p20_img0_1024x531.png"), 0.85, 1.65, height=4.15)
    add_pic_safe(s20, os.path.join(img_dir, "cad_sapphire_retaining_ring.png"), 9.85, 1.70, width=2.55)
    add_pic_safe(s20, os.path.join(img_dir, "cad_bpw34_model.png"), 9.85, 3.60, height=2.15)

    c20_bot = add_card(s20, 0.85, 5.90, 11.60, 1.20, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tf20 = c20_bot.text_frame
    p20 = tf20.paragraphs[0]
    r20 = p20.add_run()
    r20.text = "การอภิปรายผลการจำลอง BPW 34 (Discussion): "
    r20.font.name = FONT_THAI
    r20.font.size = Pt(18.5)
    r20.font.bold = True
    r20.font.color.rgb = C_GREEN_TEXT
    r20_b = p20.add_run()
    r20_b.text = (
        "1) BPW 34 ตัวเปล่า (แถวล่าง) ตอบสนองใกล้เคียง Ideal Cosine แต่เมื่ออยู่ใต้กระจกแซฟไฟร์แบนและแหวนล็อก (แถวบน) ที่มุม >70° ค่าจะดรอปลงตามธรรมชาติของกระจกแบน (ต่างจากเครื่องราคาแพงที่ใช้โดมแก้ว)  "
        "2) แนวทางจัดการในระบบ: แม้หัววัดไม่ตรง Ideal 100% แต่ใช้งานได้จริง โดยวัดค่าจริงตามมุม (0°–85°) ในแล็บฟิสิกส์ แล้วใช้พิกัดติดตั้งโซลาร์เซลล์ + เวลาจริง (รู้มุมดวงอาทิตย์) มาชดเชยในซอฟต์แวร์"
    )
    r20_b.font.name = FONT_THAI
    r20_b.font.size = Pt(17.0)
    r20_b.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 21: ผลงานการบัดกรีประกอบบอร์ดต้นแบบ Perfboard จริง
    # -------------------------------------------------------------------------
    s21 = new_slide(C_WHITE)
    add_clean_header(
        s21,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (3/7)",
        "ผลงานการบัดกรีประกอบบอร์ดต้นแบบ Perfboard แบบแยกด้านรับแสงและด้านวงจร",
        "ด้านบนติดตั้งเฉพาะหัววัดแสงและ NTC เพื่อให้เปิดมุมรับแสง 180° ไร้เงาบัง ส่วนด้านล่างรวมโมดูลวงจรครบทั้ง 6 ภาค"
    )
    add_pic_safe(s21, os.path.join(img_dir, "perfboard_top_view.jpg"), 0.85, 2.0, width=3.75)
    add_pic_safe(s21, os.path.join(img_dir, "perfboard_side_ntc_view.jpg"), 4.78, 2.0, width=3.75)
    add_pic_safe(s21, os.path.join(img_dir, "perfboard_bottom_modules.jpg"), 8.71, 2.0, width=3.75)

    caps_s21 = [
        (0.85, "1. ด้านบนบอร์ด (Top Optical Side)", "ติดตั้งเฉพาะเซนเซอร์แสงกึ่งกลางบอร์ดและ NTC ไร้ชิ้นส่วนสูงบังแสงที่มุมเอียง"),
        (4.78, "2. มุมมองด้านข้างและการแนบ NTC", "ดัดหัววัด NTC 10 kOhm (สีเขียว) แนบสนิทกับฐานโลหะใต้ตัวถังเซนเซอร์เพื่อวัดอุณหภูมิจริง"),
        (8.71, "3. ด้านล่างบอร์ด (Bottom Electronics)", "ประกอบด้วย LM2596, TPS7A4901 (+5.2V), REF5045 (4.5V), OPA380, ADS1220 และ ESP32-C6")
    ]
    for left, ct, cb in caps_s21:
        card = add_card(s21, left, 5.05, 3.75, 1.85, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
        tf = card.text_frame
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
        rb.font.size = Pt(17.5)
        rb.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 22: การวิเคราะห์และอภิปรายผลการดีบักภาคจ่ายไฟจริง (พร้อมรูปวัดมิเตอร์จริง 5.194V และ 4.495V!)
    # -------------------------------------------------------------------------
    s22 = new_slide(C_WHITE)
    add_clean_header(
        s22,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (4/7)",
        "การวิเคราะห์และอภิปรายผลการดีบักภาคจ่ายไฟ LDO TPS7A4901 และ REF5045",
        tag_color=C_ORANGE_TAG
    )
    c22_red = add_card(s22, 0.75, 1.80, 4.15, 5.20, bg_rgb=C_RED_BG, border_rgb=C_RED_BORDER)
    tfr22 = c22_red.text_frame
    p = tfr22.paragraphs[0]
    r = p.add_run()
    r.text = "ปัญหาทางวิศวกรรมที่พบขณะทดสอบ"
    r.font.name = FONT_THAI
    r.font.size = Pt(20.5)
    r.font.bold = True
    r.font.color.rgb = C_RED_TEXT

    probs22 = [
        ("1. แผ่นแปลงไม่มี Footprint ขา 9 (Thermal Pad):", "เมื่อบัดกรี TPS7A4901 ลงแผ่นแปลง MSOP-8 ทั่วไปที่ไม่มีแผ่นกราวด์ใต้ท้องชิป ไฟออกตกเหลือ < 0.8 V เพราะไอซีบังคับใช้ขา 9 เป็นกราวด์หลัก"),
        ("2. กระแสรั่วไหลขา 5 (Trim/NR) ของ REF5045:", "โหนดความต้านทานสูง (High-Z) เกิดกระแสรั่วบนบอร์ดทดลองทำให้ไฟอ้างอิงตกเหลือ 4.382 V"),
        ("3. ไฟเลี้ยง USB ดรอป & ค่า ESR ตัวเก็บประจุ:", "สาย USB แรงดันตกต่ำกว่า 4.70 V ทำให้ REF5045 หลุดสภาวะควบคุม (Dropout)")
    ]
    for ph, pb in probs22:
        pp = tfr22.add_paragraph()
        pp.space_before = Pt(8)
        rh = pp.add_run()
        rh.text = ph + " "
        rh.font.name = FONT_THAI
        rh.font.size = Pt(17.5)
        rh.font.bold = True
        rh.font.color.rgb = C_RED_TEXT
        rb = pp.add_run()
        rb.text = pb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(16.8)
        rb.font.color.rgb = C_TEXT_BLACK

    c22_grn = add_card(s22, 5.02, 1.80, 4.18, 5.20, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tfg22 = c22_grn.text_frame
    p = tfg22.paragraphs[0]
    r = p.add_run()
    r.text = "การแก้ไขและผลวัดแรงดันจริงบนบอร์ด"
    r.font.name = FONT_THAI
    r.font.size = Pt(20.5)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_TEXT

    sols22 = [
        ("1. ทดสอบ TPS7A4901 บน PCB ที่มี Footprint ขา 9:", "เนื่องจากไม่มีแผ่นแปลงตรงรุ่น จึงนำไอซีไปบัดกรีทดสอบบนแผ่น PCB รุ่นแรกที่มีแผ่นกราวด์ขา 9 วัดไฟออกได้จริง 5.194 V (HI: 5.201 V) ตรงพิกัด +5.20 V"),
        ("2. ปลดขา 5 (DNP Trim) & ใส่ C Low-ESR 10 µF:", "ขจัดกระแสรั่วไหลและรักษาเสถียรภาพลูปควบคุม วัดแรงดัน REF5045 บน Perfboard ได้จริง 4.495 V (HI: 4.496 V) ตรงพิกัด +4.500 V"),
        ("3. สถาปัตยกรรมจ่ายไฟ 2 ชั้น (LM2596 ➔ LDO):", "สร้าง Headroom 700 mV เหนือ 4.500 V ไร้ผลกระทบจากการดึงกระแสส่ง Wi-Fi ของ ESP32-C6")
    ]
    for sh, sb in sols22:
        ps = tfg22.add_paragraph()
        ps.space_before = Pt(8)
        rh = ps.add_run()
        rh.text = sh + " "
        rh.font.name = FONT_THAI
        rh.font.size = Pt(17.5)
        rh.font.bold = True
        rh.font.color.rgb = C_GREEN_TEXT
        rb = ps.add_run()
        rb.text = sb
        rb.font.name = FONT_THAI
        rb.font.size = Pt(16.8)
        rb.font.color.rgb = C_TEXT_BLACK

    add_pic_safe(s22, os.path.join(img_dir, "meas_tps7a4901_5v194.jpg"), 9.35, 1.80, width=3.25)
    c22_cap1 = add_card(s22, 9.35, 3.63, 3.25, 0.42, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    c22_cap1.text_frame.margin_left = Inches(0.04)
    c22_cap1.text_frame.margin_right = Inches(0.04)
    c22_cap1.text_frame.margin_top = Inches(0.04)
    c22_cap1.text_frame.margin_bottom = Inches(0.04)
    p22_c1 = c22_cap1.text_frame.paragraphs[0]
    p22_c1.alignment = PP_ALIGN.CENTER
    r22_c1 = p22_c1.add_run()
    r22_c1.text = "TPS7A4901 บน PCB = 5.194 V (HI: 5.201V)"
    r22_c1.font.name = FONT_THAI
    r22_c1.font.size = Pt(15.5)
    r22_c1.font.bold = True
    r22_c1.font.color.rgb = C_BLUE_TAG

    add_pic_safe(s22, os.path.join(img_dir, "meas_ref5045_4v495.jpg"), 9.68, 4.18, height=2.28)
    c22_cap2 = add_card(s22, 9.35, 6.54, 3.25, 0.42, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_TEXT)
    c22_cap2.text_frame.margin_left = Inches(0.04)
    c22_cap2.text_frame.margin_right = Inches(0.04)
    c22_cap2.text_frame.margin_top = Inches(0.04)
    c22_cap2.text_frame.margin_bottom = Inches(0.04)
    p22_c2 = c22_cap2.text_frame.paragraphs[0]
    p22_c2.alignment = PP_ALIGN.CENTER
    r22_c2 = p22_c2.add_run()
    r22_c2.text = "REF5045 บน Perfboard = 4.495 V (HI: 4.496V)"
    r22_c2.font.name = FONT_THAI
    r22_c2.font.size = Pt(15.5)
    r22_c2.font.bold = True
    r22_c2.font.color.rgb = C_GREEN_TEXT

    # -------------------------------------------------------------------------
    # SLIDE 23: ผลการอ่านสัญญาณจริง 24 บิตใน 4 สภาวะแสง และการอภิปรายผล (Discussion)
    # -------------------------------------------------------------------------
    s23 = new_slide(C_WHITE)
    add_clean_header(
        s23,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (5/7)",
        "ผลการอ่านสัญญาณจริง 24 บิตครบทั้ง 4 ช่องสัญญาณ และการอภิปรายผล (Discussion)"
    )
    rows_s23 = [
        ["สภาวะแสงทดสอบจริง", "AIN0: V_op", "AIN1: V_bias", "AIN2: Temp REF", "AIN3: NTC (Temp)", "แรงดันสุทธิ (V_net)", "กระแส I_P / ความเข้มแสง"],
        ["1. ปิดทึบมืดสนิท (Dark)", "0.4185 V", "0.4180 V", "0.5940 V (32.20 °C)", "2.0942 V (28.48 °C)", "+0.50 mV", "0.227 µA (~0.09 W/m²)"],
        ["2. แสงสว่างปกติในห้อง", "0.4405 V", "0.4180 V", "0.5942 V (32.27 °C)", "2.0940 V (28.49 °C)", "+22.50 mV", "10.23 µA (~3.93 W/m²)"],
        ["3. โคมไฟส่องระยะใกล้", "1.5420 V", "0.4180 V", "0.5948 V (32.50 °C)", "2.0885 V (28.62 °C)", "+1,124.0 mV", "510.9 µA (~196.5 W/m²)"],
        ["4. ไฟฉายแรงสูงจ่อชิด (R1=2.2k)", "4.4998 V (เต็ม)", "0.4181 V", "0.5951 V (32.61 °C)", "2.0810 V (28.80 °C)", "+4,081.7 mV", "1,855 µA (~714 W/m²)"]
    ]
    tab23 = s23.shapes.add_table(5, 7, Inches(0.85), Inches(1.68), Inches(11.63), Inches(2.90)).table
    w23 = [2.35, 1.30, 1.25, 1.72, 1.72, 1.40, 1.89]
    for j, w in enumerate(w23):
        tab23.columns[j].width = Inches(w)

    for i, row in enumerate(rows_s23):
        for j, val in enumerate(row):
            cell = tab23.cell(i, j)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_NAVY_HEADER if i == 0 else (C_BLUE_BG if i == 4 else (C_ROW_ALT if i % 2 == 0 else C_WHITE))
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = val
            r.font.name = FONT_THAI
            r.font.size = Pt(17 if i == 0 else 16.5)
            r.font.bold = (i == 0 or j == 6)
            r.font.color.rgb = C_WHITE if i == 0 else (C_BLUE_TAG if j == 6 else C_TEXT_BLACK)

    c23_disc = add_card(s23, 0.85, 4.82, 11.63, 2.28, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf23_d = c23_disc.text_frame
    p23_1 = tf23_d.paragraphs[0]
    r23_1 = p23_1.add_run()
    r23_1.text = "การอภิปรายผลการทดลอง (Experimental Discussion):"
    r23_1.font.name = FONT_THAI
    r23_1.font.size = Pt(19)
    r23_1.font.bold = True
    r23_1.font.color.rgb = C_BLUE_TAG

    disc_items = [
        "1) เสถียรภาพของ V_bias: ช่อง AIN1 นิ่งสนิทที่ 0.4180–0.4181 V ตลอดทุกสภาวะแสง ยืนยันการทำงานของ Virtual Short บน OPA380 และช่วยยกสัญญาณในที่มืด (+0.50 mV) พ้นโซนอิ่มตัว",
        "2) การอิ่มตัวตามทฤษฎีและการขยายพิกัด: ที่ค่าเริ่มต้น R1 = 2.2 kOhm เมื่อส่องไฟฉายแรงสูง เอาต์พุตอิ่มตัวที่ 4.4998 V (~714 W/m²) ตรงตามสมการ พิสูจน์ว่าการปรับ R1 เป็น 390 Ohm จะรองรับ 4,000 W/m² ได้ทันที",
        "3) การตรวจจับความร้อนจริงของ NTC: เมื่อส่องไฟฉายระยะใกล้ อุณหภูมิที่ NTC อ่านได้เพิ่มขึ้นจาก 28.48 °C เป็น 28.80 °C ทันที พร้อมสำหรับการคำนวณชดเชยสัมประสิทธิ์ +0.18 %/K"
    ]
    for ditem in disc_items:
        pd = tf23_d.add_paragraph()
        pd.space_before = Pt(2)
        rd = pd.add_run()
        rd.text = ditem
        rd.font.name = FONT_THAI
        rd.font.size = Pt(17)
        rd.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 24 (New): ผลการออกแบบต้นแบบ Web Dashboard (UI Mockup 1/2 - Overview & PR)
    # -------------------------------------------------------------------------
    s23_web1 = new_slide(C_WHITE)
    add_clean_header(
        s23_web1,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (6/7)",
        "ผลการออกแบบต้นแบบ Web Dashboard (1/2): หน้าหลักประเมินค่า PR แบบเวลาจริง",
        "เชื่อมโยงค่าความเข้มแสงสุทธิ (W/m²) จากหัววัดเข้ากับข้อมูลกำลังไฟฟ้าเพื่อคำนวณ Performance Ratio (PR) อัตโนมัติ"
    )
    add_card(s23_web1, 0.80, 1.92, 7.35, 5.10, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    add_pic_safe(s23_web1, os.path.join(web_dir, "slide15_pic01.jpg"), 0.92, 2.02, width=7.10)

    web1_cards = [
        (
            "1. ประมวลผลค่า Raw & คำนวณค่า PR",
            "Data Logger รับค่า Raw จากหัววัดมาคำนวณชดเชยเป็นค่า Irradiance (เช่น 876 W/m²) เทียบกับกำลังไฟ AC Power (1.68 kW) เพื่อคำนวณค่า PR (87.3%) บนเว็บ",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "2. เครื่องมือตรวจสอบหน้างานบน Web",
            "เว็บแอปพลิเคชันรันอยู่บนเซิร์ฟเวอร์ให้วิศวกรเปิดดูได้ทันที เพื่อแยกแยะว่าวันที่ผลิตไฟลดลงเกิดจากเมฆบัง หรือเกิดจากแผงโซลาร์สกปรก/เสื่อมสภาพ",
            C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "3. พัฒนา Data Logger เองคู่กับหัววัด 3,000฿",
            "เฉพาะหัววัด 24 บิตใช้งบเพียง < 3,000 บาท โดยส่งค่า Raw ทุกช่อง (ผ่าน Wi-Fi หรือ RS-485 Modbus) เข้าสู่ระบบ Data Logger ที่ทีมพัฒนาขึ้นเองตามมาตรฐาน",
            C_AMBER_BG, C_ORANGE_TAG
        )
    ]
    for w_idx, (wt, wb, wbg, wtc) in enumerate(web1_cards):
        top_w = 1.92 + w_idx * 1.74
        cw = add_card(s23_web1, 8.35, top_w, 4.15, 1.62, bg_rgb=wbg, border_rgb=wtc)
        tfw = cw.text_frame
        pw1 = tfw.paragraphs[0]
        rw1 = pw1.add_run()
        rw1.text = wt
        rw1.font.name = FONT_THAI
        rw1.font.size = Pt(18.5)
        rw1.font.bold = True
        rw1.font.color.rgb = wtc
        pw2 = tfw.add_paragraph()
        pw2.space_before = Pt(3)
        rw2 = pw2.add_run()
        rw2.text = wb
        rw2.font.name = FONT_THAI
        rw2.font.size = Pt(16.5)
        rw2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 25 (New): ผลการออกแบบต้นแบบ Web Dashboard (UI Mockup 2/2 - Analytics, Environmental & Modbus)
    # -------------------------------------------------------------------------
    s23_web2 = new_slide(C_WHITE)
    add_clean_header(
        s23_web2,
        "บทที่ 4 : ผลการทดลองและอภิปรายผล (7/7)",
        "ผลการออกแบบต้นแบบ Web Dashboard (2/2): กราฟรายวันและพารามิเตอร์ไฟฟ้า"
    )
    add_card(s23_web2, 0.85, 1.68, 5.68, 2.82, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    add_pic_safe(s23_web2, os.path.join(web_dir, "slide16_pic02.jpg"), 0.95, 1.75, width=5.48)

    add_card(s23_web2, 6.80, 1.68, 5.68, 2.82, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    add_pic_safe(s23_web2, os.path.join(web_dir, "slide17_pic01.jpg"), 6.90, 1.75, width=5.48)

    cw2_l = add_card(s23_web2, 0.85, 4.60, 5.68, 1.05, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tfw2_l = cw2_l.text_frame
    p = tfw2_l.paragraphs[0]
    r = p.add_run()
    r.text = "1. แท็บ Environmental & Historical Analytics: "
    r.font.name = FONT_THAI
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG
    rb = p.add_run()
    rb.text = "แสดงกราฟความเข้มรังสีอาทิตย์ตลอด 24 ชั่วโมง (W/m²) และอุณหภูมิหัววัดจาก ESP32-C6 แบบ Real-time"
    rb.font.name = FONT_THAI
    rb.font.size = Pt(17)
    rb.font.color.rgb = C_TEXT_BLACK

    cw2_r = add_card(s23_web2, 6.80, 4.60, 5.68, 1.05, bg_rgb=C_GREEN_BG, border_rgb=C_GREEN_BORDER)
    tfw2_r = cw2_r.text_frame
    p = tfw2_r.paragraphs[0]
    r = p.add_run()
    r.text = "2. แท็บ System Parameters (Modbus Inverter): "
    r.font.name = FONT_THAI
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = C_GREEN_TEXT
    rb = p.add_run()
    rb.text = "ดึงค่ากำลังไฟ แรงดัน และกระแสไฟฟ้า 3 เฟสจากอินเวอร์เตอร์ผ่าน Modbus มาแสดงร่วมกับค่าแสงในหน้าเดียว"
    rb.font.name = FONT_THAI
    rb.font.size = Pt(17)
    rb.font.color.rgb = C_TEXT_BLACK

    cw2_disc = add_card(s23_web2, 0.85, 5.76, 11.63, 1.32, bg_rgb=C_AMBER_BG, border_rgb=C_ORANGE_TAG)
    tfw2_d = cw2_disc.text_frame
    p = tfw2_d.paragraphs[0]
    r = p.add_run()
    r.text = "การอภิปรายผลสถาปัตยกรรม Data Logger & Web (System Discussion): "
    r.font.name = FONT_THAI
    r.font.size = Pt(18.5)
    r.font.bold = True
    r.font.color.rgb = C_ORANGE_TAG
    rb = p.add_run()
    rb.text = (
        "การส่งค่า Raw ทุกช่องสัญญาณ (V_op, V_bias, V_NTC, V_TempREF) จากหัววัดผ่าน Wi-Fi หรือสาย RS-485 (Modbus) เข้าสู่ Data Logger ที่พัฒนาขึ้นเอง "
        "ช่วยให้บันทึกข้อมูลดิบและปรับปรุงสมการชดเชย (ทั้งอุณหภูมิและมุมดวงอาทิตย์ตามพิกัด/เวลา) บนเซิร์ฟเวอร์เพื่อแสดงผลบน Web Dashboard ได้อย่างยืดหยุ่น"
    )
    rb.font.name = FONT_THAI
    rb.font.size = Pt(17)
    rb.font.color.rgb = C_TEXT_BLACK

    # =========================================================================
    # บทที่ 5: สรุปความก้าวหน้าและแผนการดำเนินงานในโครงงาน 2 (Slides 24 - 27)
    # =========================================================================

    # SLIDE 24: การออกแบบระเบียบวิธีทดสอบทางทัศนศาสตร์จริง (1/2 - Vernier LabQuest 2)
    s24 = new_slide(C_WHITE)
    add_clean_header(
        s24,
        "บทที่ 5 : แผนการทดสอบร่วมกับภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ (1/2)",
        "การออกแบบระเบียบวิธีการทดลองวัดการส่องผ่านสเปกตรัมของกระจกแซฟไฟร์"
    )
    add_pic_safe(s24, os.path.join(img_dir, "labquest2_vernier_spectrometer.png"), 1.05, 1.90, height=3.1)
    add_pic_safe(s24, os.path.join(ext_dir, "p23_img1_816x480.png"), 6.65, 1.90, height=3.1)

    c24_bot = add_card(s24, 2.65, 5.25, 8.0, 1.75, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf24 = c24_bot.text_frame
    p24 = tf24.paragraphs[0]
    p24.alignment = PP_ALIGN.CENTER
    r24 = p24.add_run()
    r24.text = "1. การตรวจวัดค่าร้อยละการส่องผ่านสเปกตรัม (%Transmittance)"
    r24.font.name = FONT_THAI
    r24.font.size = Pt(21)
    r24.font.bold = True
    r24.font.color.rgb = C_TEXT_BLACK
    p24_2 = tf24.add_paragraph()
    p24_2.alignment = PP_ALIGN.CENTER
    p24_2.space_before = Pt(6)
    r24_2 = p24_2.add_run()
    r24_2.text = "ใช้เครื่องสเปกโตรมิเตอร์ Vernier LabQuest 2 วัดการส่องผ่านของแผ่นกระจกแซฟไฟร์จริงในช่วงความยาวคลื่น 350 – 900 nm เพื่อยืนยันค่าคงตัวทางแสง (%T ≈ 85%)"
    r24_2.font.name = FONT_THAI
    r24_2.font.size = Pt(19)
    r24_2.font.color.rgb = C_TEXT_BODY

    # -------------------------------------------------------------------------
    # SLIDE 25: การออกแบบระเบียบวิธีทดสอบทางทัศนศาสตร์จริง (2/2 - Rotation Stage & ลำแสงขนานเสมือน)
    # -------------------------------------------------------------------------
    s25 = new_slide(C_WHITE)
    add_clean_header(
        s25,
        "บทที่ 5 : แผนการทดสอบร่วมกับภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ (2/2)",
        "การทดสอบการตอบสนองเชิงมุมบนแท่นหมุนตามเกณฑ์ ISO 9060:2018 Annex A.5"
    )
    add_pic_safe(s25, os.path.join(ext_dir, "p33_img1_1024x576.png"), 0.85, 1.80, width=4.45)
    add_pic_safe(s25, os.path.join(ext_dir, "p33_img0_763x674.png"), 5.50, 1.80, height=2.95)
    add_pic_safe(s25, os.path.join(iso_dir, "iso_p21_annex_a5.png"), 9.05, 1.72, height=3.20)

    c25_l = add_card(s25, 0.85, 5.15, 5.65, 1.85, bg_rgb=C_CARD_GRAY, border_rgb=C_BORDER_GRAY)
    tf25_l = c25_l.text_frame
    p = tf25_l.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "2. ทดสอบมุมตกกระทบบนแท่นหมุน (ISO 9060 Annex A.5)"
    r.font.name = FONT_THAI
    r.font.size = Pt(19.5)
    r.font.bold = True
    r.font.color.rgb = C_TEXT_BLACK
    p2 = tf25_l.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "ยึดหัววัดบนแท่นหมุนวัดค่าที่อ่านได้จริงในแต่ละมุม (0°–85°) เพื่อนำไปใช้คู่กับพิกัดที่ติดตั้งโซลาร์เซลล์และเวลาจริง (คำนวณตำแหน่งมุมดวงอาทิตย์) ในการชดเชยค่าความเข้มแสงในระบบซอฟต์แวร์"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(17.5)
    r2.font.color.rgb = C_TEXT_BODY

    c25_r = add_card(s25, 6.75, 5.15, 5.73, 1.85, bg_rgb=C_BLUE_BG, border_rgb=C_BLUE_TAG)
    tf25_r = c25_r.text_frame
    p = tf25_r.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "3. การจัดชุดกำเนิดลำแสงขนานเสมือน (Collimated Beam)"
    r.font.name = FONT_THAI
    r.font.size = Pt(19.5)
    r.font.bold = True
    r.font.color.rgb = C_BLUE_TAG
    p2 = tf25_r.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "ใช้แหล่งกำเนิดแสงแบบจุด (Point Source) ร่วมกับชุดเลนส์ขยายและปรับลำแสงขนาน (Collimating Lens) ตามคำแนะนำของอาจารย์ที่ปรึกษาด้านฟิสิกส์ เพื่อให้แสงตกกระทบสม่ำเสมอ"
    r2.font.name = FONT_THAI
    r2.font.size = Pt(17.5)
    r2.font.color.rgb = C_TEXT_BLACK

    # -------------------------------------------------------------------------
    # SLIDE 26: สถานะปัจจุบัน การอภิปรายสรุปภาพรวม และแผนการดำเนินงานสู่โครงงาน 2 (ล้อตาม 6 ขั้นตอนหลัก)
    # -------------------------------------------------------------------------
    s26 = new_slide(C_WHITE)
    add_clean_header(
        s26,
        "บทที่ 5 : สรุปความก้าวหน้า การอภิปรายผลภาพรวม และแผนงานต่อยอด",
        "สถานะปัจจุบันและแผนการดำเนินงานสู่โครงงาน 2 (Roadmap 6 ขั้นตอนหลัก)"
    )
    roadmap_cards = [
        (
            "ขั้นตอนที่ 1 • สำเร็จแล้ว 100%",
            "จำลองแสง 3 มิติ (FreeCAD & Mitsuba 3)",
            "สร้าง 3D CAD ทั้ง BPX 61 และ BPW 34 รัน Ray-Tracing พิสูจน์จุดตัดแสง 44° สำเร็จ",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "ขั้นตอนที่ 2 • สำเร็จแล้ว 100%",
            "เคสปิดผนึก & กระจกแซฟไฟร์ (Al2O3)",
            "ออกแบบเคส 3D CAD และแหวนล็อกกระจกแซฟไฟร์ความแข็ง 9 Mohs ป้องกัน UV และความชื้น",
            C_BLUE_BG, C_BLUE_TAG
        ),
        (
            "ขั้นตอนที่ 3 • สำเร็จแล้ว 100%",
            "วงจร TIA 24 บิต & ภาคจ่ายไฟ",
            "ออกแบบวงจร OPA380 TIA (V_bias = 0.418 V), ADS1220 24-bit, NTC และไฟอ้างอิง REF5045",
            C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "ขั้นตอนที่ 4 • สำเร็จแล้ว 65%",
            "บอร์ด Perfboard จริง & 4-Layer PCB",
            "บัดกรีเทสบน Perfboard อ่านค่า 24 บิตแล้ว และออกแบบ 4-Layer PCB แล้ว (รอบัดกรีเทสลงแผ่น PCB จริงในโครงงาน 2)",
            C_GREEN_BG, C_GREEN_TEXT
        ),
        (
            "ขั้นตอนที่ 5 • โครงงาน 2 (ระยะที่ 1)",
            "เก็บกราฟ Cosine จริงในแล็บฟิสิกส์",
            "วัดสเปกตรัมด้วย LabQuest 2 และหมุนวัดค่าจริงตามมุม 0°–85° บนแท่นหมุน Optical Rotation Stage",
            C_AMBER_BG, C_ORANGE_TAG
        ),
        (
            "ขั้นตอนที่ 6 • โครงงาน 2 (ระยะที่ 2)",
            "Data Logger, ชดเชยพิกัดเวลา & Web",
            "ส่งค่า Raw (Wi-Fi / RS-485) เข้า Data Logger ชดเชยด้วยผลแล็บฟิสิกส์ + พิกัดเวลาจริง และสอบเทียบกลางแจ้ง",
            C_AMBER_BG, C_ORANGE_TAG
        )
    ]
    for idx, (tag, rtitle, rbody, bg_c, tc) in enumerate(roadmap_cards):
        col = idx % 3
        row = idx // 3
        left = 0.85 + col * 3.94
        top = 1.65 + row * 1.92
        c = add_card(s26, left, top, 3.75, 1.80, bg_rgb=bg_c, border_rgb=tc, border_pt=1.5)
        tf = c.text_frame
        p_t = tf.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = tag
        r_t.font.name = FONT_THAI
        r_t.font.size = Pt(16.5)
        r_t.font.bold = True
        r_t.font.color.rgb = tc

        p_h = tf.add_paragraph()
        p_h.space_before = Pt(2)
        p_h.space_after = Pt(2)
        r_h = p_h.add_run()
        r_h.text = rtitle
        r_h.font.name = FONT_THAI
        r_h.font.size = Pt(18.5)
        r_h.font.bold = True
        r_h.font.color.rgb = C_TEXT_BLACK

        p_b = tf.add_paragraph()
        r_b = p_b.add_run()
        r_b.text = rbody
        r_b.font.name = FONT_THAI
        r_b.font.size = Pt(16.0)
        r_b.font.color.rgb = C_TEXT_BODY

    c26_disc = add_card(s26, 0.85, 5.56, 11.63, 1.58, bg_rgb=C_DARK_BG, border_rgb=None)
    tf26_d = c26_disc.text_frame
    p26_1 = tf26_d.paragraphs[0]
    r26_1 = p26_1.add_run()
    r26_1.text = "การอภิปรายผลสรุปภาพรวมของโครงงาน (Overall Project Discussion):"
    r26_1.font.name = FONT_THAI
    r26_1.font.size = Pt(18.0)
    r26_1.font.bold = True
    r26_1.font.color.rgb = C_ORANGE_TAG

    p26_2 = tf26_d.add_paragraph()
    p26_2.space_before = Pt(2)
    r26_2 = p26_2.add_run()
    r26_2.text = (
        "• ด้านฮาร์ดแวร์และทัศนศาสตร์ (ขั้นตอนที่ 1–4): พิสูจน์จุดตัดแสงด้วย Ray-Tracing ออกแบบเคสกระจกแซฟไฟร์ และสร้างวงจร OPA380 TIA + ADS1220 24 บิตอ่านค่าจริงสำเร็จ\n"
        "• ด้านการชดเชยและระบบข้อมูล (ขั้นตอนที่ 5–6): เก็บกราฟมุมตกกระทบจริง 0°–85° ในแล็บฟิสิกส์ เพื่อนำไปชดเชยร่วมกับพิกัดเวลาและอุณหภูมิในระบบ Data Logger & Web"
    )
    r26_2.font.name = FONT_THAI
    r26_2.font.size = Pt(16.5)
    r26_2.font.color.rgb = C_WHITE

    # -------------------------------------------------------------------------
    # SLIDE 27: Closing Slide (Dark Navy Theme)
    # -------------------------------------------------------------------------
    s27 = new_slide(C_DARK_BG, show_num=False)
    bar_end = s27.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.36), Inches(2.6), Inches(0.6), Inches(0.06))
    bar_end.fill.solid()
    bar_end.fill.fore_color.rgb = RGBColor(249, 115, 22)
    bar_end.line.fill.background()

    tb_end = s27.shapes.add_textbox(Inches(1.0), Inches(2.85), Inches(11.33), Inches(3.0))
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
    re3.text = "เตชินท์ รัฐวงษา (66010296)  •  ณัฐพัชร์ รัตนนวงศ์ (66010256)  •  พีรวิช เฉลิมเกียรติ (66010591)"
    re3.font.name = FONT_THAI
    re3.font.size = Pt(19)
    re3.font.color.rgb = RGBColor(203, 213, 225)

    out_path = os.path.join(script_dir, "69_IoT_Pyranometer_Capstone1_Slides_V2_Clean7Min.pptx")
    prs.save(out_path)
    print(f"[OK] Saved new V2 presentation (without overwriting original) to: {out_path}")
    return out_path


if __name__ == "__main__":
    create_presentation_v2()
