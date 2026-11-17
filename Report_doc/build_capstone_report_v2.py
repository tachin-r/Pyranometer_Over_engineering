"""
========================================================================================
สคริปต์คอมไพล์เล่มรายงานความก้าวหน้าโครงงานปริญญานิพนธ์ 1 (Capstone Project 1 Progress Report)
โครงงาน: การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (High-Precision Pyranometer)
หลักสูตรสองปริญญา: วศ.บ. (วิศวกรรมระบบไอโอทีและสารสนเทศ) และ วท.บ. (ฟิสิกส์อุตสาหกรรม)
คณะวิศวกรรมศาสตร์ สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง (สจล.)

ขอบเขตรายงาน: 5 บทหลัก (บทที่ 1 บทนำ, บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง, บทที่ 3 การออกแบบและวิธีดำเนินงาน, บทที่ 4 ผลการจำลองและทดสอบต้นแบบ, บทที่ 5 สรุปผล ความก้าวหน้า และแผนการดำเนินงานต่อยอด)
จัดรูปแบบตามมาตรฐานวิทยานิพนธ์ สจล. (KMITL Thesis Format 100%)
========================================================================================
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT

# นำเข้า Helper Functions
sys.path.append(os.path.dirname(__file__))
from helpers import (
    setup_kmitl_page,
    set_font_sarabun,
    add_page_number_to_footer,
    add_page_number_to_header,
    set_outline_level,
    set_section_numbering,
    add_thesis_chapter_heading,
    add_section_heading,
    add_subsection_heading,
    add_thai_body_p,
    add_numbered_item,
    add_bullet_item,
    add_toc_line,
    set_table_clean_borders,
    set_table_no_borders,
    set_cell_shading,
    set_cell_margins,
    add_table_caption,
    add_figure_image,
    add_two_figure_images_side_by_side,
    add_two_figure_images_stacked,
    add_photo_placeholder,
    add_equation_box
)

def create_report_v2():
    doc = docx.Document()
    script_dir = os.path.dirname(__file__)
    img_dir = os.path.join(script_dir, "image")

    # ตั้งค่า Base Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'TH SarabunPSK'
    normal_style.font.size = Pt(16)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)

    # =========================================================================
    # SECTION 1: ปกนอก ปกใน และหน้ารับรองปริญญานิพนธ์ (ไม่มีเลขหน้า ตามแบบ 69_IoT_66011392.pdf)
    # =========================================================================
    sec1 = doc.sections[0]
    setup_kmitl_page(sec1)

    # 1.1 หน้าปกภาษาไทย (รูปแบบเดียวกับหน้า 1 ของ 69_IoT_66011392.pdf)
    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(24)

    p_th_title = doc.add_paragraph()
    p_th_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_th_title.paragraph_format.space_after = Pt(12)
    p_th_title.paragraph_format.line_spacing = 1.2
    r_th = p_th_title.add_run(
        "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"
    )
    set_font_sarabun(r_th, size_pt=18, bold=True)

    p_en_title = doc.add_paragraph()
    p_en_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_title.paragraph_format.space_after = Pt(72)
    p_en_title.paragraph_format.line_spacing = 1.2
    r_en = p_en_title.add_run(
        "DEVELOPMENT OF HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM"
    )
    set_font_sarabun(r_en, size_pt=16, bold=True)

    p_authors = doc.add_paragraph()
    p_authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_authors.paragraph_format.space_after = Pt(80)
    p_authors.paragraph_format.line_spacing = 1.25
    r_auth = p_authors.add_run(
        "เตชินท์   รัฐวงษา\n"
        "TECHIN   RATTAWONGSA\n"
        "ณัฐพัชร์  รัตนนวงศ์\n"
        "NATTAPAT  RATANANWONG\n"
        "พีรวิช   เฉลิมเกียรติ\n"
        "PEERAWICH CHALOEMKIAT"
    )
    set_font_sarabun(r_auth, size_pt=16, bold=False)

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.paragraph_format.line_spacing = 1.2
    r_foot = p_footer.add_run(
        "ปริญญานิพนธ์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต\n"
        "สาขาวิชาวิศวกรรมระบบไอโอทีและสารสนเทศ\n"
        "และปริญญาวิทยาศาสตรบัณฑิตสาขาวิชาฟิสิกส์อุตสาหกรรม\n"
        "ภาควิชาวิศวกรรมไอโอทีและสารสนเทศ\n"
        "คณะวิศวกรรมศาสตร์\n"
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง\n"
        "ปีการศึกษา 2569"
    )
    set_font_sarabun(r_foot, size_pt=16, bold=True)

    doc.add_page_break()

    # 1.2 หน้าปกภาษาอังกฤษ (รูปแบบเดียวกับหน้า 2 ของ 69_IoT_66011392.pdf พร้อมแก้คำสะกด SCIENCE ให้ถูกต้อง)
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(36)

    p_en_cov = doc.add_paragraph()
    p_en_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_cov.paragraph_format.space_after = Pt(90)
    p_en_cov.paragraph_format.line_spacing = 1.2
    r_encov = p_en_cov.add_run(
        "DEVELOPMENT OF HIGH-PRECISION SOLAR IRRADIANCE\n"
        "MEASUREMENT SYSTEM"
    )
    set_font_sarabun(r_encov, size_pt=16, bold=True)

    p_auth_en = doc.add_paragraph()
    p_auth_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth_en.paragraph_format.space_after = Pt(90)
    p_auth_en.paragraph_format.line_spacing = 1.25
    r_authen = p_auth_en.add_run(
        "TECHIN   RATTAWONGSA\n"
        "NATTAPAT  RATANANWONG\n"
        "PEERAWICH CHALOEMKIAT"
    )
    set_font_sarabun(r_authen, size_pt=16, bold=False)

    p_en_sub = doc.add_paragraph()
    p_en_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_sub.paragraph_format.line_spacing = 1.2
    r_ensub = p_en_sub.add_run(
        "A THESIS IS SUBMITTED IN PARTIAL FULFILLMENT\n"
        "OF THE REQUIREMENT FOR THE DEGREE OF\n"
        "BACHELOR OF ENGINEERING IN\n"
        "IOT SYSTEM AND INFORMATION ENGINEERING AND\n"
        "BACHELOR OF SCIENCE IN INDUSTRIAL PHYSICS\n"
        "DEPARTMENT OF IoT AND INFORMATION ENGINEERING\n"
        "SCHOOL OF ENGINEERING\n"
        "KING MONGKUT'S INSTITUTE OF TECHNOLOGY LADKRABANG\n"
        "ACADEMIC YEAR 2026"
    )
    set_font_sarabun(r_ensub, size_pt=15, bold=True)

    doc.add_page_break()

    # 1.3 หน้ารับรองปริญญานิพนธ์ (รูปแบบเดียวกับหน้า 3 ของ 69_IoT_66011392.pdf และตารางลายมือชื่อตามรูปที่กำหนด)
    t_app_info = doc.add_table(rows=6, cols=2)
    t_app_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t_app_info)
    col_widths = [Inches(1.8), Inches(4.4)]
    for row in t_app_info.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    app_data = [
        ("หัวข้อปริญญานิพนธ์", "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"),
        ("Thesis Title", "DEVELOPMENT OF HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM"),
        ("ชื่อนักศึกษา", "นายเตชินท์  รัฐวงษา       รหัสนักศึกษา 66010296\nนายณัฐพัชร์ รัตนนวงศ์      รหัสนักศึกษา 66010256\nนายพีรวิช  เฉลิมเกียรติ    รหัสนักศึกษา 66010591"),
        ("ระดับปริญญา", "วิศวกรรมศาสตรบัณฑิตและวิทยาศาสตรบัณฑิต"),
        ("สาขาวิชา", "วิศวกรรมระบบไอโอทีและสารสนเทศ และฟิสิกส์อุตสาหกรรม"),
        ("ปริญญานิพนธ์ปีการศึกษา", "2569")
    ]

    for idx, (label, val) in enumerate(app_data):
        c0 = t_app_info.cell(idx, 0)
        c1 = t_app_info.cell(idx, 1)
        set_cell_margins(c0, top=40, bottom=40, left=0, right=60)
        set_cell_margins(c1, top=40, bottom=40, left=40, right=0)
        r0 = c0.paragraphs[0].add_run(label)
        set_font_sarabun(r0, size_pt=16, bold=True)
        r1 = c1.paragraphs[0].add_run(val)
        set_font_sarabun(r1, size_pt=16)

    p_app_body = doc.add_paragraph()
    p_app_body.paragraph_format.first_line_indent = Inches(0.5)
    p_app_body.paragraph_format.space_before = Pt(14)
    p_app_body.paragraph_format.space_after = Pt(14)
    p_app_body.paragraph_format.line_spacing = 1.2
    r_app_txt = p_app_body.add_run(
        "คณะวิศวกรรมศาสตร์ สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง "
        "อนุมัติให้ปริญญานิพนธ์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต "
        "(วิศวกรรมระบบไอโอทีและสารสนเทศ) และวิทยาศาสตรบัณฑิต (ฟิสิกส์อุตสาหกรรม) ประจำปีการศึกษา 2569"
    )
    set_font_sarabun(r_app_txt, size_pt=16)

    # ตารางคณะกรรมการสอบ (ตรงตามรูปตัวอย่าง 100%: พื้นขาวล้วน ช่องลายมือชื่อเว้นว่างให้กรรมการเซ็นจริง)
    t_comm = doc.add_table(rows=6, cols=2)
    t_comm.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_comm, border_color="000000", border_sz="6")
    for r in t_comm.rows:
        r.cells[0].width = Inches(3.0)
        r.cells[1].width = Inches(3.0)

    c_th0 = t_comm.cell(0, 0)
    c_th1 = t_comm.cell(0, 1)
    set_cell_margins(c_th0, top=60, bottom=60, left=80, right=80)
    set_cell_margins(c_th1, top=60, bottom=60, left=80, right=80)
    c_th0.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c_th1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_th0 = c_th0.paragraphs[0].add_run("คณะกรรมการสอบ")
    set_font_sarabun(r_th0, size_pt=16, bold=True)
    r_th1 = c_th1.paragraphs[0].add_run("ลายมือชื่อ")
    set_font_sarabun(r_th1, size_pt=16, bold=True)

    comm_members = [
        "ประธานกรรมการสอบ",
        "กรรมการ",
        "กรรมการ",
        "กรรมการและอาจารย์ที่ปรึกษา",
        "กรรมการและอาจารย์ที่ปรึกษาร่วม"
    ]

    for idx, role in enumerate(comm_members, start=1):
        cell_role = t_comm.cell(idx, 0)
        cell_sig = t_comm.cell(idx, 1)
        set_cell_margins(cell_role, top=220, bottom=80, left=80, right=80)
        set_cell_margins(cell_sig, top=220, bottom=80, left=80, right=80)
        cell_role.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_role = cell_role.paragraphs[0].add_run(role)
        set_font_sarabun(r_role, size_pt=16)
        r_sig = cell_sig.paragraphs[0].add_run("")
        set_font_sarabun(r_sig, size_pt=16)

    # =========================================================================
    # SECTION 2: ส่วนนำ (บทคัดย่อ, ABSTRACT, กิตติกรรมประกาศ, สารบัญ) -> เลขหน้าโรมัน i, ii, iii...
    # =========================================================================
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec2)
    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False
    set_section_numbering(sec2, fmt="lowerRoman", start=1)
    add_page_number_to_footer(sec2.footer)

    # 2.1 บทคัดย่อภาษาไทย (หน้า i - มีบล็อกข้อมูลหัวเรื่องเหมือนหน้า 4 ของ 69_IoT_66011392.pdf)
    t_abs_th = doc.add_table(rows=8, cols=2)
    t_abs_th.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t_abs_th)
    for row in t_abs_th.rows:
        row.cells[0].width = Inches(2.0)
        row.cells[1].width = Inches(4.2)

    abs_th_meta = [
        ("หัวข้อปริญญานิพนธ์", "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"),
        ("Thesis Title", "DEVELOPMENT OF HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM"),
        ("ชื่อนักศึกษา", "นายเตชินท์  รัฐวงษา       รหัสนักศึกษา 66010296\nนายณัฐพัชร์ รัตนนวงศ์      รหัสนักศึกษา 66010256\nนายพีรวิช  เฉลิมเกียรติ    รหัสนักศึกษา 66010591"),
        ("ระดับปริญญา", "วิศวกรรมศาสตรบัณฑิต และวิทยาศาสตรบัณฑิต"),
        ("สาขาวิชา", "วิศวกรรมระบบไอโอทีและสารสนเทศ และฟิสิกส์อุตสาหกรรม"),
        ("ปีการศึกษา", "2569"),
        ("อาจารย์ที่ปรึกษาปริญญานิพนธ์", "ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ"),
        ("อาจารย์ที่ปรึกษาปริญญานิพนธ์ร่วม", "ผศ.ดร.ประธาน บูรณศิริ")
    ]
    for idx, (label, val) in enumerate(abs_th_meta):
        c0 = t_abs_th.cell(idx, 0)
        c1 = t_abs_th.cell(idx, 1)
        set_cell_margins(c0, top=25, bottom=25, left=0, right=40)
        set_cell_margins(c1, top=25, bottom=25, left=20, right=0)
        r0 = c0.paragraphs[0].add_run(label)
        set_font_sarabun(r0, size_pt=15, bold=True)
        r1 = c1.paragraphs[0].add_run(val)
        set_font_sarabun(r1, size_pt=15)

    p_abs_title = doc.add_paragraph()
    p_abs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs_title.paragraph_format.space_before = Pt(14)
    p_abs_title.paragraph_format.space_after = Pt(10)
    r_abst = p_abs_title.add_run("บทคัดย่อ")
    set_font_sarabun(r_abst, size_pt=18, bold=True)
    set_outline_level(p_abs_title, 0)

    add_thai_body_p(
        doc,
        "ปริญญานิพนธ์นี้นำเสนอการวิจัย ออกแบบ และสร้างต้นแบบระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง "
        "(High-Precision Solar Irradiance Measurement System) โดยมุ่งเน้นการแก้ปัญหาความคลาดเคลื่อนทางทัศนศาสตร์เชิงมุม (Cosine Error) "
        "และเสถียรภาพของวงจรตรวจวัดสัญญาณแอนะล็อกความละเอียดสูง ในขั้นแรกได้ทำการวิเคราะห์เปรียบเทียบเซนเซอร์โฟโตไดโอดชนิดซิลิคอนระหว่างรุ่น BPX 61 "
        "(ตัวถังโลหะ TO-39) และรุ่น BPW 34 (ตัวถังพลาสติกแบนราบ) ซึ่งพบว่าทั้งสองรุ่นใช้ชิปซิลิคอน (Silicon Die) ขนาดเดียวกันคือ 2.65 × 2.65 mm² "
        "(พื้นที่รับแสง 7.02 mm²) และมีคุณสมบัติทางไฟฟ้าเหมือนกันทุกประการ แต่ผลจากการจำลองทางทัศนศาสตร์ (Optical Ray-tracing Simulation) "
        "พบว่าเมื่อใส่กระจกหน้าต่าง Schott 8250 บนกระป๋อง TO-39 การหักเหของแสงทำให้ขอบกระป๋องบังแสงเร็วขึ้นจนเกิดจุดตัดแสงตั้งแต่ 44 องศา "
        "โครงงานจึงเลือกใช้เซนเซอร์ BPW 34 ที่มีระนาบรับแสงแบนราบร่วมกับหน้าต่างกระจกแซฟไฟร์ใส (Sapphire Optical Window)"
    )
    add_thai_body_p(
        doc,
        "ในภาคฮาร์ดแวร์และแบบจำลองคณิตศาสตร์ ได้ออกแบบวงจรขยายกระแสเป็นแรงดัน (Transimpedance Amplifier: TIA) ระบบไฟเลี้ยงเดี่ยวด้วยไอซี OPA380 "
        "โดยป้อนแรงดันไบแอสอ้างอิง V_bias (0.409 V หรือปรับลดเป็น 0.144–0.202 V) ที่ขา Non-Inverting (+IN) เพื่อยกจุดทำงานให้พ้นย่านอิ่มตัวขอบล่าง 100 mV "
        "ของออปแอมป์โดยเกิดกระแสรั่วไหลต่ำเพียง 30–84 pA (ต่ำกว่าการไบแอส 10 V ถึง 24–65 เท่า) พร้อมทั้งวิเคราะห์การเลือกค่าตัวต้านทานป้อนกลับ R1 "
        "ตั้งแต่ 2,200 Ω ไปจนถึง 390 Ω เพื่อขยายเพดานการวัดภายใต้กระจกแซฟไฟร์ใสให้รองรับความเข้มรังสีสูงสุดถึง 4,000 W/m² เทียบเท่าเครื่องมือวัดมาตรฐาน "
        "ISO 9060 Class A โดยให้ค่าความไว ~1 mV ต่อ W/m² ร่วมกับการส่งข้อมูลดิบ 4 ช่องสัญญาณจาก 24-bit ADC (ADS1220), "
        "แหล่งจ่ายแรงดันอ้างอิง REF5045 และหัววัดอุณหภูมิ NTC Thermistor (TTC3A103F39HAEY) ผ่านไมโครคอนโทรลเลอร์ ESP32-C6 ไปประมวลผลสมการชดเชยบนตัว Data Logger"
    )

    doc.add_page_break()

    # 2.2 บทคัดย่อภาษาอังกฤษ (หน้า ii - มีบล็อกข้อมูลหัวเรื่องเหมือนหน้า 5 ของ 69_IoT_66011392.pdf)
    t_abs_en = doc.add_table(rows=7, cols=2)
    t_abs_en.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_no_borders(t_abs_en)
    for row in t_abs_en.rows:
        row.cells[0].width = Inches(1.8)
        row.cells[1].width = Inches(4.4)

    abs_en_meta = [
        ("Thesis Title", "DEVELOPMENT OF HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM"),
        ("Student", "Mr. Techin Rattawongsa        Student ID. 66010296\nMr. Nattapat Ratananwong      Student ID. 66010256\nMr. Peerawich Chaloemkiat     Student ID. 66010591"),
        ("Degree", "Bachelor of Engineering and Bachelor of Science"),
        ("Program", "IoT System and Information Engineering and Industrial Physics"),
        ("Academic Year", "2026"),
        ("Thesis Advisor", "Asst. Prof. Dr. Thanawit Anuwongpinit"),
        ("Thesis Co-Advisor", "Asst. Prof. Dr. Prathan Buranasiri")
    ]
    for idx, (label, val) in enumerate(abs_en_meta):
        c0 = t_abs_en.cell(idx, 0)
        c1 = t_abs_en.cell(idx, 1)
        set_cell_margins(c0, top=25, bottom=25, left=0, right=40)
        set_cell_margins(c1, top=25, bottom=25, left=20, right=0)
        r0 = c0.paragraphs[0].add_run(label)
        set_font_sarabun(r0, size_pt=15, bold=True)
        r1 = c1.paragraphs[0].add_run(val)
        set_font_sarabun(r1, size_pt=15)

    p_absen_title = doc.add_paragraph()
    p_absen_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_absen_title.paragraph_format.space_before = Pt(14)
    p_absen_title.paragraph_format.space_after = Pt(10)
    r_absen = p_absen_title.add_run("ABSTRACT")
    set_font_sarabun(r_absen, size_pt=18, bold=True)
    set_outline_level(p_absen_title, 0)

    add_thai_body_p(
        doc,
        "This thesis presents the research, design, and prototyping of a high-precision solar irradiance measurement system, "
        "focusing on mitigating directional cosine error and front-end analog instrumentation limitations. Comparative analysis between the BPX 61 (TO-39 metal can) "
        "and BPW 34 (flat clear plastic package) silicon PIN photodiodes revealed that both devices share the exact same silicon die (2.65 × 2.65 mm², 7.02 mm² active area). "
        "Optical ray-tracing simulations demonstrated that introducing the Schott 8250 glass window on the TO-39 package causes premature optical cutoff at 44° due to refractive bending. "
        "Consequently, the flat BPW 34 package was selected and paired with a synthetic sapphire optical window."
    )
    add_thai_body_p(
        doc,
        "In the analog front-end (AFE), a single-supply transimpedance amplifier (TIA) was designed around the auto-zero OPA380 operational amplifier. "
        "Applying a controlled low positive bias voltage V_bias (0.409 V, or optimized to 0.144–0.202 V) at the non-inverting input lifts the zero-light baseline above the 100 mV "
        "negative-rail saturation floor while keeping dark current at just 30–84 pA (24–65 times lower than conventional 10 V biasing). Furthermore, scaling the feedback resistor R1 "
        "from 2,200 Ω down to 390 Ω extends the full-scale unattenuated irradiance range under the clear sapphire window up to 4,000 W/m² (matching flagship ISO 9060 Class A specifications "
        "with a 1 mV/(W/m²) sensitivity and 0.002 W/m² resolution via the 24-bit ADS1220 ADC). A complete 4-stage mathematical compensation pipeline integrating real-time differential "
        "baseline subtraction, dark-current thermal modeling, photodiode temperature coefficient correction (+0.18%/K via TTC3A103F39HAEY NTC), and REF5045 thermal drift correction "
        "is transmitted via an ESP32-C6 microcontroller to a Data Logger for real-time compensation, solar angle correction, and photovoltaic performance auditing."
    )

    doc.add_page_break()

    # 2.3 กิตติกรรมประกาศ (หน้า iii)
    p_ack_title = doc.add_paragraph()
    p_ack_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ack_title.paragraph_format.space_before = Pt(10)
    p_ack_title.paragraph_format.space_after = Pt(12)
    r_ack = p_ack_title.add_run("กิตติกรรมประกาศ")
    set_font_sarabun(r_ack, size_pt=18, bold=True)
    set_outline_level(p_ack_title, 0)

    add_thai_body_p(
        doc,
        "ปริญญานิพนธ์ฉบับนี้สำเร็จลุล่วงไปได้ด้วยดี เนื่องจากความเมตตาและความช่วยเหลืออย่างดียิ่งจากหลายภาคส่วน "
        "คณะผู้ดำเนินโครงงานขอกราบขอบพระคุณ ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ อาจารย์ที่ปรึกษาปริญญานิพนธ์ และ ผศ.ดร.ประธาน บูรณศิริ "
        "อาจารย์ที่ปรึกษาปริญญานิพนธ์ร่วม ตลอดจนคณาจารย์ภาควิชาวิศวกรรมไอโอทีและสารสนเทศ คณะวิศวกรรมศาสตร์ และภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ "
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง ที่ได้กรุณาถ่ายทอดองค์ความรู้ ให้คำแนะนำเชิงวิชาการอย่างลึกซึ้ง "
        "ทั้งในด้านการออกแบบวงจรอิเล็กทรอนิกส์ความแม่นยำสูงและการสนับสนุนเครื่องมือวัดทางทัศนศาสตร์ระดับห้องปฏิบัติการ"
    )
    add_thai_body_p(
        doc,
        "ขอขอบคุณคณาจารย์และเจ้าหน้าที่ห้องปฏิบัติการทัศนศาสตร์ ภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล. "
        "ที่ได้ให้ความอนุเคราะห์ในการเปิดให้ยืมใช้เครื่องตรวจวัดสเปกตรัม Vernier LabQuest 2 และอุปกรณ์ชุดแท่นหมุนวัดองศา "
        "ตลอดจนให้คำปรึกษาในการออกแบบระเบียบวิธีการทดลองเพื่อประเมินคุณลักษณะทางแสงของเซนเซอร์และกระจกแซฟไฟร์ในโครงงานนี้ "
        "ท้ายที่สุดนี้ ขอกราบขอบพระคุณบิดามารดาและครอบครัวที่เป็นกำลังใจสำคัญในการศึกษาเสมอมา"
    )

    p_ack_sign = doc.add_paragraph()
    p_ack_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ack_sign.paragraph_format.space_before = Pt(16)
    r_asign = p_ack_sign.add_run(
        "คณะผู้จัดทำปริญญานิพนธ์\n"
        "เตชินท์  รัฐวงษา\n"
        "ณัฐพัชร์ รัตนนวงศ์\n"
        "พีรวิช  เฉลิมเกียรติ"
    )
    set_font_sarabun(r_asign, size_pt=16, bold=False)

    doc.add_page_break()

    # 2.4 สารบัญ (Table of Contents - หน้า iv)
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_head.paragraph_format.space_before = Pt(10)
    p_toc_head.paragraph_format.space_after = Pt(12)
    r_toc = p_toc_head.add_run("สารบัญ")
    set_font_sarabun(r_toc, size_pt=18, bold=True)
    set_outline_level(p_toc_head, 0)

    add_toc_line(doc, "บทคัดย่อ", "i")
    add_toc_line(doc, "ABSTRACT", "ii")
    add_toc_line(doc, "กิตติกรรมประกาศ", "iii")
    add_toc_line(doc, "สารบัญ", "iv")
    add_toc_line(doc, "สารบัญตาราง", "v")
    add_toc_line(doc, "สารบัญรูป", "vi")
    add_toc_line(doc, "บทที่ 1 บทนำ", "1", is_bold=True)
    add_toc_line(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา", "1", level=1)
    add_toc_line(doc, "1.2 วัตถุประสงค์ของโครงงาน", "2", level=1)
    add_toc_line(doc, "1.3 ขอบเขตของโครงงาน", "2", level=1)
    add_toc_line(doc, "1.4 ประโยชน์ที่คาดว่าจะได้รับ", "3", level=1)
    add_toc_line(doc, "1.5 แผนการดำเนินงานโครงงาน", "3", level=1)
    add_toc_line(doc, "บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "5", is_bold=True)
    add_toc_line(doc, "2.1 มาตรฐานการวัดรังสีดวงอาทิตย์ (ISO 9060:2018)", "5", level=1)
    add_toc_line(doc, "2.2 การวิเคราะห์เปรียบเทียบชิปเซนเซอร์ BPX 61 และ BPW 34", "7", level=1)
    add_toc_line(doc, "2.3 ทฤษฎีโฟโตไดโอด วงจรขยาย TIA และการออกแบบพิกัด 4,000 W/m²", "10", level=1)
    add_toc_line(doc, "2.4 คุณสมบัติทางแสงของกระจกแซฟไฟร์ (Sapphire Optical Window)", "16", level=1)
    add_toc_line(doc, "2.5 ไอซีแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)", "17", level=1)
    add_toc_line(doc, "2.6 แหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูง (REF5045) และ LDO (TPS7A4901)", "18", level=1)
    add_toc_line(doc, "บทที่ 3 การออกแบบและวิธีดำเนินงาน", "20", is_bold=True)
    add_toc_line(doc, "3.1 สถาปัตยกรรมระบบวัด การแมปช่องสัญญาณ Perfboard vs PCB และเฟิร์มแวร์", "20", level=1)
    add_toc_line(doc, "3.2 การออกแบบโครงสร้าง 3D CAD เคสกระจกแซฟไฟร์ และแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน", "24", level=1)
    add_toc_line(doc, "3.3 การออกแบบข้อเสนอระเบียบวิธีการทดลองจริงทางทัศนศาสตร์", "26", level=1)
    add_toc_line(doc, "บทที่ 4 ผลการจำลองและทดสอบต้นแบบ", "29", is_bold=True)
    add_toc_line(doc, "4.1 ผลการจำลองทางทัศนศาสตร์และการค้นพบความผิดเพี้ยนจากกระจกหน้าต่าง", "29", level=1)
    add_toc_line(doc, "4.2 ผลการสร้างบอร์ดต้นแบบ Perfboard และการแก้ปัญหาภาคจ่ายไฟจริง", "32", level=1)
    add_toc_line(doc, "4.3 ผลการทดสอบอ่านสัญญาณจริง 24 บิตใน 4 สภาวะแสง", "35", level=1)
    add_toc_line(doc, "4.4 ต้นแบบหน้าต่างแสดงผลบนเว็บ (Web Dashboard Mockup)", "37", level=1)
    add_toc_line(doc, "4.5 ขอบเขตงานทดสอบที่อยู่ระหว่างดำเนินการในโครงงาน 2", "38", level=1)
    add_toc_line(doc, "บทที่ 5 สรุปผล ความก้าวหน้า และแผนการดำเนินงานต่อยอด", "39", is_bold=True)
    add_toc_line(doc, "5.1 สรุปสถานะความก้าวหน้าตามแผนงาน 6 ขั้นตอนหลัก", "39", level=1)
    add_toc_line(doc, "5.2 อภิปรายผลภาพรวมและแนวทางการดำเนินงานสู่โครงงาน 2", "41", level=1)
    add_toc_line(doc, "เอกสารอ้างอิง", "42", is_bold=True)
    add_toc_line(doc, "ภาคผนวก ก รายการอุปกรณ์ต้นแบบ (BOM)", "43", is_bold=True)
    add_toc_line(doc, "ภาคผนวก ข เอกสารมาตรฐานสากลอ้างอิง ISO 9060:2018 (Table 1)", "45", is_bold=True)

    doc.add_page_break()

    # 2.5 สารบัญตาราง (List of Tables - หน้า v)
    p_lot_head = doc.add_paragraph()
    p_lot_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lot_head.paragraph_format.space_before = Pt(10)
    p_lot_head.paragraph_format.space_after = Pt(12)
    r_lot = p_lot_head.add_run("สารบัญตาราง")
    set_font_sarabun(r_lot, size_pt=18, bold=True)
    set_outline_level(p_lot_head, 0)

    p_col_t = doc.add_paragraph()
    p_col_t.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), WD_TAB_ALIGNMENT.RIGHT)
    r_ct1 = p_col_t.add_run("ตารางที่")
    set_font_sarabun(r_ct1, size_pt=16, bold=True)
    r_ct2 = p_col_t.add_run("\tหน้า")
    set_font_sarabun(r_ct2, size_pt=16, bold=True)

    add_toc_line(doc, "ตารางที่ 1.1 แผนการดำเนินงานและกิจกรรมโครงงานตลอดปีการศึกษา", "4")
    add_toc_line(doc, "ตารางที่ 2.1 เกณฑ์มาตรฐานการจำแนกประเภท Pyranometer ตาม ISO 9060:2018", "5")
    add_toc_line(doc, "ตารางที่ 2.2 การเปรียบเทียบพารามิเตอร์ทางเทคนิคระหว่าง OSRAM BPX 61 กับ BPW 34", "7")
    add_toc_line(doc, "ตารางที่ 2.3 การเปรียบเทียบรูปแบบการไบแอสขา +IN (V_bias) ในวงจร TIA ไฟเลี้ยงเดี่ยว", "12")
    add_toc_line(doc, "ตารางที่ 2.4 การออกแบบค่าตัวต้านทานป้อนกลับ R1 (R_TIA) ตามพิกัดความเข้มแสงสูงสุดถึง 4,000 W/m²", "14")
    add_toc_line(doc, "ตารางที่ 3.1 ตารางกำหนดช่องสัญญาณ Input MUX ของ ADS1220 (เปรียบเทียบบอร์ด Perfboard และ PCB)", "21")
    add_toc_line(doc, "ตารางที่ 3.2 ตารางกำหนดขาสัญญาณ SPI ระหว่าง ESP32-C6 กับ ADS1220", "22")
    add_toc_line(doc, "ตารางที่ 4.1 ตารางเปรียบเทียบผลการจำลองเชิงมุม: BPX 61 vs BPW 34", "31")
    add_toc_line(doc, "ตารางที่ 4.2 ตารางเปรียบเทียบเสถียรภาพแรงดันอ้างอิง REF5045 และการวัดแรงดันจริงด้วยมัลติมิเตอร์", "34")
    add_toc_line(doc, "ตารางที่ 4.3 ข้อมูลการอ่านค่าแรงดันจริงภายใต้สภาวะแสงต่างๆ จากบอร์ดต้นแบบ Perfboard", "36")
    add_toc_line(doc, "ตารางที่ 5.1 สรุปสถานะความก้าวหน้าการดำเนินงานโครงงาน (Roadmap 6 ขั้นตอนหลัก)", "40")

    doc.add_page_break()

    # 2.6 สารบัญรูป (List of Figures - หน้า vi)
    p_lof_head = doc.add_paragraph()
    p_lof_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lof_head.paragraph_format.space_before = Pt(10)
    p_lof_head.paragraph_format.space_after = Pt(12)
    r_lof = p_lof_head.add_run("สารบัญรูป")
    set_font_sarabun(r_lof, size_pt=18, bold=True)
    set_outline_level(p_lof_head, 0)

    p_col_f = doc.add_paragraph()
    p_col_f.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), WD_TAB_ALIGNMENT.RIGHT)
    r_cf1 = p_col_f.add_run("รูปที่")
    set_font_sarabun(r_cf1, size_pt=16, bold=True)
    r_cf2 = p_col_f.add_run("\tหน้า")
    set_font_sarabun(r_cf2, size_pt=16, bold=True)

    add_toc_line(doc, "รูปที่ 2.1 นิยามและสมการการตอบสนองเชิงมุม (Cosine Error) จากเอกสาร ISO 9060:2018 (Annex A.5)", "6")
    add_toc_line(doc, "รูปที่ 2.2 คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) ของ BPX 61 เทียบกับ BPW 34", "8")
    add_toc_line(doc, "รูปที่ 2.3 มิติและโครงสร้างกายภาพตัดขวางของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34", "9")
    add_toc_line(doc, "รูปที่ 2.4 คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) และกราฟ Dark Current จาก Datasheet BPW 34", "11")
    add_toc_line(doc, "รูปที่ 2.5 ค่าพารามิเตอร์อุณหภูมิ (TC_I = 0.18 %/K) และความจุ C_0 = 72 pF จาก Datasheet BPW 34", "11")
    add_toc_line(doc, "รูปที่ 2.6 ผังวงจรจริงของภาคขยายกระแสแสง Transimpedance Amplifier (TIA) ด้วยไอซี OPA380", "13")
    add_toc_line(doc, "รูปที่ 2.7 กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Transmission)", "16")
    add_toc_line(doc, "รูปที่ 2.8 บล็อกไดอะแกรมโครงสร้างวงจรภายในของไอซี 24-bit ADC (ADS1220)", "17")
    add_toc_line(doc, "รูปที่ 2.9 วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR", "19")
    add_toc_line(doc, "รูปที่ 3.1 สถาปัตยกรรมระบบวัดและผังวงจรการเชื่อมต่อไอซี ADS1220xPW ใน Schematic จริง", "23")
    add_toc_line(doc, "รูปที่ 3.2 แบบจำลอง 3D CAD โครงสร้างภายในของเซนเซอร์ BPX 61 (TO-39) และ BPW 34 (DIL)", "24")
    add_toc_line(doc, "รูปที่ 3.3 แบบจำลอง 3D CAD ของกล่องเคสสำหรับบอร์ด Perfboard และแหวนล็อกกระจกแซฟไฟร์", "25")
    add_toc_line(doc, "รูปที่ 3.4 ภาพจำลอง 3 มิติของแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน (BPX 61 vs BPW 34)", "25")
    add_toc_line(doc, "รูปที่ 3.5 หน้าจอแสดงผลการวัดสเปกตรัมของชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์", "27")
    add_toc_line(doc, "รูปที่ 3.6 แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุม (Goniometer Cosine Test Jig)", "28")
    add_toc_line(doc, "รูปที่ 4.1 ผลการจำลองการตอบสนองเชิงมุมของ BPX 61: กรณีไม่มีกระจก เทียบกับ มีกระจก SCHOTT 8250", "30")
    add_toc_line(doc, "รูปที่ 4.2 ผลการจำลองการตอบสนองเชิงมุมและ Heatmap ของ BPW 34 พร้อมการประเมิน ISO 9060", "31")
    add_toc_line(doc, "รูปที่ 4.3 ภาพถ่ายจริงของบอร์ดต้นแบบ Perfboard (70×90 mm) ด้านบน ด้านข้าง และด้านล่าง", "32")
    add_toc_line(doc, "รูปที่ 4.4 ภาพถ่ายผลการวัดแรงดันภาคจ่ายไฟและแรงดันอ้างอิงจริงด้วยมัลติมิเตอร์ (5.194 V & 4.495 V)", "34")
    add_toc_line(doc, "รูปที่ 4.5 ต้นแบบหน้าต่างแสดงผลบนเว็บ (Web Dashboard Mockup) ทั้ง 3 มุมมอง", "37")
    add_toc_line(doc, "รูปที่ ข.1 ตารางเกณฑ์การจำแนกประเภท Pyranometer จากเอกสารมาตรฐานสากล ISO 9060:2018 (Table 1)", "45")

    # =========================================================================
    # SECTION 3: เนื้อหาหลัก บทที่ 1 ถึง 3 (เลขหน้าอารบิก 1, 2, 3... มุมบนขวา ตามแบบ 69_IoT_66011392.pdf)
    # =========================================================================
    sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec3)
    sec3.header.is_linked_to_previous = False
    sec3.footer.is_linked_to_previous = False
    set_section_numbering(sec3, fmt="decimal", start=1)
    add_page_number_to_header(sec3.header, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # -------------------------------------------------------------------------
    # บทที่ 1: บทนำ
    # -------------------------------------------------------------------------
    add_thesis_chapter_heading(doc, "1", "บทนำ")

    add_section_heading(doc, "1.1", "ความเป็นมาและความสำคัญของปัญหา")
    add_thai_body_p(
        doc,
        "ข้อมูลความเข้มการแผ่รังสีแสงอาทิตย์รวมในแนวราบ (Global Horizontal Irradiance: GHI, หน่วย W/m²) "
        "เป็นตัวแปรทางฟิสิกส์ที่มีความสำคัญอย่างยิ่งต่อการประเมินประสิทธิภาพและตรวจวัดสุขภาพของโรงไฟฟ้าพลังงานแสงอาทิตย์ "
        "(Photovoltaic System Performance Ratio: PR) ตลอดจนการศึกษาอุตุนิยมวิทยาพลังงานแสงอาทิตย์ เครื่องมือวัดรังสีดวงอาทิตย์ "
        "(Pyranometer) เกรดมาตรฐานอ้างอิงสากลตาม ISO 9060:2018 เช่น ระดับ Class A มักใช้เทคโนโลยีหัววัดแบบเทอร์โมไพล์ (Thermopile) "
        "และครอบด้วยโดมแก้วออปติคอลคู่เจียระไนความแม่นยำสูง ซึ่งมีราคาจำหน่ายสูงมาก (มากกว่า 80,000 ถึง 150,000 บาทต่อเครื่อง) "
        "ทำให้เป็นอุปสรรคสำคัญในการติดตั้งสถานีตรวจวัดแบบกระจายตัวหนาแน่น (Distributed Sensor Network)"
    )
    add_thai_body_p(
        doc,
        "การพัฒนาไพราโนมิเตอร์ราคาประหยัดโดยใช้เซนเซอร์สารกึ่งตัวนำชนิดซิลิคอนโฟโตไดโอด (Silicon Photodiode) "
        "เป็นแนวทางที่ได้รับความสนใจอย่างกว้างขวางเนื่องจากมีเวลาตอบสนองรวดเร็วและให้สัญญาณไฟฟ้าที่ชัดเจน "
        "อย่างไรก็ตาม การนำโฟโตไดโอดมาสร้างเป็นไพราโนมิเตอร์ต้องเผชิญกับความท้าทายหลัก 2 ประการ ได้แก่: "
        "(1) ความคลาดเคลื่อนเชิงมุมตกกระทบ (Cosine Error) ซึ่งเกิดจากลักษณะทางกายภาพของโครงสร้างตัวถังและการหักเหของกระจกหน้าต่าง "
        "และ (2) ความเสถียรและความเงียบของวงจรขยายสัญญาณ ซึ่งต้องการวงจรแปลงสัญญาณดิจิทัลที่มีความละเอียดสูงและแหล่งจ่ายแรงดันอ้างอิงที่นิ่งสนิท"
    )
    add_thai_body_p(
        doc,
        "โครงงานนี้จึงมุ่งวิจัยและพัฒนาต้นแบบระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง โดยทำการวิเคราะห์เปรียบเทียบชิปเซนเซอร์ "
        "จำลองพฤติกรรมเชิงแสงเพื่อศึกษาผลกระทบของการหักเหแสงผ่านกระจกหน้าต่าง สร้างและทดสอบวงจรอ่านสัญญาณความละเอียด 24 บิตบน Breadboard "
        "และออกแบบระเบียบวิธีการทดลองทางทัศนศาสตร์ร่วมกับภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ เพื่อเตรียมการสอบเทียบตามมาตรฐานสากลต่อไป"
    )

    add_section_heading(doc, "1.2", "วัตถุประสงค์ของโครงงาน")
    add_numbered_item(doc, "1.", "เพื่อศึกษาและวิเคราะห์เปรียบเทียบคุณลักษณะทางแสงและกายภาพระหว่างโฟโตไดโอดรุ่น BPX 61 และ BPW 34 ในการประยุกต์ใช้งานเป็นหัววัดไพราโนมิเตอร์", bold_prefix="การคัดเลือกหัววัด: ")
    add_numbered_item(doc, "2.", "เพื่อออกแบบ สร้าง และทดสอบวงจรอ่านสัญญาณต้นแบบบน Breadboard โดยใช้ตัวแปลงสัญญาณ 24-bit ADC (ADS1220) และวิเคราะห์เสถียรภาพของ REF5045 ต่อค่า ESR ของตัวเก็บประจุ", bold_prefix="การสร้างและดีบักวงจร: ")
    add_numbered_item(doc, "3.", "เพื่อจำลองพฤติกรรมทางทัศนศาสตร์เชิงมุม (Optical Ray-tracing Simulation) และวิเคราะห์ผลกระทบของโครงสร้างตัวถังและกระจกหน้าต่างต่อความคลาดเคลื่อนโคไซน์", bold_prefix="การจำลองทางทัศนศาสตร์: ")
    add_numbered_item(doc, "4.", "เพื่อออกแบบระเบียบวิธีการทดลองจริง (Experimental Protocols) สำหรับการวัดค่าการส่องผ่านสเปกตรัม (Spectrophotometer) และการวัดการตอบสนองเชิงมุมร่วมกับคณะวิทยาศาสตร์", bold_prefix="การออกแบบการทดลอง: ")

    add_section_heading(doc, "1.3", "ขอบเขตของโครงงาน (Capstone Project 1)")
    add_bullet_item(doc, bullet_char="•", text="การศึกษาและวิเคราะห์สเปกตรัมของโฟโตไดโอดซิลิคอนในช่วงความยาวคลื่น 400 ถึง 1,100 nm")
    add_bullet_item(doc, bullet_char="•", text="การออกแบบและประกอบวงจรทดสอบต้นแบบบน Breadboard พร้อมหัววัด NTC ชดเชยอุณหภูมิ และเก็บบันทึกข้อมูลผ่าน Serial Monitor ใน 4 สภาวะแสง")
    add_bullet_item(doc, bullet_char="•", text="การดีบักและวิเคราะห์เสถียรภาพแรงดันอ้างอิง 4.500 V (REF5045) ต่อค่า ESR ของตัวเก็บประจุโหลด เพื่อเตรียมการลงบอร์ด Perfboard")
    add_bullet_item(doc, bullet_char="•", text="การจำลองพฤติกรรมเชิงมุมตกกระทบ (0 ถึง 85 องศา) บนแบบจำลองเรขาคณิต 3 มิติของตัวถัง TO-39 (มีและไม่มีกระจก) เทียบกับ BPW 34")
    add_bullet_item(doc, bullet_char="•", text="การจัดทำคู่มือและระเบียบวิธีทดสอบ (Testing Protocols) 2 ชุด เพื่อใช้กับเครื่องมือวัดของภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล.")

    add_section_heading(doc, "1.4", "ประโยชน์ที่คาดว่าจะได้รับ")
    add_bullet_item(doc, bullet_char="•", text="ได้องค์ความรู้เชิงลึกเกี่ยวกับผลกระทบของการหักเหแสงผ่านกระจกหน้าต่างต่อคุณลักษณะการตอบสนองเชิงมุมของไพราโนมิเตอร์")
    add_bullet_item(doc, bullet_char="•", text="ได้ต้นแบบวงจรอ่านสัญญาณแอนะล็อกความละเอียดสูง 24 บิตบน Perfboard ที่ผ่านการทดสอบเก็บข้อมูลจริงและเข้าใจแนวทางการแก้ปัญหาภาคจ่ายไฟ")
    add_bullet_item(doc, bullet_char="•", text="ได้กรอบระเบียบวิธีการทดลองทางทัศนศาสตร์ที่ชัดเจน พร้อมสำหรับการทดสอบและสอบเทียบจริงในโครงงาน 2")
    add_bullet_item(doc, bullet_char="•", text="สามารถต่อยอดเป็นชุดเครื่องมือตรวจรับงานและตรวจสอบคุณภาพระบบโซลาร์เซลล์แบบพกพา (On-site Solar PV Audit / Commissioning) สำหรับกลุ่มบ้านพักอาศัยหรือระบบขนาดเล็ก (งบติดตั้ง 50,000–150,000 บาท) ที่ไม่คุ้มค่าต่อการติดตั้งไพราโนมิเตอร์ถาวร แต่ต้องการตรวจวัดค่า Performance Ratio (PR) ว่าได้มาตรฐานตามที่ผู้ติดตั้งคาดการณ์ไว้หรือไม่")

    add_section_heading(doc, "1.5", "แผนการดำเนินงานโครงงาน")
    add_thai_body_p(
        doc,
        "แผนการดำเนินงานโครงงาน Capstone Project 1 มีรายละเอียดขั้นตอนและผลสัมฤทธิ์ที่ส่งมอบดังระบุในตารางที่ 1.1"
    )

    # ตารางที่ 1.1
    add_table_caption(doc, "1.1", "แผนการดำเนินงานและกิจกรรมโครงงานตลอดปีการศึกษา")
    t_plan = doc.add_table(rows=6, cols=3)
    t_plan.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_plan, border_color="000000", border_sz="4")
    plan_headers = ["ระยะเวลา", "กิจกรรมและขั้นตอนการดำเนินงาน", "ผลสัมฤทธิ์ (Deliverables)"]
    for j, h in enumerate(plan_headers):
        cell = t_plan.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    plan_data = [
        ["เดือน 1–2", "ศึกษามาตรฐาน ISO 9060 และสเปกตรัมโฟโตไดโอด BPX 61 vs BPW 34", "ข้อสรุปการคัดเลือกเซนเซอร์และกระจกแซฟไฟร์"],
        ["เดือน 3–4", "ต่อวงจรทดสอบบน Breadboard และตรวจพบปัญหา Leakage ขา Trim", "วงจรแก้ปัญหา High-Z Node และ DNP C11"],
        ["เดือน 5", "จำลอง Ray-tracing เงากระป๋อง TO-39 และพบการเบี่ยงเบนเมื่อใส่กระจก Schott", "ผลซิมเปรียบเทียบ BPX 61 vs BPW 34"],
        ["เดือน 6", "วิเคราะห์ปัญหาเสถียรภาพ REF5045 ต่อ ESR ของ C และสั่งซื้อ C Low-ESR", "แนวทางแก้ไขวงจรและรอ C เพื่อลง Perfboard วันที่ 30 ต.ค."],
        ["เดือน 7 (ปัจจุบัน)", "ประสานงานคณะวิทยาศาสตร์ และออกแบบระเบียบวิธีทดสอบ 2 ชุด", "ข้อเสนอ Protocol การทดลอง และเล่มรายงาน Project 1"]
    ]
    for i, row in enumerate(plan_data):
        for j, val in enumerate(row):
            cell = t_plan.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=14)

    # -------------------------------------------------------------------------
    # บทที่ 2: ทฤษฎีและงานวิจัยที่เกี่ยวข้อง (Literature Review)
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "2", "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง\n(Literature Review)")

    add_section_heading(doc, "2.1", "มาตรฐานการวัดรังสีดวงอาทิตย์ (ISO 9060:2018)")
    add_thai_body_p(
        doc,
        "มาตรฐานสากล ISO 9060:2018 กำหนดเกณฑ์และคุณลักษณะเฉพาะของเครื่องวัดรังสีแสงอาทิตย์ (Pyranometer) "
        "โดยแบ่งชั้นคุณภาพออกเป็น Class A, Class B และ Class C ตามพารามิเตอร์สำคัญในตารางที่ 2.1 "
        "ซึ่งพารามิเตอร์ที่เป็นอุปสรรคสำคัญที่สุดในการออกแบบหัววัดแบบโซลิดสเตตคือ Directional Response (Cosine Error) "
        "ซึ่งมาตรฐานกำหนดให้มีความคลาดเคลื่อนไม่เกิน ±20 W/m² สำหรับ Class B และไม่เกิน ±10 W/m² สำหรับ Class A"
    )

    # ตารางที่ 2.1
    add_table_caption(doc, "2.1", "เกณฑ์มาตรฐานการจำแนกประเภท Pyranometer ตาม ISO 9060:2018")
    t_iso = doc.add_table(rows=6, cols=4)
    t_iso.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_iso, border_color="000000", border_sz="4")
    iso_headers = ["พารามิเตอร์สมรรถนะ", "Class A", "Class B", "Class C"]
    for j, h in enumerate(iso_headers):
        cell = t_iso.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    iso_data = [
        ["เวลาการตอบสนอง (Response Time, 95%)", "< 10 s (Fast < 0.5 s)", "< 20 s", "< 30 s"],
        ["การเลื่อนวนจากอุณหภูมิ (-10 ถึง 40 °C)", "±1%", "±2%", "±4%"],
        ["ความไม่เป็นเชิงเส้น (Non-linearity)", "±0.5%", "±1%", "±3%"],
        ["การตอบสนองต่อทิศทาง (Directional / Cosine Error)", "±10 W/m²", "±20 W/m²", "±30 W/m²"],
        ["ความไม่เสถียรต่อปี (Non-stability per year)", "±0.8%", "±1.5%", "±3%"]
    ]
    for i, row in enumerate(iso_data):
        for j, val in enumerate(row):
            cell = t_iso.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=14)

    add_thai_body_p(
        doc,
        "นอกจากเกณฑ์ในตารางที่ 2.1 (เอกสารต้นฉบับ ISO 9060:2018 Table 1 แสดงในภาคผนวก ข) มาตรฐานสากลยังกำหนดนิยาม "
        "และสมการคำนวณความคลาดเคลื่อนเชิงทิศทาง (Directional Error: Δ_1000(θ, ψ) และ Cosine Error: δ_cos(θ, ψ)) ไว้ในภาคผนวก Annex A.5 "
        "โดยกำหนดให้ทดสอบที่มุมตกกระทบสำคัญ 40°, 60°, 70°, 80° และ 95° เทียบกับมุมตั้งฉาก 0° ดังแสดงในรูปที่ 2.1"
    )
    iso_a5_path = os.path.join(img_dir, "iso9060_crops", "iso_p21_annex_a5.png")
    if os.path.exists(iso_a5_path):
        add_figure_image(
            doc,
            iso_a5_path,
            fig_no="2.1",
            title="นิยามและสมการการตอบสนองเชิงมุม (Directional Response / Cosine Error) จากเอกสาร ISO 9060:2018 (Annex A.5)",
            width_inches=5.4
        )

    add_section_heading(doc, "2.2", "การวิเคราะห์เปรียบเทียบชิปเซนเซอร์ BPX 61 และ BPW 34")
    add_thai_body_p(
        doc,
        "ในการเลือกเซนเซอร์ตรวจจับรังสีแสงอาทิตย์ ได้ทำการศึกษารายละเอียดจากเอกสารข้อมูลทางเทคนิค (Datasheet) ของผู้ผลิต ams-OSRAM "
        "ระหว่างรุ่น BPX 61 (Metal Can TO-39) และรุ่น BPW 34 (Flat Clear Plastic) ดังเปรียบเทียบในตารางที่ 2.2"
    )

    # ตารางที่ 2.2
    add_table_caption(doc, "2.2", "การเปรียบเทียบพารามิเตอร์ทางเทคนิคระหว่าง OSRAM BPX 61 กับ BPW 34")
    t_cmp = doc.add_table(rows=8, cols=3)
    t_cmp.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_cmp, border_color="000000", border_sz="4")
    cmp_headers = ["พารามิเตอร์ทางเทคนิค", "OSRAM BPX 61", "OSRAM BPW 34"]
    for j, h in enumerate(cmp_headers):
        cell = t_cmp.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    cmp_data = [
        ["ขนาดพื้นที่รับแสง (Radiant Sensitive Area)", "7.02 mm²", "7.02 mm² (ชิปตัวเดียวกัน)"],
        ["ขนาดของชิปซิลิคอน (Dimensions of Active Chip)", "2.65 x 2.65 mm", "2.65 x 2.65 mm (ชิปตัวเดียวกัน)"],
        ["ช่วงความยาวคลื่นตอบสนอง (Spectral Range)", "420 ... 1120 nm", "420 ... 1120 nm"],
        ["ความยาวคลื่นที่ไวสูงสุด (Peak Wavelength)", "920 nm", "920 nm"],
        ["ความไวของชิปที่ 850 nm (Spectral Sensitivity)", "0.62 A/W", "0.62 A/W"],
        ["ความจุไฟฟ้าของรอยต่อ (Capacitance, V_R = 0V)", "72 pF", "72 pF"],
        ["สัมประสิทธิ์อุณหภูมิ (TC_I)", "+0.18 %/K", "+0.18 %/K"]
    ]
    for i, row in enumerate(cmp_data):
        for j, val in enumerate(row):
            cell = t_cmp.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=14)

    add_thai_body_p(
        doc,
        "จากข้อมูลเชิงเปรียบเทียบในตารางที่ 2.2 จะเห็นได้อย่างชัดเจนว่าทั้ง BPX 61 และ BPW 34 ใช้ชิปซิลิคอน (Silicon Die) "
        "เทคโนโลยีเดียวกันทุกประการ มีพื้นที่รับแสง 7.02 mm² เท่ากัน และมีพฤติกรรมทางสเปกตรัมเหมือนกัน "
        "แต่จุดที่แตกต่างกันอย่างสิ้นเชิงคือโครงสร้างตัวถังภายนอก (Packaging):"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="BPX 61 บรรจุในตัวถังโลหะ TO-39 โดยชิปซิลิคอนถูกติดตั้งจมลึกลงไปในหลุมกระป๋องถึง 1.55 ถึง 1.75 mm ใต้กระจกหน้าต่างขนาดเส้นผ่านศูนย์กลาง 6.0 mm "
        "โครงสร้างหลุมนี้ทำให้เกิดการบดบังแสงเชิงมุมอย่างรุนแรง ดังที่ปรากฏในกราฟ Directional Characteristics ใน Datasheet (รูปที่ 2.1) "
        "ซึ่งสัญญาณแสงจะตกฮวบอย่างรวดเร็วที่มุมเกิน 60 องศา และถูกตัดแสงสิ้นเชิงที่ 70 องศา (รูปที่ 2.2)"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="BPW 34 บรรจุในตัวถังพลาสติกใสแบบแบนราบ ผิวชิปอยู่ชิดระนาบผิวด้านหน้าโดยไม่มีขอบกระป๋องยกสูง ทำให้มีมุมรับแสงกว้างมาก (Half Angle ±60 องศา) "
        "และตอบสนองต่อแสงอาทิตย์ในมุมเอียงได้ใกล้เคียงกับกฎโคไซน์ตามธรรมชาติมากกว่า"
    )

    # รูปที่ 2.2 Polar Plot
    p_bpx61 = os.path.join(img_dir, "(Directional Polar Plot) ของ BPX 61 .png")
    p_bpw34 = os.path.join(img_dir, "(Directional Polar Plot) ของ BPW 34.png")
    if os.path.exists(p_bpx61) and os.path.exists(p_bpw34):
        add_two_figure_images_side_by_side(
            doc,
            p_bpx61, "(ก) BPX 61 (TO-39 ขอบกระป๋องบังแสงที่มุมสูง)",
            p_bpw34, "(ข) BPW 34 (ตัวถังแบนราบ Half-angle ±60°)",
            fig_no="2.2",
            main_title="คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) จาก Datasheet",
            col_width_inch=2.7
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.2",
            title="คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) ของ BPX 61 เทียบกับ BPW 34",
            description="ภาพกราฟ Polar Plot จาก Datasheet ของ OSRAM แสดงมุมรับแสงที่แตกต่างกันอย่างสิ้นเชิง",
            instructions="นำภาพ Polar Plot จาก Datasheet มาจัดวาง"
        )

    # รูปที่ 2.3 Dimension
    dim_bpx61 = os.path.join(img_dir, "dimention(BPX61).png")
    dim_bpw34 = os.path.join(img_dir, "dimention (BPW34).png")
    if os.path.exists(dim_bpx61) and os.path.exists(dim_bpw34):
        add_two_figure_images_side_by_side(
            doc,
            dim_bpx61, "(ก) ตัวถัง TO-39 (BPX 61) ขอบลึก 1.55–1.75 มม.",
            dim_bpw34, "(ข) ตัวถังแบนราบ (BPW 34) ชิปอยู่ชิดผิวหน้า",
            fig_no="2.3",
            main_title="มิติและโครงสร้างกายภาพของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34",
            col_width_inch=2.7
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.3",
            title="มิติและโครงสร้างกายภาพตัดขวางของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34",
            description="ภาพแสดงโครงสร้างหลุมลึก 1.55–1.75 มม. ของกระป๋อง TO-39 เทียบกับตัวถังแบนราบของ BPW 34",
            instructions="จัดวางภาพเปรียบเทียบมิติตัวถังจาก Datasheet หน้า 9"
        )

    # -------------------------------------------------------------------------
    # 2.3 ทฤษฎีโฟโตไดโอด วงจรขยาย TIA และการคำนวณชดเชยอย่างเป็นระบบ
    # -------------------------------------------------------------------------
    add_section_heading(doc, "2.3", "ทฤษฎีโฟโตไดโอด วงจรขยายกระแสแสง (TIA) และการชดเชยทางคณิตศาสตร์")
    add_thai_body_p(
        doc,
        "การตรวจวัดรังสีดวงอาทิตย์ด้วยเซนเซอร์สารกึ่งตัวนำชนิดซิลิคอนโฟโตไดโอด (Silicon PIN Photodiode) "
        "ภายใต้หน้าต่างกระจกแซฟไฟร์ใส (Clear Synthetic Sapphire Window, ค่าการส่องผ่านแสง T_sapphire ≈ 85% โดยไม่ใช้แผ่นกรองแสงขุ่น PTFE Diffuser) "
        "ต้องอาศัยการวิเคราะห์เชิงลึกทั้งด้านฟิสิกส์ของรอยต่อ p-i-n ภายใต้กระแสแสงความเข้มสูง, การจัดระดับแรงดันไบแอสในระบบไฟเลี้ยงเดี่ยว (Single-Supply Biasing), "
        "การกำหนดอัตราขยายของวงจร Transimpedance Amplifier (TIA) เพื่อรองรับพิกัดสูงสุดระดับเครื่องมือวัดชั้นหนึ่ง (ISO 9060 Class A: 4,000 W/m²) "
        "ตลอดจนการสร้างระบบสมการชดเชยความคลาดเคลื่อนทางคณิตศาสตร์ที่อ้างอิงจากเอกสารข้อมูลทางเทคนิค (Datasheet) ของผู้ผลิตโดยตรง"
    )

    add_subsection_heading(doc, "2.3.1", "ฟิสิกส์ของโฟโตไดโอด BPW 34 และการพิสูจน์ความเป็นเชิงเส้นที่ความเข้มแสงสูง")
    add_thai_body_p(
        doc,
        "เมื่อโฟตอนของแสงอาทิตย์ที่มีพลังงานสูงกว่าแถบพลังงานต้องห้ามของซิลิคอน (Bandgap Energy, E_g ≈ 1.12 eV) "
        "ตกกระทบลงบนชั้นรอยต่อของโฟโตไดโอด BPW 34 หรือ BPX 61 (พื้นที่รับแสงกัมมันต์ A_active = 7.02 mm²) "
        "จะเกิดกระแสโฟโตเคอร์เรนต์ (Photocurrent, I_P) ซึ่งแปรผันตรงกับความเข้มรังสีแสงอาทิตย์ตกกระทบ (Irradiance, E หรือ G ในหน่วย W/m²) ดังสมการ:"
    )
    add_equation_box(doc, "I_P = S_eff · G = (R_λ · A_active · T_sapphire) · G", eq_no="2.1")
    add_thai_body_p(
        doc,
        "ภายใต้สเปกตรัมแสงอาทิตย์มาตรฐาน AM1.5G โฟโตไดโอดตระกูล BPW 34 มีค่าความไวกระแสแสงเปลือยประมาณ 3.06 µA ต่อ W/m² "
        "(หรือประมาณ 3.06 mA ที่ความเข้มแสงอาทิตย์มาตรฐาน 1,000 W/m²) เมื่อติดตั้งภายใต้หน้าต่างกระจกแซฟไฟร์ใสที่มีค่าการส่องผ่านสุทธิ "
        "T_sapphire ≈ 85% (หักการสูญเสียจากการสะท้อนแบบเฟรสเนลทั้งสองผิวหน้าประมาณ 14–15%) จะได้ค่าความไวสุทธิของหัววัด (Effective Irradiance Sensitivity, S_eff) เท่ากับ:"
    )
    add_equation_box(doc, "S_eff ≈ 3.06 µA/(W/m²) × 0.85 ≈ 2.60 µA ต่อ 1 W/m²", eq_no="2.2")
    add_thai_body_p(
        doc,
        "ประเด็นทางวิศวกรรมที่สำคัญอย่างยิ่งเมื่อนำเซนเซอร์ BPW 34 ไปรับแสงอาทิตย์โดยตรงผ่านกระจกแซฟไฟร์ใสโดยไม่มีแผ่นลดทอนแสง (No Diffuser) "
        "คือคำถามที่ว่า 'ตัวชิปซิลิคอนจะเกิดการอิ่มตัวทางแสง (Optical Saturation) หรือได้รับความเสียหายจากความร้อนหรือไม่?' "
        "คณะผู้จัดทำได้ทำการพิสูจน์ทางคณิตศาสตร์จากเอกสารข้อมูลทางเทคนิคของ ams-OSRAM BPW 34 (Version 1.6, หน้า 2–3) พบข้อเท็จจริง 3 ประการดังนี้:"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="ขีดจำกัดการทนกำลังงานความร้อน (Maximum Power Dissipation): จากสเปกชีต BPW 34 หน้า 2 ระบุพิกัดกำลังงานสูงสุด P_tot(max) = 150 mW ที่อุณหภูมิ 25 °C "
        "ในขณะที่พลังงานแสงอาทิตย์ที่ความเข้มมาตรฐาน 1,000 W/m² ตกกระทบลงบนพื้นที่ชิป 7.02 mm² คิดเป็นกำลังงานทางแสงเพียง P_opt = 0.1 W/cm² × 0.0702 cm² = 7.02 mW "
        "และแม้ในสภาวะความเข้มรังสีสูงสุดระดับพิกัดเครื่องมือวัด Class A ที่ 4,000 W/m² กำลังงานแสงที่ตกกระทบมีค่าเพียง 28.08 mW ซึ่งยังต่ำกว่าพิกัดสูงสุด 150 mW ถึงกว่า 5.3 เท่า"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="การวิเคราะห์ความต้านทานอนุกรมภายใน (R_S) และเงื่อนไขความเป็นเชิงเส้นที่กระแสแสงสูง: ในสเปกชีต BPW 34 หน้า 3 ระบุว่าที่กระแสไบแอสตรงสูง I_F = 100 mA "
        "แรงดันตกคร่อมไดโอดมีค่า V_F = 1.3 V ซึ่งคิดเป็นความต้านทานอนุกรมในสภาวะฉีดพาหะระดับสูง (Conductivity Modulation) เพียง ~4.5 Ω อย่างไรก็ตาม เมื่อโฟโตไดโอดทำงานในโหมดรับแสงที่แรงดันไบแอสย้อนกลับต่ำ "
        "(Low Reverse Bias: 0.1–0.6 V) ชั้นปลอดพาหะ (Depletion Layer) จะกางออกเพียง ~25% ของความหนาชั้น Intrinsic (สังเกตจากค่า C_J ≈ 55–72 pF เทียบกับ 15 pF เมื่อ Deplete เต็มที่ในหน้า 6) "
        "ประกอบกับกระแสโฮลที่เกิดบริเวณกลางแผ่นชิปขนาด 2.65 × 2.65 mm ต้องไหลออกทางด้านข้างผ่านชั้นผิวหน้า P+ ที่บางมากไปยังขอบโลหะ ทำให้เกิดความต้านทานแผ่นตามแนวราบ (Lateral Sheet Resistance) "
        "รวมกับความต้านทานเนื้อซิลิคอนส่วนที่ยังไม่ Deplete คิดเป็นความต้านทานอนุกรมสมมูลขณะรับแสง R_S(photo) ≈ 30–50 Ω ดังนั้น เงื่อนไขสำคัญที่จะทำให้เซนเซอร์มีความเป็นเชิงเส้น 100% โดยไม่เกิดการอิ่มตัวภายใน "
        "คือต้องรักษาแรงดันไบแอสย้อนกลับ V_bias ให้สูงพอที่จะหักล้างแรงดันตกคร่อมภายใน (I_P · R_S(photo)) ไม่ให้รอยต่อ p-n พลิกกลับเป็นไบแอสตรงเข้าใกล้ V_OC (0.365 V)"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="ความสามารถในการรับกระแสของออปแอมป์ (Op-Amp Short-Circuit Current Capability): จากสเปกชีตของไอซีออปแอมป์ความแม่นยำสูง Texas Instruments OPA380 (SBOS291G, หน้า 3) "
        "ระบุพิกัดกระแสเอาต์พุตสูงสุด (Short-Circuit Current, I_SC) สูงถึง 150 mA ซึ่งสามารถรองรับกระแสโฟโตเคอร์เรนต์ 2.6 mA ถึง 10.4 mA ได้อย่างสมบูรณ์โดยไม่เกิดสภาวะ Current Limiting"
    )

    add_subsection_heading(doc, "2.3.2", "การวิเคราะห์แรงดันไบแอส (V_bias) ในระบบไฟเลี้ยงเดี่ยว: 0 V vs 10 V vs Controlled Low-Bias")
    add_thai_body_p(
        doc,
        "ในตำราวงจรโฟโตไดโอดทั่วไป มักแบ่งโหมดการทำงานออกเป็นโหมด Photovoltaic (Zero Bias: V_R = 0 V) สำหรับงานความแม่นยำสูง และโหมด Photoconductive "
        "(High Reverse Bias: เช่น V_R = 5 V ถึง 10 V) สำหรับงานสื่อสารแสงความเร็วสูง อย่างไรก็ตาม เมื่อนำมาออกแบบเป็นวงจรแอนะล็อกฟรอนต์เอนด์ที่ใช้แหล่งจ่ายไฟเดี่ยว "
        "(Single-Supply Operation: V- = GND = 0 V, V+ = +5.2 V) ร่วมกับไอซี OPA380 คณะผู้จัดทำได้วิเคราะห์เปรียบเทียบทางเลือกการให้แรงดันไบแอสทั้ง 3 รูปแบบ "
        "จากข้อมูลสเปกชีตจริง (รูปที่ 2.4, รูปที่ 2.5 และตารางที่ 2.3) ดังนี้:"
    )

    # รูปที่ 2.4 IV Curve
    iv_curve_path = os.path.join(img_dir, "photodiode_iv_curve.png")
    if os.path.exists(iv_curve_path):
        add_figure_image(
            doc,
            iv_curve_path,
            fig_no="2.4",
            title="คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) ของโฟโตไดโอดและการจัดระดับจุดทำงาน",
            width_inches=5.6
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.4",
            title="คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) ของโฟโตไดโอดและการจัดระดับจุดทำงาน",
            description="กราฟแสดงเส้นคุณลักษณะ I-V ภายใต้ความเข้มแสงต่างๆ และการแบ่งย่านการทำงาน",
            instructions="วางภาพกราฟ I-V Curves"
        )

    # รูปที่ 2.5 ตารางพารามิเตอร์อุณหภูมิและความจุจาก Datasheet BPW 34
    bpw_tc_path = os.path.join(img_dir, "v2_extracted", "p16_img0_703x261.png")
    if os.path.exists(bpw_tc_path):
        add_figure_image(
            doc,
            bpw_tc_path,
            fig_no="2.5",
            title="ค่าพารามิเตอร์สัมประสิทธิ์อุณหภูมิ (TC_I = 0.18 %/K) และความจุไฟฟ้า (C_0 = 72 pF) จาก Datasheet BPW 34",
            width_inches=5.4
        )

    add_numbered_item(
        doc,
        "1.",
        "จากสเปกชีต BPW 34 หน้า 3 และหน้า 5 พิสูจน์อย่างชัดเจนว่าที่ความสว่าง 1,000 lx กระแสแสง I_P มีค่าเท่ากับ 80 µA เท่ากันทุกประการทั้งที่ V_R = 0 V และ V_R = 5 V "
        "(การเพิ่มแรงดันไบแอสเป็น 10 V ไม่ได้ช่วยให้รับแสงได้มากขึ้น) สาเหตุที่วงจรไฟเบอร์ออปติกใช้ไฟไบแอส 10 V มีเพียงเพื่อลดความจุไฟฟ้าของรอยต่อ (C_J ลดจาก 72 pF เหลือ 18 pF) "
        "สำหรับการรับส่งข้อมูลระดับ MHz แต่ในการวัดรังสีดวงอาทิตย์ (20 SPS หรือ DC) การจ่ายไฟไบแอส 10 V กลับส่งผลเสียอย่างร้ายแรง เพราะทำให้กระแสรั่วในความมืด (Dark Current, I_R) "
        "พุ่งสูงขึ้นถึง 2,000 pA (2 nA) ที่ 25 °C (สูงกว่าการใช้งานที่แรงดันต่ำถึง 24 เท่า) และพุ่งทะลุ 100,000 pA ที่อุณหภูมิ 65 °C",
        bold_prefix="เหตุผลที่ไม่ใช้แรงดันไบแอสสูง 5 V – 10 V: "
    )
    add_numbered_item(
        doc,
        "2.",
        "จากสเปกชีต OPA380 หน้า 3 และหน้า 10 (หัวข้อ Achieving Output Swing to the Op Amp Negative Rail) ระบุว่าในระบบไฟเลี้ยงเดี่ยว (V- = GND) "
        "ขาเอาต์พุตของออปแอมป์ไม่สามารถสวิงลงแตะระดับ 0.000 V ได้จริง แต่จะติดขีดจำกัดอิ่มตัวด้านล่าง (Negative Rail Saturation Floor) ที่ประมาณ +60 mV (ค่าปกติ) ถึง +100 mV (ค่าสูงสุด) "
        "หากถอดวงจร V_bias ออกโดยต่อขา +IN ลงกราวด์ (0 V) เอาต์พุตในความมืดจะค้างอยู่ที่ ~60–100 mV ทำให้เกิดย่านบอดของสัญญาณ (Dead-Zone) ในช่วงแสงน้อยและไม่สามารถวัดค่า Zero-Offset ที่แท้จริงได้",
        bold_prefix="เหตุผลที่ไม่สามารถตัด V_bias ออกเป็น 0 V ในระบบไฟเลี้ยงเดี่ยว: "
    )
    add_numbered_item(
        doc,
        "3.",
        "วงจรที่คณะผู้จัดทำออกแบบใช้ตัวแบ่งแรงดันความแม่นยำสูงจาก V_REF = 4.500 V (ผ่านตัวต้านทาน R3 = 100 kΩ และ R4 พร้อมตัวเก็บประจุกรองความถี่ C3 = 1 µF) "
        "ป้อนเข้าที่ขา +IN ของ OPA380 ในขณะที่ขา Anode ของโฟโตไดโอดต่อลงกราวด์โดยตรง (0 V) และขา Cathode ต่อเข้าที่ขา -IN ซึ่งตรงตามวงจรมาตรฐานในสเปกชีต OPA380 หน้า 1 "
        "(Figure 1: Single-Supply Connection) และหน้า 9 (Figure 4: Filtered Reverse Bias Voltage) ทุกประการ ทั้งนี้ ห้ามย้ายขา Anode ของไดโอดไปต่อร่วมกับจุด V_bias โดยเด็ดขาด "
        "เนื่องจากวงจรแบ่งแรงดัน R3 || R4 มีความต้านทานสมมูลเทเวนิน R_Th ≈ 9.09 kΩ หากกระแสแสงระดับ mA ไหลย้อนเข้าสู่จุด V_bias จะทำให้แรงดันฐานพังทลายทันที",
        bold_prefix="สถาปัตยกรรม Controlled Low-Bias ตามมาตรฐาน TI OPA380: "
    )
    add_numbered_item(
        doc,
        "4.",
        "เมื่อใช้งานภายใต้กระจกแซฟไฟร์ใส (ไม่มี PTFE Diffuser) ที่ความเข้มแสงอาทิตย์ 1,500 W/m² กระแสแสงจะสูงถึง I_P ≈ 3.8–3.9 mA ซึ่งทำให้เกิดแรงดันตกคร่อมความต้านทานภายในชิป "
        "V_drop,internal = I_P × R_S(photo) ≈ 3.8 mA × 50 Ω = 0.190 V หากปรับลด V_bias ลงต่ำเกินไป (เช่น 0.14 V) แรงดันตกคร่อมภายใน 0.190 V จะดันให้รอยต่อภายในชิปพลิกเข้าสู่สภาวะไบแอสตรง "
        "และทำให้กราฟสูญเสียความเป็นเชิงเส้นทันที ในทางกลับกัน การคงค่าเริ่มต้น R3 = 100 kΩ, R4 = 10 kΩ (V_bias = 0.409 V, วัดจริง 0.418 V) จะช่วยรักษาแรงดันไบแอสย้อนกลับสุทธิภายในรอยต่อ "
        "(V_R,net = 0.418 V - 0.190 V = +0.228 V > 0 V) ทำให้เซนเซอร์มีความเป็นเชิงเส้น 100% ตลอดช่วงแดดธรรมชาติไปจนถึง ~2,300 W/m² และหากต้องการขยายพิกัดสูงสุดไปถึงระดับเรือธง 4,000 W/m² "
        "(I_P ≈ 10.4 mA, V_drop,internal ≈ 0.31–0.52 V) สามารถใช้ R4 = 10 kΩ ถึง 15 kΩ (V_bias = 0.418 V ถึง 0.587 V) คู่กับ R1 = 360–390 Ω เพื่อรักษารอยต่อให้เป็น Reverse Bias ตลอดย่าน 4,000 W/m² "
        "(ส่วนการลดค่า R4 เหลือ 3.3–4.7 kΩ เหมาะเฉพาะกรณีที่มีแผ่นลดทอนแสงหรือวัดแสงในร่มที่กระแส I_P < 1.5 mA เท่านั้น) ดังสรุปในตารางที่ 2.3",
        bold_prefix="ความสำคัญของ V_bias = 0.418 V (และ 0.58 V) ในการป้องกันการอิ่มตัวภายในเมื่อรับแดดแรงผ่านกระจกใส: "
    )

    # ตารางที่ 2.3 เปรียบเทียบการตั้งค่า Vbias
    add_table_caption(doc, "2.3", "การวิเคราะห์เปรียบเทียบทางเลือกการตั้งค่าแรงดันไบแอส (V_bias) ต่อความเป็นเชิงเส้นที่ความเข้มแสงสูง")
    t_mode = doc.add_table(rows=6, cols=6)
    t_mode.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_mode, border_color="000000", border_sz="4")
    mode_headers = [
        "รูปแบบการตั้งค่าไบแอส\n(R3 = 100 kΩ คงที่)",
        "แรงดัน V_bias\n(V)",
        "สถานะเอาต์พุต\nOPA380 ในความมืด",
        "พิกัดกระแสสูงสุดที่ยัง\nLinear 100% (R_S≈50Ω)",
        "กระแสรั่วมืด I_dark\nที่ 25°C / 65°C",
        "ความเหมาะสมทางวิศวกรรม"
    ]
    for j, h in enumerate(mode_headers):
        cell = t_mode.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=55, bottom=55, left=50, right=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=12, bold=True)

    mode_data = [
        [
            "1. ถอด V_bias ออก\n(+IN ต่อลง GND)",
            "0.000 V",
            "ติดพื้นอิ่มตัว\n(+60 ถึง +100 mV)",
            "ไม่ Linear เมื่อแดดแรง\n(รอยต่อไบแอสตรงทันที)",
            "~0 pA",
            "ห้ามใช้ในไฟเลี้ยงเดี่ยว (เสียทั้งย่านมืดและอิ่มตัวเร็วเมื่อแดดจัด)"
        ],
        [
            "2. ไบแอสต่ำ\n(R4 = 3.3k – 4.7k)",
            "0.144 – 0.202 V",
            "พ้นพื้นอิ่มตัว\n100 mV",
            "I_P ≤ 3.0–4.0 mA\n(~1,100–1,500 W/m²)",
            "~29–40 pA /\n~1.5–2.0 nA",
            "เหมาะเฉพาะกรณีใช้ในร่มหรือมีแผ่นกรองแสง (ถ้ารับแดดเกิน 1,500 W/m² จะเริ่มโค้ง)"
        ],
        [
            "3. ค่ามาตรฐานบนบอร์ด\n(ใช้ R4 = 10 kΩ)",
            "0.4091 V\n(วัดจริง 0.418 V)",
            "อยู่ในย่านเชิงเส้น\nสมบูรณ์แบบ",
            "I_P ≤ 6.5–8.0 mA\n(~2,500–3,000 W/m²)",
            "~84 pA /\n~4.2 nA",
            "แนะนำเป็นค่าหลัก! ป้องกันรอยต่อไบแอสตรงได้ตลอดช่วงแดดจัดผ่านกระจกใส"
        ],
        [
            "4. ไบแอสพิกัดสูงสุด\n(ใช้ R4 = 15 kΩ)",
            "0.5870 V\n(~587 mV)",
            "อยู่ในย่านเชิงเส้น\nสมบูรณ์แบบ",
            "I_P ≥ 10.5 mA\n(Linear เต็ม 4,000 W/m²)",
            "~105 pA /\n~5.2 nA",
            "แนะนำเมื่อต้องการทำรุ่นท็อป 4,000 W/m² (คู่กับ R1 = 360–390 Ω) Linear 100%"
        ],
        [
            "5. ไบแอสย้อนกลับสูง\n(External 10.0 V)",
            "10.000 V",
            "เกินพิกัดไฟเลี้ยง\nOPA380 (Max 5.5V)",
            "ใช้งานไม่ได้\nในระบบ 5.2 V",
            "2,000 pA /\n100,000 pA",
            "ไม่เหมาะสมอย่างยิ่ง (I_dark สูงขึ้น 24 เท่าโดยไม่ได้เพิ่มกระแสแสง)"
        ]
    ]
    for i, row in enumerate(mode_data):
        for j, val in enumerate(row):
            cell = t_mode.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=45, right=45)
            p = cell.paragraphs[0]
            if j in [0, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_subsection_heading(doc, "2.3.3", "สถาปัตยกรรมวงจรขยาย TIA และการคำนวณเลือกค่า R1 รองรับพิกัดสูงสุด 4,000 W/m²")
    add_thai_body_p(
        doc,
        "ผังวงจรภาคขยายกระแสแสง Transimpedance Amplifier (TIA) ที่ออกแบบและใช้งานจริงในโครงงานนี้ แสดงดังรูปที่ 2.6 "
        "ประกอบด้วยไอซี OPA380AIDR (IC1) รับไฟเลี้ยง +5.2 V จากเรกูเลเตอร์สัญญาณรบกวนต่ำ พร้อมตัวเก็บประจุบายพาส C4 = 1 µF ชิดขา V+ "
        "วงจรป้อนกลับประกอบด้วยตัวต้านทานกำหนดอัตราขยาย R1 (R_TIA) ขนานอยู่กับตัวเก็บประจุชดเชยเฟส C1 = 10 pF (ชนิด C0G/NP0) "
        "และผ่านวงจรกรองความถี่ต่ำ RC เอาต์พุต (R2 = 100 Ω, C2 = 0.1 µF, ความถี่ตัด f_c ≈ 15.9 kHz) เพื่อป้องกันสัญญาณรบกวนความถี่สูงและทำหน้าที่เป็นชุดขับประจุเข้าสู่ขา ADC"
    )

    # รูปที่ 2.6 TIA Schematic จริงจาก KiCad
    tia_kicad_path = os.path.join(img_dir, "tia_schematic_kicad.png")
    tia_arch_path = os.path.join(img_dir, "tia_circuit_architecture.png")
    if os.path.exists(tia_kicad_path):
        add_figure_image(
            doc,
            tia_kicad_path,
            fig_no="2.6",
            title="ผังวงจรขยายกระแสแสง Transimpedance Amplifier (TIA) ด้วยไอซี OPA380 และวงจรกรอง V_bias ที่ใช้งานจริง",
            width_inches=5.8
        )
    elif os.path.exists(tia_arch_path):
        add_figure_image(
            doc,
            tia_arch_path,
            fig_no="2.6",
            title="สถาปัตยกรรมวงจรขยายกระแสแสงแบบ Transimpedance Amplifier (TIA)",
            width_inches=5.8
        )

    add_thai_body_p(
        doc,
        "จากสเปกชีต OPA380 หน้า 11 (สมการที่ 2: Maximizing Transimpedance Bandwidth) ค่าตัวเก็บประจุป้อนกลับ C1 ที่เหมาะสมที่สุดเพื่อให้ได้การตอบสนองความถี่แบบ "
        "Maximally Flat (2nd-Order Butterworth) คำนวณได้จากสมการ:"
    )
    add_equation_box(doc, "C_1 = √[ C_TOT / (2π · R_1 · GBW) ]", eq_no="2.3")
    add_thai_body_p(
        doc,
        "เมื่อแทนค่าความจุรวมขาเข้า C_TOT = C_J(BPW34 @ 0.4V) + C_in(OPA380) ≈ 65 pF + 4 pF = 69 pF, ค่า GBW = 90 MHz และ R_1 = 2,200 Ω "
        "จะได้ค่าตามทฤษฎี C_1 = 9.82 pF ซึ่งตรงกับค่าอุปกรณ์มาตรฐาน C1 = 10 pF (C0G/NP0) ที่คณะผู้จัดทำเลือกใช้อย่างแม่นยำ "
        "โดยฟังก์ชันถ่ายโอนแรงดันเอาต์พุตของวงจร TIA กำหนดโดยสมการที่ 2.4:"
    )
    add_equation_box(doc, "V_op = V_bias + (I_P · R_1)", eq_no="2.4")
    add_thai_body_p(
        doc,
        "สำหรับการเลือกค่าตัวต้านทานป้อนกลับ R1 (R_TIA) เพื่อกำหนดพิกัดความเข้มรังสีสูงสุด (Maximum Irradiance Range, G_max) "
        "เนื่องจากไอซีแปลงสัญญาณ ADS1220 ใช้แรงดันอ้างอิงภายนอก V_REF = 4.500 V ดังนั้นแรงดันเอาต์พุตสูงสุดที่ยังไม่เกิดการอิ่มตัวของ ADC คือ V_op(max) = 4.500 V "
        "(คิดเป็นช่วงสวิงสัญญาณสุทธิ ΔV_max = 4.500 V - V_bias ≈ 4.082 V เมื่อ V_bias = 0.418 V หรือ ≈ 4.30 V เมื่อลด V_bias เป็น 0.20 V) "
        "คณะผู้จัดทำได้คำนวณออกแบบตารางเลือกค่า R1 เพื่อให้บอร์ดต้นแบบสามารถปรับระดับพิกัดการวัดได้ตั้งแต่การทดสอบในร่ม (2,200 Ω) "
        "ไปจนถึงพิกัดสูงสุดระดับเรือธงมาตรฐานสากล ISO 9060 Class A ที่ 4,000 W/m² (เทียบเท่าไพราโนมิเตอร์ชั้นนำของโลก เช่น Kipp & Zonen CMP22 / SMP11 และ Hukseflux SR30) "
        "เพื่อรองรับปรากฏการณ์ขอบเมฆรวมแสง (Cloud Enhancement / Edge-of-Cloud Effect) ในเขตร้อนได้โดยไม่มีวันอิ่มตัว ดังแสดงในตารางที่ 2.4"
    )

    # ตารางที่ 2.4 การเลือกค่า R1 รองรับพิกัดถึง 4,000 W/m²
    add_table_caption(doc, "2.4", "การคำนวณเลือกค่าตัวต้านทานป้อนกลับ R1 (R_TIA) เพื่อกำหนดพิกัดความเข้มแสงสูงสุดถึงระดับ Class A (4,000 W/m²)")
    t_rtia = doc.add_table(rows=6, cols=6)
    t_rtia.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_rtia, border_color="000000", border_sz="4")
    rtia_headers = [
        "ระดับพิกัด / การใช้งาน",
        "ค่าตัวต้านทาน R1\n(วิธีต่อค่ามาตรฐาน)",
        "กระแสสูงสุด I_P(max)\nก่อน ADC อิ่มตัว",
        "ความเข้มแสงสูงสุด G_max\n(กระจก Sapphire ใส)",
        "ความไวแรงดัน\n(mV ต่อ 1 W/m²)",
        "ความละเอียดใช้งานจริง\n(W/m² @ ~1.7 µV)"
    ]
    for j, h in enumerate(rtia_headers):
        cell = t_rtia.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=55, bottom=55, left=45, right=45)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=12, bold=True)

    rtia_data = [
        [
            "1. ค่าเริ่มต้นปัจจุบัน\n(ในร่ม / มีแผ่นกรองแสง)",
            "2,200 Ω\n(2.2 kΩ เดี่ยว)",
            "1.855 mA",
            "~714 W/m²\n(หรือ ~2,000 W/m² เมื่อมี Diffuser)",
            "5.72 mV / (W/m²)",
            "0.0003 W/m²"
        ],
        [
            "2. แดดจัดปกติ\n(Clear Sky Standard)",
            "1,100 Ω\n(2.2 kΩ ขนาน 2 ตัว)",
            "3.711 mA",
            "~1,430 W/m²\n(1,500 W/m² ที่ V_bias=0.2V)",
            "2.86 mV / (W/m²)",
            "0.0006 W/m²"
        ],
        [
            "3. แดดเขตร้อน + ขอบเมฆ\n(Tropical Cloud-Edge)",
            "680 Ω\n(หรือ 2.2k || 1k = 688 Ω)",
            "6.003 mA",
            "~2,310 W/m²",
            "1.77 mV / (W/m²)",
            "0.0010 W/m²"
        ],
        [
            "4. พิกัดสูงพิเศษ\n(Extended Solar Test)",
            "470 Ω\n(ค่ามาตรฐาน 0.1%)",
            "8.685 mA",
            "~3,340 W/m²",
            "1.22 mV / (W/m²)",
            "0.0014 W/m²"
        ],
        [
            "5. พิกัดสูงสุดเรือธง Class A\n(ISO 9060 Flagship Spec)",
            "390 Ω\n(หรือ 2.2k || 470Ω = 387 Ω)",
            "10.467 mA",
            "~4,025 W/m²\n(ครอบคลุมเต็ม 4,000 W/m²)",
            "1.01 mV / (W/m²)",
            "0.0017 W/m²"
        ]
    ]
    for i, row in enumerate(rtia_data):
        for j, val in enumerate(row):
            cell = t_rtia.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=45, right=45)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_thai_body_p(
        doc,
        "จากตารางที่ 2.4 จะเห็นได้ว่า ด้วยคุณสมบัติด้านความละเอียด 24 บิตของไอซี ADS1220 (1 LSB = 4.500 V / 2²³ = 0.5364 µV และมีระดับสัญญาณรบกวนอาร์เอ็มเอสที่อัตรา 20 SPS เพียง ~1.5 µV_rms) "
        "แม้จะปรับลดค่าตัวต้านทาน R1 ลงเหลือ 390 Ω เพื่อเปิดรับความเข้มรังสีสูงสุดถึงระดับเรือธง 4,000 W/m² วงจรยังคงให้ความไวสัญญาณที่ลงตัวพอดีคือ ≈ 1.01 mV ต่อ 1 W/m² "
        "และมีความละเอียดในการจำแนกความเข้มแสงสูงถึง 0.0017 W/m² (ละเอียดกว่าเกณฑ์ความต้องการของ ISO 9060 Class A ที่กำหนดความละเอียดไว้ 1 W/m² ถึงกว่า 500 เท่า)"
    )

    add_subsection_heading(doc, "2.3.4", "ระบบสมการคำนวณชดเชยทางคณิตศาสตร์ 4 ขั้นตอน (4-Stage Systematic Compensation Pipeline)")
    add_thai_body_p(
        doc,
        "เพื่อให้การประมวลผลความเข้มรังสีดวงอาทิตย์มีความแม่นยำสูงสุดโดยไม่ต้องดัดแปลงโครงสร้างฮาร์ดแวร์ให้ซับซ้อน คณะผู้จัดทำได้พัฒนาระเบียบวิธีคำนวณชดเชยทางคณิตศาสตร์อย่างเป็นระบบ 4 ขั้นตอน "
        "(4-Stage Systematic Mathematical Compensation Pipeline) ซึ่งนำสัญญาณจากทั้ง 4 ช่องของไอซี ADS1220 (ได้แก่ V_op, V_bias, V_NTC และ V_TEMP_REF) "
        "มาผสานการทำงานร่วมกันในเฟิร์มแวร์ของไมโครคอนโทรลเลอร์ ESP32-C6 ดังนี้:"
    )
    add_numbered_item(
        doc,
        "ขั้นที่ 1",
        "คำนวณอุณหภูมิของชิปสร้างแรงดันอ้างอิง REF5045 จากแรงดันขา TEMP (V_TEMP_REF ≈ 509 mV ที่ 0 °C หรือ ≈ 575 mV ที่ 25 °C, ความชัน 2.64 mV/°C) "
        "เพื่อชดเชยการเลื่อนของแรงดันอ้างอิง V_REF(T) ตามสัมประสิทธิ์อุณหภูมิ α_REF (3 ppm/°C) ก่อนนำไปคำนวณแปลงรหัสดิจิทัล 24 บิตทุกช่องสัญญาณ:",
        bold_prefix="การชดเชยความแปรปรวนของแรงดันอ้างอิงตามอุณหภูมิ (Reference Drift Compensation): "
    )
    add_equation_box(doc, "T_REF = (V_TEMP_REF - 0.509) / 0.00264  ,   V_REF(T) = V_REF(25°C) · [1 + α_REF · (T_REF - 25.0)]", eq_no="2.5")
    add_numbered_item(
        doc,
        "ขั้นที่ 2",
        "เนื่องจากแรงดัน V_bias ที่ขา +IN ของออปแอมป์ (คำนวณหรือวัดจริงจากช่องสัญญาณ V_bias) จะปรากฏเป็นระดับฐานเดียวกันที่ขาเอาต์พุต V_op ด้วยอัตราขยาย 1.000 เท่าเสมอ "
        "การนำค่าที่อ่านได้จริงทั้งสองช่องมาหักลบกันในซอฟต์แวร์ (Software Pseudo-Differential) พร้อมหักลบค่าออฟเซ็ตตกค้างในความมืด (ΔV_zero_cal) "
        "จะช่วยกำจัดความคลาดเคลื่อนจากความคลาดเคลื่อนของตัวต้านทาน R3, R4 และการแกว่งของแหล่งจ่ายไฟออกไปได้อย่างสมบูรณ์ 100%:",
        bold_prefix="การหักลบแรงดันฐานแบบผลต่างเทียมและการคำนวณกระแสแสงดิบ (Ratiometric Differential Subtraction): "
    )
    add_equation_box(doc, "ΔV_sig = (V_op - V_bias) - ΔV_zero_cal  ,   I_P(raw) = ΔV_sig / R_1", eq_no="2.6")
    add_numbered_item(
        doc,
        "ขั้นที่ 3",
        "คำนวณอุณหภูมิผิวของหัววัด (T_diode ในหน่วย °C) จากวงจรแบ่งแรงดันเทอร์มิสเตอร์ NTC 10 kΩ (B₂₅/₈₅ = 3975 K ตามสเปกชีตรุ่น TTC3A103F39HAEY รหัส 39H) ที่ติดตั้งแนบชิดกับตัวโฟโตไดโอด "
        "แล้วนำไปคำนวณชดเชยคุณสมบัติทางความร้อนของซิลิคอนไดโอด BPW 34 ทั้ง 2 ส่วนจากสเปกชีต หน้า 5 ได้แก่: "
        "(ก) กระแสรั่วในความมืด I_dark(T) ซึ่งเพิ่มขึ้น 2 เท่าทุกๆ อุณหภูมิที่สูงขึ้น 10 °C และ "
        "(ข) ค่าสัมประสิทธิ์อุณหภูมิของความไวกระแสแสง (Temperature Coefficient of Photocurrent, TC_I ≈ +0.18 %/K หรือ α_I = +0.0018 K⁻¹):",
        bold_prefix="การชดเชยกระแสรั่วมืดและสัมประสิทธิ์อุณหภูมิของไดโอด (Photodiode Thermal Compensation): "
    )
    add_equation_box(doc, "I_dark(T) = I_dark(25°C) · 2^((T_diode - 25.0) / 10.0)", eq_no="2.7")
    add_equation_box(doc, "I_P(25°C) = [ I_P(raw) - I_dark(T) ] / [ 1 + α_I · (T_diode - 25.0) ]", eq_no="2.8")
    add_numbered_item(
        doc,
        "ขั้นที่ 4",
        "นำกระแสแสงสุทธิที่ปรับเทียบกลับสู่อุณหภูมิมาตรฐาน 25 °C แล้ว มาหารด้วยค่าความไวสุทธิของหัววัด (S_eff ≈ 2.60 µA ต่อ W/m² หรือค่าคงที่จากการสอบเทียบกับไพราโนมิเตอร์มาตรฐาน) "
        "เพื่อให้ได้ค่าความเข้มรังสีดวงอาทิตย์สุทธิ G_comp ในหน่วย W/m²:",
        bold_prefix="การแปลงเป็นค่าความเข้มรังสีดวงอาทิตย์สุทธิ (Calibrated Solar Irradiance Conversion): "
    )
    add_equation_box(doc, "G_comp [W/m²] = I_P(25°C) / S_eff", eq_no="2.9")

    add_section_heading(doc, "2.4", "คุณสมบัติทางแสงของกระจกแซฟไฟร์ (Sapphire Optical Window)")
    add_thai_body_p(
        doc,
        "แม้ว่า BPW 34 จะมีข้อดีเรื่องระนาบรับแสงแบนราบ แต่ข้อจำกัดคือตัวถังเป็นพลาสติก Epoxy ใส "
        "ซึ่งหากนำไปติดตั้งรับแสงแดดกลางแจ้งโดยตรง รังสีอัลตราไวโอเลต (UV) จะทำให้พลาสติกเสื่อมสภาพ เกิดอาการเหลืองขุ่น (Yellowing) "
        "และผิวหน้าเกิดรอยขีดข่วนจากเม็ดทรายและฝุ่นละออง เพื่อแก้ปัญหานี้ จึงได้เลือกใช้แผ่นกระจกแซฟไฟร์สังเคราะห์ (Synthetic Sapphire, Al₂O₃) "
        "มาติดตั้งครอบเหนือระนาบเซนเซอร์ โดยแซฟไฟร์มีคุณสมบัติเด่นดังนี้:"
    )
    add_bullet_item(doc, bullet_char="•", text="มีความแข็งตามสเกลโมส์ (Mohs Hardness) สูงถึงระดับ 9 ทนทานต่อการขูดขีดกลางแจ้งได้ยาวนานนับสิบปี")
    add_bullet_item(doc, bullet_char="•", text="มีดัชนีหักเหแสง n ≈ 1.76 ให้การส่องผ่านแสง (Optical Transmittance) กว้างครอบคลุม 300 ถึง 4,000 nm")
    add_bullet_item(doc, bullet_char="•", text="สูญเสียจากการสะท้อนที่ผิว (Fresnel Reflection Loss) ประมาณ 7.6% ต่อระนาบหน้าสัมผัส ทำให้มีค่าการส่องผ่านสเปกตรัมที่เสถียรและราบเรียบ (Flat Transmission Curve) ตลอดช่วงความยาวคลื่นที่เซนเซอร์ซิลิคอนตอบสนอง")

    # รูปที่ 2.7 กราฟ Sapphire Transmission
    sapphire_path = os.path.join(img_dir, "sapphire_transmission_spectrum.png")
    if os.path.exists(sapphire_path):
        add_figure_image(
            doc,
            sapphire_path,
            fig_no="2.7",
            title="กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Al₂O₃ Transmission Spectrum)",
            width_inches=5.4
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.7",
            title="กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Transmission Spectrum)",
            description="กราฟแสดงค่าร้อยละการส่องผ่านแสง (% Transmittance) ของกระจกแซฟไฟร์ที่มีค่าเสถียรประมาณ 85–88% ในช่วง 400 ถึง 1,100 nm",
            instructions="วางภาพกราฟ Transmission Curve ของผลึกแซฟไฟร์"
        )

    add_section_heading(doc, "2.5", "ไอซีแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)")
    add_thai_body_p(
        doc,
        "การตรวจวัดกระแสแสงโฟโตไดโอดที่มีการเปลี่ยนแปลงตั้งแต่ระดับ nA (ในเวลารุ่งเช้า/พลบค่ำ) จนถึงระดับหลาย mA (เวลากลางวันแดดจัด) "
        "ต้องการตัวแปลงแอนะล็อกเป็นดิจิทัลที่มี Dynamic Range สูงมาก ชิป Texas Instruments ADS1220 สถาปัตยกรรม 24-bit Delta-Sigma ADC "
        "มีวงจรขยาย Programmable Gain Amplifier (PGA) สัญญาณรบกวนต่ำมากในตัว และมีอินพุตมัลติเพล็กเซอร์ 4 ช่องสัญญาณ "
        "ที่สามารถกำหนดค่าเป็นโหมดวัดเทียบกราวด์แอนะล็อก (Single-Ended vs AVSS) ได้ทั้ง 4 ช่อง หรือโหมดวัดผลต่าง (Differential Input) ได้ 2 คู่ช่องสัญญาณ"
    )

    # รูปที่ 2.8 ADS1220 Block Diagram
    ads_diag = os.path.join(img_dir, "ADS1220_diagram.png")
    if os.path.exists(ads_diag):
        add_figure_image(
            doc,
            ads_diag,
            fig_no="2.8",
            title="บล็อกไดอะแกรมโครงสร้างวงจรภายในของไอซี 24-bit ADC (ADS1220)",
            width_inches=5.2
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.8",
            title="บล็อกไดอะแกรมโครงสร้างวงจรภายในของไอซี 24-bit ADC (ADS1220)",
            description="ภาพบล็อกไดอะแกรมแสดง Input MUX, Low-noise PGA, Delta-Sigma Modulator, Digital Filter และ SPI Interface",
            instructions="วางภาพ Functional Block Diagram จาก Datasheet ADS1220"
        )

    add_section_heading(doc, "2.6", "แหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูง (REF5045) และข้อกำหนดตัวเก็บประจุ")
    add_thai_body_p(
        doc,
        "ความเที่ยงตรงของการแปลงสัญญาณแอนะล็อกเป็นดิจิทัลขึ้นอยู่กับเสถียรภาพของแรงดันอ้างอิง (Reference Voltage, V_REF) "
        "ชิป Texas Instruments REF5045 จ่ายแรงดันอ้างอิง 4.500 V มีค่าสัมประสิทธิ์การเลื่อนตามอุณหภูมิต่ำพิเศษ (Low Drift 3 ppm/°C) "
        "อย่างไรก็ตาม การทำงานของวงจรขยายควบคุมภายใน (Internal Error Amplifier) ต้องการตัวเก็บประจุโหลดภายนอก (C_OUT หรือ C10) "
        "ที่มีค่าความต้านทานอนุกรมสมมูลต่ำ (Low Equivalent Series Resistance: Low-ESR) เพื่อรักษาเสถียรภาพของลูปป้อนกลับ (Loop Phase Margin) "
        "หากตัวเก็บประจุมีค่า ESR สูงเกินไป จะทำให้ลูปเกิดการแกว่ง (Oscillation) และส่งผลให้แรงดันเอาต์พุตเพี้ยนไปจากค่าพิกัด"
    )

    # รูปที่ 2.9 REF5045 Schematic & Layout Guide
    ref_sample = os.path.join(img_dir, "Layout sample for REF5045.png")
    ref_guide = os.path.join(img_dir, "Layout Guide REF5045.png")
    if os.path.exists(ref_sample) and os.path.exists(ref_guide):
        add_two_figure_images_side_by_side(
            doc,
            ref_sample, "(ก) ผังวงจรและการต่อ C10 ชิดขา V_OUT",
            ref_guide, "(ข) แนวทางการจัดวางเส้นทองแดง (Layout Guide)",
            fig_no="2.9",
            main_title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR",
            col_width_inch=2.8
        )
    elif os.path.exists(ref_sample):
        add_figure_image(doc, ref_sample, fig_no="2.9", title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR", width_inches=4.8)
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.9",
            title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR",
            description="แผนผังวงจรแสดงตำแหน่งของตัวเก็บประจุ C10 (10 µF) ชิดขาเอาต์พุต และขา 5 Trim/NR",
            instructions="วางภาพ Schematic วงจรของ REF5045 จาก Datasheet"
        )

    # -------------------------------------------------------------------------
    # บทที่ 3: การออกแบบและวิธีดำเนินงาน (System Design & Methodology)
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "3", "การออกแบบและวิธีดำเนินงาน\n(System Design & Research Methodology)")

    add_section_heading(doc, "3.1", "สถาปัตยกรรมระบบวัดและผังวงจรรวม (System Architecture & Hardware Interfacing)")
    add_thai_body_p(
        doc,
        "ระบบวัดไพราโนมิเตอร์ต้นแบบที่พัฒนาขึ้นในโครงงานนี้ ได้รับการออกแบบเชิงวิศวกรรมให้มีความแม่นยำสูง "
        "โดยแบ่งสถาปัตยกรรมการทำงานออกเป็น 4 ภาคหลักที่ประสานการทำงานร่วมกันอย่างสมบูรณ์ ได้แก่: "
        "(1) ภาคหัววัดสัญญาณแสงและวงจรขยายกระแสแสง (Optical Sensing & Transimpedance Amplifier Front-End) "
        "ประกอบด้วยโฟโตไดโอดซิลิคอน (พิจารณาเปรียบเทียบระหว่าง BPW 34 และ BPX 61) สวมครอบด้วยหน้าต่างกระจกแซฟไฟร์ (Sapphire Al₂O₃) "
        "ทำหน้าที่แปลงความเข้มรังสีแสงอาทิตย์เป็นสัญญาณกระแสแสง (Photocurrent: I_P) จากนั้นส่งผ่านเข้าสู่วงจรขยายกระแสเป็นแรงดัน "
        "Transimpedance Amplifier (TIA) โดยใช้ไอซี Zero-Drift Op-Amp รุ่น OPA380AIDR พร้อมตัวต้านทานป้อนกลับ R1 (R_TIA) "
        "ที่สามารถเลือกค่าได้ตามพิกัดความเข้มแสง (2,200 Ω สำหรับการทดสอบความไวสูงในร่ม หรือ 1,100 Ω / 680 Ω / 390 Ω สำหรับรับแสงอาทิตย์กลางแจ้งสูงสุดถึง 4,000 W/m²) "
        "เพื่อแปลงกระแสโฟโตเคอร์เรนต์ให้อยู่ในรูปแรงดันไฟฟ้า V_op, "
        "(2) ภาคตรวจวัดอุณหภูมิหลายจุดเพื่อการชดเชยทางความร้อน (Multi-Point Thermal Monitoring & Compensation) "
        "ประกอบด้วยการวัดอุณหภูมิผิวของหัววัดแสงด้วยเซนเซอร์ NTC Thermistor ขนาด 10 kΩ (ความแม่นยำ ±1%) รุ่น TTC3A103F39HAEY "
        "(ค่าสัมประสิทธิ์ B₂₅/₈₅ = 3975 K) ติดตั้งแนบชิดกับตัวโฟโตไดโอด ต่ออนุกรมกับตัวต้านทานอ้างอิง 10 kΩ (0.1%) เข้ากับไฟอ้างอิง +4.5Vref "
        "เพื่อนำไปคำนวณชดเชยกระแสรั่วมืด I_dark(T) และค่าสัมประสิทธิ์อุณหภูมิของกระแสแสง (TC_I ≈ +0.18%/K) "
        "ร่วมกับการตรวจวัดอุณหภูมิของไอซีแรงดันอ้างอิง REF5045 ผ่านขา 3 (TEMP) และอุณหภูมิรอยต่อภายในชิป ADC ADS1220, "
        "(3) ภาคแปลงสัญญาณแอนะล็อกเป็นดิจิทัลความละเอียดสูง 24 บิตและเรกูเลเตอร์สัญญาณรบกวนต่ำพิเศษ (24-bit ADC, Precision Reference & Ultra-Low-Noise LDO) "
        "ใช้ไอซีเรกูเลเตอร์ TPS7A4901DGNR แปลงไฟเข้า DC เป็นไฟเลี้ยงแอนะล็อกสะอาด +5.2VLDO ป้อนให้แก่ไอซีสร้างแรงดันอ้างอิง REF5045 "
        "(V_REF = 4.500 V, 3 ppm/°C) ซึ่งจ่ายแรงดันอ้างอิงเข้าที่ขั้ว REFP0 และ REFN0 ของไอซี ADS1220 "
        "โดยกำหนดค่ารีจิสเตอร์ของ ADS1220 ให้ทำงานในโหมด Single-Ended เทียบกับกราวด์แอนะล็อก (AVSS) ทุกช่องสัญญาณ, "
        "ปิดการทำงานของวงจรขยายสัญญาณภายใน (Gain = 1, PGA Bypass) เพื่อรองรับช่วงแรงดันอินพุตเต็มสเกล 0 ถึง 4.500 V "
        "และกำหนดอัตราการสุ่มอ่านข้อมูลที่ 20 SPS (Normal Mode) เพื่อกรองสัญญาณรบกวน 50 Hz/60 Hz จากระบบสายส่งไฟฟ้า, "
        "และ (4) ภาคประมวลผลไมโครคอนโทรลเลอร์และการสื่อสารสู่ Data Logger (ESP32-C6 RISC-V Controller & Data Logger Platform) "
        "ใช้บอร์ดไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini อ่านค่าดิบ (Raw Data) ครบทั้ง 4 ช่องผ่านบัส SPI ส่งผ่าน Wi-Fi / RS-485 ไปยังตัว Data Logger เพื่อรันสมการชดเชย 4 ขั้นตอน, สมการชดเชยมุมตกกระทบ และแสดงผลผ่าน Web Dashboard แบบเรียลไทม์ "
        "ดังแสดงในบล็อกไดอะแกรมและผังวงจรรูปที่ 3.1 โดยมีรายละเอียดการกำหนดช่องสัญญาณมัลติเพล็กเซอร์ (Input MUX) ของไอซี ADS1220 "
        "เปรียบเทียบระหว่างบอร์ดต้นแบบบัดกรีมือ (Perfboard Prototype) และแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom PCB L1_Analog) ดังแสดงในตารางที่ 3.1"
    )

    # ตารางที่ 3.1
    add_table_caption(doc, "3.1", "การกำหนดช่องสัญญาณอินพุตมัลติเพล็กเซอร์ (Input MUX) ของไอซี ADS1220 บนบอร์ด Perfboard และแผ่นวงจรพิมพ์ PCB")
    t_mux = doc.add_table(rows=6, cols=6)
    t_mux.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_mux, border_color="000000", border_sz="4")
    mux_headers = [
        "ช่อง MUX",
        "ค่า Config\nReg 0",
        "สัญญาณบนบอร์ด\nPerfboard จริง",
        "สัญญาณบนบอร์ด\nPCB (L1_Analog)",
        "โหมดการวัด",
        "หน้าที่การทำงานและสมการคำนวณในระบบชดเชย 4 ขั้นตอน"
    ]
    for j, h in enumerate(mux_headers):
        cell = t_mux.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=55, bottom=55, left=45, right=45)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=12, bold=True)

    mux_data = [
        [
            "AIN0",
            "0x81\n(0b10000001)",
            "V_op\n(เอาต์พุต TIA)",
            "V_bias\n(แรงดันฐาน +IN)",
            "Single-Ended\nvs AVSS\nGain=1, Bypass",
            "บนบอร์ด Perfboard ใช้รับสัญญาณเอาต์พุต V_op จากวงจรขยายกระแสแสง OPA380 (ส่วนบนแผ่น PCB สลับใช้รับ V_bias เพื่อความสะดวกในการเดินลายทองแดง)"
        ],
        [
            "AIN1",
            "0x91\n(0b10010001)",
            "V_bias\n(แรงดันฐาน +IN)",
            "V_op\n(เอาต์พุต TIA)",
            "Single-Ended\nvs AVSS\nGain=1, Bypass",
            "นำไปหักลบแบบผลต่างเทียมในซอฟต์แวร์ (ขั้นที่ 2):\nΔV_sig = (V_op - V_bias) - ΔV_zero\nI_P(raw) = ΔV_sig / R_1"
        ],
        [
            "AIN2",
            "0xA1\n(0b10100001)",
            "Temp_REF\n(ขา 3 REF5045)",
            "Temp_REF\n(ขา 3 REF5045)",
            "Single-Ended\nvs AVSS\nGain=1, Bypass",
            "วัดแรงดันขา TEMP ของ REF5045 เพื่อคำนวณอุณหภูมิชิปอ้างอิง:\nT_REF = (V_AIN2 - 0.509) / 0.00264\nและชดเชย V_REF(T) แบบเรียลไทม์ (ขั้นที่ 1)"
        ],
        [
            "AIN3",
            "0xB1\n(0b10110001)",
            "V_NTC\n(เทอร์มิสเตอร์ 10k)",
            "V_NTC\n(เทอร์มิสเตอร์ 10k)",
            "Single-Ended\nvs AVSS\nGain=1, Bypass",
            "วัดแรงดันตกคร่อม NTC คำนวณอุณหภูมิผิวหัววัด T_diode (B25/85 = 3975 K) เพื่อชดเชย I_dark(T) และ TC_I = +0.18%/K (ขั้นที่ 3)"
        ],
        [
            "Internal\nTemp (TS=1)",
            "Reg 1 = 0x02\n(bit TS = 1)",
            "ไดโอดอุณหภูมิ\nภายใน ADS1220",
            "ไดโอดอุณหภูมิ\nภายใน ADS1220",
            "Internal Sensor\n(14-bit Mode)",
            "วัดอุณหภูมิรอยต่อซิลิคอนภายในตัวชิป ADC เพื่อติดตามเสถียรภาพทางความร้อนของระบบ (0.03125 °C ต่อ LSB)"
        ]
    ]
    for i, row in enumerate(mux_data):
        for j, val in enumerate(row):
            cell = t_mux.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=45, right=45)
            p = cell.paragraphs[0]
            if j in [0, 1, 2, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_thai_body_p(
        doc,
        "จากตารางที่ 3.1 ในเฟิร์มแวร์ภาษา C++ (user_arduino_code.cpp) คณะผู้จัดทำได้สร้างตัวแปรสลับโหมดฮาร์ดแวร์ "
        "(#define TARGET_PERFBOARD) เพื่อให้สามารถสลับตำแหน่งช่องสัญญาณระหว่างบอร์ด Perfboard (AIN0 = V_op, AIN1 = V_bias) "
        "และแผ่นวงจรพิมพ์ PCB L1_Analog (AIN0 = V_bias, AIN1 = V_op) ได้ทันทีโดยไม่ต้องแก้ไขสมการคำนวณหลัก "
        "สำหรับการเชื่อมต่อสัญญาณดิจิทัลระหว่างไอซี ADS1220 กับบอร์ดไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini "
        "ใช้บัสสื่อสารอนุกรมความเร็วสูง SPI ที่ความถี่สัญญาณนาฬิกา 1 MHz ในโหมด SPI Mode 1 (CPOL = 0, CPHA = 1) "
        "โดยมีรายละเอียดการกำหนดขาสัญญาณจริงที่ใช้งานบนบอร์ดต้นแบบ ดังแสดงในตารางที่ 3.2"
    )

    # ตารางที่ 3.2
    add_table_caption(doc, "3.2", "การกำหนดการเชื่อมต่อขาสัญญาณระหว่างไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini กับไอซี ADS1220")
    t_pin = doc.add_table(rows=7, cols=5)
    t_pin.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_pin, border_color="000000", border_sz="4")
    pin_headers = ["ขา ESP32-C6", "ตัวแปรในเฟิร์มแวร์", "ขาไอซี ADS1220", "ทิศทางสัญญาณ (I/O)", "หน้าที่และรายละเอียดทางเทคนิค"]
    for j, h in enumerate(pin_headers):
        cell = t_pin.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=13, bold=True)

    pin_data = [
        [
            "GPIO 5",
            "PIN_DRDY",
            "ขา 14 (DRDY)",
            "Input (Pull-up)",
            "สัญญาณแจ้งเตือนข้อมูลพร้อมอ่าน (Active-Low Data Ready) พร้อมระบบตรวจจับ Watchdog Timeout 150 ms และ Auto-Recovery Reset"
        ],
        [
            "GPIO 4",
            "PIN_MISO",
            "ขา 15 (DOUT/DRDY)",
            "Input (ADC -> MCU)",
            "สายรับข้อมูลดิจิทัลอนุกรม (SPI MISO) รับผลการแปลงค่าขนาด 24 บิต (3 ไบต์) จาก ADC เข้าสู่ไมโครคอนโทรลเลอร์"
        ],
        [
            "GPIO 3",
            "PIN_MOSI",
            "ขา 16 (DIN)",
            "Output (MCU -> ADC)",
            "สายส่งคำสั่งควบคุม (SPI MOSI) สำหรับส่งคำสั่งรีเซ็ต (0x06), สั่งเริ่มแปลงค่า (0x08) และตั้งค่ารีจิสเตอร์ WREG"
        ],
        [
            "GPIO 2",
            "PIN_SCLK",
            "ขา 1 (SCLK)",
            "Output (MCU -> ADC)",
            "สัญญาณนาฬิกาบัส SPI ความถี่ 1 MHz โหมด SPI_MODE1 (CPOL=0, CPHA=1) อ่านข้อมูลในจังหวะขอบขาลง (Falling Edge)"
        ],
        [
            "GPIO 1",
            "PIN_CS",
            "ขา 2 (CS)",
            "Output (Active-Low)",
            "สัญญาณเลือกชิป (Chip Select) ควบคุมการเปิด-ปิดเฟรมการสื่อสาร SPI เพื่อป้องกันสัญญาณรบกวนบนบัสข้อมูล"
        ],
        [
            "3V3 / GND",
            "Power Supply",
            "DVDD / DGND, AVSS",
            "Power Lines",
            "แรงดันไฟเลี้ยงภาคดิจิทัล 3.3 V (DVDD) และจุดรวมกราวด์อ้างอิง (Star Ground ร่วมกับ AVSS ของภาคแอนะล็อก)"
        ]
    ]
    for i, row in enumerate(pin_data):
        for j, val in enumerate(row):
            cell = t_pin.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=50, right=50)
            p = cell.paragraphs[0]
            if j in [0, 1, 2, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_thai_body_p(
        doc,
        "ในด้านสถาปัตยกรรมเฟิร์มแวร์บนไมโครคอนโทรลเลอร์ ESP32-C6 ได้รับการออกแบบให้ทำงานในลักษณะ Non-blocking "
        "โดยใช้ฟังก์ชัน millis() กำหนดจังหวะการอ่านข้อมูลทุกๆ 1 วินาที (READ_INTERVAL = 1000 ms) "
        "ในแต่ละรอบการทำงาน เฟิร์มแวร์จะทำการวนรอบสลับอ่านสัญญาณทีละช่องสัญญาณตามลำดับ AIN0 -> AIN1 -> AIN2 -> AIN3 "
        "พร้อมสั่งอ่านเซนเซอร์อุณหภูมิภายในตัวชิป ADS1220 แล้วส่งเข้าสู่ฟังก์ชัน runSystematicCompensationPipeline() "
        "เพื่อประมวลผลสมการชดเชย 4 ขั้นตอนทันที พร้อมทั้งให้บริการหน้าจอ Real-time Web Dashboard (เส้นทาง '/') และ RESTful JSON API (เส้นทาง '/data') "
        "ผ่านเครือข่าย Wi-Fi 2.4 GHz ของชิป ESP32-C6 ดังแสดงผังวงจรเชื่อมต่อในรูปที่ 3.1"
    )

    # รูปที่ 3.1 บล็อกไดอะแกรมระบบและผังวงจร ADS1220
    sys_diag_path = os.path.join(img_dir, "system_block_diagram.png")
    ads_sch_path = os.path.join(img_dir, "ads1220_schematic_kicad.png")
    if os.path.exists(sys_diag_path) and os.path.exists(ads_sch_path):
        add_two_figure_images_stacked(
            doc,
            sys_diag_path, "(ก) บล็อกไดอะแกรมสถาปัตยกรรมระบบวัดไพราโนมิเตอร์และการประมวลผลชดเชย 4 ขั้นตอน",
            ads_sch_path, "(ข) ผังวงจรการเชื่อมต่อไอซี ADS1220, แหล่งจ่ายอ้างอิง REF5045 และวงจรแบ่งแรงดัน NTC บนโปรแกรม KiCad",
            fig_no="3.1",
            main_title="สถาปัตยกรรมระบบวัดไพราโนมิเตอร์ต้นแบบและผังวงจรภาคแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต",
            width_inches=5.6
        )
    elif os.path.exists(sys_diag_path):
        add_figure_image(
            doc,
            sys_diag_path,
            fig_no="3.1",
            title="บล็อกไดอะแกรมสถาปัตยกรรมระบบวัดไพราโนมิเตอร์ต้นแบบ",
            width_inches=5.8
        )

    add_section_heading(doc, "3.2", "การออกแบบโครงสร้าง 3D CAD เคสกระจกแซฟไฟร์ และแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน")
    add_thai_body_p(
        doc,
        "คณะผู้จัดทำได้สร้างแบบจำลอง 3 มิติ (3D CAD Model) ของเซนเซอร์ทั้งสองตระกูลอย่างละเอียดในโปรแกรม FreeCAD ดังแสดงในรูปที่ 3.2 "
        "ได้แก่ เซนเซอร์ BPX 61 ในตัวถังโลหะ TO-39 (ประกอบด้วยส่วน Header, โลหะ Can, Glass Window และแผ่นชิป Active Die ขนาด 2.65 × 2.65 mm²) "
        "และเซนเซอร์ BPW 34 ในตัวถังพลาสติกแบนราบ DIL (แสดงแผ่นชิปซิลิคอนขนาด 2.65 × 2.65 mm² วางใกล้ผิวหน้าด้านบนและขาโลหะทั้งสองข้าง) "
        "เพื่อนำไปใช้ในการจำลองแสงทางทัศนศาสตร์และการออกแบบกล่องเคส"
    )

    # รูปที่ 3.2 โมเดล CAD ของ BPX 61 และ BPW 34
    cad1_p = os.path.join(img_dir, "CAD1.png")
    cad_bpw34_p = os.path.join(img_dir, "cad_bpw34_model.png")
    if os.path.exists(cad1_p) and os.path.exists(cad_bpw34_p):
        add_two_figure_images_side_by_side(
            doc,
            cad1_p, "(ก) โมเดล 3D CAD ของเซนเซอร์ BPX 61 (ตัวถังโลหะ TO-39 ลึก 1.55–1.75 mm)",
            cad_bpw34_p, "(ข) โมเดล 3D CAD ของเซนเซอร์ BPW 34 (ตัวถังแบนราบ DIL ชิปขนาด 2.65×2.65 mm²)",
            fig_no="3.2",
            main_title="แบบจำลอง 3D CAD เปรียบเทียบโครงสร้างทางเรขาคณิตระหว่างเซนเซอร์ BPX 61 และ BPW 34",
            col_width_inch=2.8
        )

    add_thai_body_p(
        doc,
        "นอกจากนี้ จากการวิเคราะห์การเสื่อมสภาพของตัวถังเรซิน Epoxy (Photo-oxidation จากรังสี UV และการขูดขีดจากฝุ่นทราย) "
        "และขาโลหะที่มีส่วนผสมของเงินซึ่งไวต่อความชื้น คณะผู้จัดทำจึงออกแบบกล่องเคสต้นแบบ 3 มิติแบบปิดผนึก (Custom 3D-Printed Enclosure) "
        "สำหรับแผงวงจรขนาด 70×90 mm พร้อมฝาแหวนล็อกกระจกหน้าต่างแซฟไฟร์สังเคราะห์ (Synthetic Sapphire Window, ความแข็ง 9 Mohs, "
        "การส่องผ่าน ~85% ช่วง 300–1100 nm) แบบมีร่องขันล็อก 6 จุด (Sapphire Retaining Bezel Ring) และช่องร้อยสายไฟกันน้ำด้านข้างดังแสดงในรูปที่ 3.3"
    )

    # รูปที่ 3.3 โมเดล 3D CAD ของกล่องเคส Perfboard และแหวนล็อกกระจกแซฟไฟร์
    cad_case_p = os.path.join(img_dir, "cad_perfboard_case_assy.png")
    cad_ring_p = os.path.join(img_dir, "cad_sapphire_retaining_ring.png")
    if os.path.exists(cad_case_p) and os.path.exists(cad_ring_p):
        add_two_figure_images_side_by_side(
            doc,
            cad_case_p, "(ก) โครงสร้างกล่องเคสสำหรับบอร์ด Perfboard พร้อมช่องหน้าต่างรับแสงทรงกลมด้านบน",
            cad_ring_p, "(ข) แหวนกดล็อกหน้าต่างกระจกแซฟไฟร์ (Retaining Bezel Ring)",
            fig_no="3.3",
            main_title="แบบจำลอง 3D CAD ของกล่องบรรจุวงจรต้นแบบ Perfboard และชุดยึดหน้าต่างกระจกแซฟไฟร์",
            col_width_inch=2.8
        )

    add_thai_body_p(
        doc,
        "สำหรับการพัฒนาฮาร์ดแวร์ระดับผลิตจริงเพื่อขจัดปัญหาความต้านทานแฝงและสัญญาณรบกวนจากการต่อสายลอยบนบอร์ดทดลอง "
        "คณะผู้จัดทำได้ออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom 4-Layer Circular Sensor Head PCB) รุ่น 'V1.1 L1 by Tachin' บนโปรแกรม KiCad 10.0 "
        "โดยจัดสรรโครงสร้างแบบคู่ขนาน 2 รูปแบบ ได้แก่ รุ่น L1_Analog (สำหรับ BPX 61 ในตัวถัง TO-39 Can) และรุ่น L1_Analog_BPW34 (สำหรับ BPW 34 ในตัวถัง Flat DIL Plastic) "
        "ทั้งนี้ จากการตรวจสอบเอกสารข้อมูลทางเทคนิคของ ams OSRAM ยืนยันว่าเซนเซอร์ทั้งสองรุ่นใช้เนื้อชิปซิลิคอน (Active Chip Die) ขนาดเดียวกันเป๊ะ "
        "คือ 2.65 × 2.65 mm² (พื้นที่รับแสง 7.02 mm²) จึงเป็นการทดลองแบบควบคุมตัวแปร (Controlled Experiment) เพื่อเปรียบเทียบผลของตัวถังและหน้าต่างทางทัศนศาสตร์อย่างแท้จริง "
        "โดยทั้งสองรุ่นผ่านการตรวจ DRC 100% ไร้ข้อผิดพลาด (0 Violations / 0 Unconnected items) ดังแสดงในรูปที่ 3.4"
    )

    # รูปที่ 3.4 ภาพเรนเดอร์ 3D PCB V1.1 L1 (BPX 61 vs BPW 34)
    p_bpx = os.path.join(img_dir, "pcb_3d_render_iso.png")
    p_bpw = os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png")
    if os.path.exists(p_bpx) and os.path.exists(p_bpw):
        add_two_figure_images_side_by_side(
            doc,
            p_bpx, "(ก) รุ่นที่ 1: ตัวถังโลหะ TO-39 Can (BPX 61)",
            p_bpw, "(ข) รุ่นที่ 2: ตัวถังพลาสติกแบน Flat DIL (BPW 34)",
            fig_no="3.4",
            main_title="ภาพจำลอง 3 มิติของแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน (Custom 4-Layer PCB: BPX 61 vs BPW 34)",
            col_width_inch=2.85
        )

    add_section_heading(doc, "3.3", "การออกแบบข้อเสนอระเบียบวิธีการทดลองจริงทางทัศนศาสตร์")
    add_thai_body_p(
        doc,
        "เนื่องจากการจำลองแสงด้วยคอมพิวเตอร์มีข้อจำกัดในการจำลองการสะท้อนหลายชั้น (Multiple Internal Reflections) และการกระเจิงที่ผิวสัมผัสจริง "
        "คณะผู้จัดทำจึงได้จัดทำ 'ข้อเสนอระเบียบวิธีการทดลอง (Proposed Experimental Methodology)' เพื่อนำไปทดสอบจริงในห้องปฏิบัติการฟิสิกส์ คณะวิทยาศาสตร์ สจล. "
        "ในระยะถัดไปของโครงงาน (Capstone Project 2) โดยครอบคลุมทั้งแนวทางในอุดมคติและแนวทางสำรองเชิงปฏิบัติ (Practical Contingency Setup) ดังนี้:"
    )

    add_subsection_heading(doc, "3.3.1", "ชุดการทดลองที่ 1: การวัดค่าการส่องผ่านสเปกตรัมด้วย Vernier LabQuest 2 Spectrometer (350–900 nm)")
    add_thai_body_p(
        doc,
        "การทดสอบนี้มีวัตถุประสงค์เพื่อตรวจสอบคุณสมบัติการส่องผ่านแสง (Spectral Transmittance: %T) ของแผ่นกระจกแซฟไฟร์สังเคราะห์ "
        "(Synthetic Sapphire Optical Window) ความหนา 1.0 mm เทียบกับอากาศเปล่า (Baseline) ว่ามีค่าการส่องผ่านคงที่และราบเรียบ (Flat) "
        "ครอบคลุมย่านความไวแสงของซิลิคอนไดโอด BPW 34 หรือไม่ โดยจากการสำรวจเครื่องมือในห้องปฏิบัติการภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล. "
        "พบว่ามีชุดตรวจวัดสเปกตรัม 'Vernier LabQuest 2 Connected Science System' พร้อมโพรบใยแก้วนำแสง (Optical Fiber) ช่วงคลื่น 350.1–899.7 nm (รูปที่ 3.5) โดยมีระเบียบวิธีทดสอบดังนี้:"
    )
    add_numbered_item(doc, "1.", "เปิดเครื่อง Vernier LabQuest 2 เชื่อมต่อหัววัดสเปกโตรมิเตอร์และโพรบใยแก้วนำแสง (Optical Fiber Probe)", bold_prefix="การเตรียมเครื่องมือ: ")
    add_numbered_item(doc, "2.", "จัดแนวหัวโพรบใยแก้วนำแสงรับแสงคงที่จากหลอดกำเนิดแสง และบันทึกสเปกตรัมเส้นฐานอ้างอิง I_0(λ) (100% Baseline) ในช่วง 350.1 ถึง 899.7 nm", bold_prefix="การวัดเส้นอ้างอิง (Baseline): ")
    add_numbered_item(doc, "3.", "นำแผ่นกระจกแซฟไฟร์สังเคราะห์มาติดตั้งกั้นขวางหน้าปลายโพรบใยแก้วนำแสงในแนวตั้งฉาก แล้วบันทึกสเปกตรัมแสงที่ส่องผ่าน I(λ)", bold_prefix="การวัดชิ้นงานจริง: ")
    add_numbered_item(doc, "4.", "คำนวณค่าร้อยละการส่องผ่าน %T(λ) = [I(λ) / I_0(λ)] × 100% และส่งออกไฟล์ตาราง (.csv) เพื่อเทียบกับค่า 85–88% ในแบบจำลอง", bold_prefix="การประมวลผลข้อมูล: ")

    labquest_img_path = os.path.join(img_dir, "labquest2_vernier_spectrometer.png")
    if not os.path.exists(labquest_img_path):
        labquest_img_path = os.path.join(script_dir, "labquest2_vernier_spectrometer.png")
    if os.path.exists(labquest_img_path):
        add_figure_image(
            doc,
            img_path=labquest_img_path,
            fig_no="3.5",
            title="หน้าจอแสดงผลการวัดสเปกตรัมความยาวคลื่น 350.1–899.7 nm ของชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์",
            width_inches=5.2
        )

    add_subsection_heading(doc, "3.3.2", "ชุดการทดลองที่ 2: การวัดการตอบสนองเชิงมุมบนแท่นหมุน (Directional Cosine Response Test)")
    add_thai_body_p(
        doc,
        "วัตถุประสงค์เพื่อเก็บกราฟ Cosine Error จริงของหัววัด BPW 34 (ทั้งกรณีมีและไม่มีกระจกแซฟไฟร์) ที่มุมตกกระทบ 0° ถึง 85° "
        "เพื่อนำไปสร้างฟังก์ชันชดเชยเชิงมุมร่วมกับพิกัดเวลาและมุมดวงอาทิตย์จริง ระบบการทดสอบออกแบบเป็นส่วนประกอบหลักดังนี้ (รูปที่ 3.6):"
    )
    add_thai_body_p(
        doc,
        "1) แท่นหมุนปรับมุมองศา (Rotation Stage): อาศัยแท่นหมุนเชิงแสงแบบมือหมุน (Manual Optical Rotation Stage พร้อมสเกล Vernier) "
        "โดยจัดทำชุดยึดหัววัด (Test Jig) ให้ระนาบผิวรับแสงของชิปไดโอดและกระจกแซฟไฟร์วางอยู่ในแนวแกนหมุนกึ่งกลางพอดี (Co-axial Alignment) "
        "เพื่อป้องกันการเยื้องศูนย์ขณะปรับมุมตั้งแต่ 0° ถึง 85°"
    )
    add_thai_body_p(
        doc,
        "2) การจัดเตรียมแหล่งกำเนิดแสงและข้อเสนอลำแสงขนาน (Light Source & Collimation Strategy): วางแนวทางสำรองไว้ 2 รูปแบบตามความพร้อมของห้องปฏิบัติการ:"
    )
    add_bullet_item(doc, bullet_char="•", text="แนวทางตามอุดมคติ (Ideal Scenario): ใช้ชุดเครื่องจำลองแสงอาทิตย์ (Solar Simulator) หรือแหล่งกำเนิดแสงซีนอนร่วมกับเลนส์คอลลิเมเตอร์")
    add_bullet_item(doc, bullet_char="•", text="แนวทางสำรองเชิงปฏิบัติ (Practical Contingency / DIY Setup): ใช้หลอดทังสเตน-ฮาโลเจน 12V 50W/100W จ่ายไฟด้วย Bench Power Supply กระแสตรงคงที่ (Blackbody Radiation 2,800–3,000 K สเปกตรัมต่อเนื่อง 350–2,000+ nm) วางห่าง d ≥ 1.5–2.0 m ให้มุมเบี่ยงเบนลำแสงบนชิป 2.65 mm มีค่าเพียง ~0.076° (เสมือนลำแสงขนาน) พร้อมท่อบังแสง Baffle ในห้องมืด")

    blender_persp = os.path.join(img_dir, "goniometer_blender_perspective.png")
    blender_top = os.path.join(img_dir, "goniometer_blender_topdown.jpg")
    if os.path.exists(blender_persp) and os.path.exists(blender_top):
        add_two_figure_images_stacked(
            doc,
            blender_persp, "(ก) แบบจำลองสามมิติ (Perspective View) แสดงการจัดวางโคมหลอดทังสเตน ท่อ Baffle และแท่นหมุนบน Optical Table",
            blender_top, "(ข) แบบจำลองมุมมองระนาบบน (Top-down Plan View) แสดงแนวลำแสงขนานและการวัดมุมตกกระทบ θ",
            fig_no="3.6",
            main_title="แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุม (Goniometer Cosine Test Jig) พร้อมข้อเสนอชุดกำเนิดแสงแสวงเครื่อง",
            width_inches=5.6
        )

    # -------------------------------------------------------------------------
    # บทที่ 4: ผลการจำลองและทดสอบต้นแบบ (Simulation & Experimental Results)
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "4", "ผลการจำลองและทดสอบต้นแบบ\n(Simulation and Experimental Results)")

    add_section_heading(doc, "4.1", "ผลการจำลองทางทัศนศาสตร์และการค้นพบความผิดเพี้ยนจากกระจกหน้าต่าง")
    add_thai_body_p(
        doc,
        "คณะผู้จัดทำได้นำแบบจำลอง 3 มิติของเซนเซอร์ทั้ง BPX 61 และ BPW 34 จากบทที่ 3 ไปจำลองการเดินทางของลำแสงด้วยเอ็นจิน Mitsuba 3 "
        "(Optical Ray-tracing Simulation) ในช่วงมุมตกกระทบ 0° ถึง 90° และค้นพบข้อเท็จจริงทางทัศนศาสตร์ที่สำคัญยิ่งดังนี้:"
    )

    add_subsection_heading(doc, "4.1.1", "การจำลอง BPX 61: ความผิดเพี้ยนจากการหักเหของกระจกหน้าต่าง Schott 8250")
    add_thai_body_p(
        doc,
        "เมื่อจำลองพฤติกรรมเชิงมุมของ BPX 61 เปรียบเทียบกับเส้นกราฟใน Datasheet ของ ams-OSRAM พบว่า (รูปที่ 4.1):"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="กรณีจำลองเฉพาะกระป๋องเปล่าและแผ่นชิป (Without glass only can and die): กราฟการตอบสนองเชิงมุมของแบบจำลอง (เส้นสีน้ำเงิน) "
        "สอดคล้องกับเส้นสเปกใน Datasheet (เส้นสีเขียว) เป็นอย่างดี และเกิดจุดตัดแสง (Cut-off) ที่มุมประมาณ 62 องศา ใกล้เคียงกับสเปกชีต"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="กรณีจำลองแบบใส่กระจกหน้าต่างของกระป๋อง (With glass + can + die): เมื่อกำหนดคุณสมบัติของกระจกตามสเปก SCHOTT Glass 8250 "
        "(Kovar Sealing Borosilicate, ค่าดัชนีหักเห int_ior = 1.487) พบว่าการหักเหแสงที่ผิวแก้วทำให้แสงในมุมเอียงถูกเบนทิศทาง "
        "จนกระทบขอบด้านในของกระป๋องโลหะเร็วขึ้นอย่างมาก ส่งผลให้สัญญาณแสงตกฮวบและถูกตัดแสงสิ้นเชิงตั้งแต่ที่มุม 44 องศา! "
        "ซึ่งเกิดความคลาดเคลื่อนเพี้ยนไปจากสเปกของผู้ผลิตอย่างมหาศาล"
    )

    img_bpx_path = os.path.join(script_dir, "sim_bpx61_with_without_glass.png")
    if os.path.exists(img_bpx_path):
        add_figure_image(
            doc,
            img_bpx_path,
            fig_no="4.1",
            title="ผลการจำลองการตอบสนองเชิงมุมของ BPX 61: กรณีไม่มีกระจก เทียบกับ กรณีมีกระจก SCHOTT 8250",
            width_inches=5.8
        )

    add_subsection_heading(doc, "4.1.2", "การจำลอง BPW 34: การตอบสนองเชิงมุมแบบแบนราบและการประเมิน ISO 9060")
    add_thai_body_p(
        doc,
        "เมื่อเปลี่ยนมาจำลองโฟโตไดโอดแบบแบนราบ BPW 34 ซึ่งใช้ชิปซิลิคอนขนาด 7.02 mm² เดียวกัน ผลการจำลองแสดงข้อได้เปรียบที่เด่นชัด (รูปที่ 4.2):"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="BPW 34 แบบไม่มีกระจกครอบ (without flat glass): เส้นการตอบสนองเชิงมุมสอดคล้องกับกฎโคไซน์อุดมคติอย่างยอดเยี่ยม "
        "ผลการประเมิน Directional Error ตาม ISO 9060 พบว่าความคลาดเคลื่อนมีค่าแทบจะเป็นศูนย์ (ใกล้เคียง 0 W/m²) ตลอดช่วงมุม 0 ถึง 80 องศา "
        "ผ่านเกณฑ์มาตรฐาน Class A (±10 W/m²) อย่างสมบูรณ์แบบ"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="BPW 34 เมื่อติดตั้งกระจกแซฟไฟร์แบน (with Flat Sapphire Glass): การตอบสนองที่มุม 0 ถึง 50 องศายังคงดีเยี่ยม "
        "แต่ที่มุมชันเกิน 55 องศาขึ้นไป การสะท้อนที่ผิวแก้วแซฟไฟร์ (Fresnel Reflection) จะเพิ่มสูงขึ้นอย่างรวดเร็ว "
        "ส่งผลให้ค่า Error ในแบบจำลองตกลงไปแตะระดับ -80 W/m² ในช่วงมุม 60 ถึง 85 องศา"
    )

    img_bpw_path = os.path.join(script_dir, "sim_bpw34_with_without_glass.png")
    if os.path.exists(img_bpw_path):
        add_figure_image(
            doc,
            img_bpw_path,
            fig_no="4.2",
            title="ผลการจำลองการตอบสนองเชิงมุมและ Heatmap ของ BPW 34 พร้อมการประเมินความคลาดเคลื่อน ISO 9060",
            width_inches=5.8
        )

    add_thai_body_p(
        doc,
        "จากการเปรียบเทียบพฤติกรรมเชิงมุมของทั้ง BPX 61 และ BPW 34 ภายใต้เงื่อนไขที่มีและไม่มีกระจกหน้าต่าง "
        "สามารถสรุปผลการตอบสนองเชิงมุมสัมพัทธ์ (Relative Angular Response) ที่มุมตกกระทบสำคัญต่างๆ ได้ดังแสดงในตารางที่ 4.1"
    )

    add_table_caption(doc, "4.1", "การเปรียบเทียบผลการจำลองเชิงมุม: BPX 61 vs BPW 34")
    t_sim_cmp = doc.add_table(rows=6, cols=4)
    t_sim_cmp.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_sim_cmp, border_color="000000", border_sz="4")
    scmp_headers = ["มุมตกกระทบ (θ)", "BPX 61 (ไม่มีกระจก)", "BPX 61 (มีกระจก Schott)", "BPW 34 (มีกระจกแซฟไฟร์)"]
    for j, h in enumerate(scmp_headers):
        cell = t_sim_cmp.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    scmp_data = [
        ["0° (ตั้งฉาก)", "1.000 (ปกติ)", "1.000 (ปกติ)", "1.000 (ปกติ)"],
        ["30°", "0.852 (ตามสเปก)", "0.830 (เริ่มเบี่ยงเบน)", "0.860 (ตามกฎโคไซน์)"],
        ["44°", "0.680 (รับแสงได้)", "0.000 (ตัดแสง 100% เพี้ยนหนัก)", "0.710 (รับแสงได้ปกติ)"],
        ["62°", "0.050 (เริ่มตัดแสงตามสเปก)", "0.000 (มืดสนิท)", "0.450 (รับแสงได้ปกติ)"],
        ["75°", "0.000 (มืดสนิท)", "0.000 (มืดสนิท)", "0.220 (มี Fresnel Reflection Drop)"]
    ]
    for i, row in enumerate(scmp_data):
        for j, val in enumerate(row):
            cell = t_sim_cmp.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=13)

    add_subsection_heading(doc, "4.1.3", "ข้อสรุปทางวิศวกรรม: เหตุผลที่ต้องเก็บกราฟ Cosine Error จริงในห้องปฏิบัติการ")
    add_thai_body_p(
        doc,
        "ผลการจำลองพิสูจน์ชัดเจนว่าแม้ BPW 34 จะแก้ปัญหาการบดบังของขอบกระป๋องโลหะได้ดีกว่า BPX 61 อย่างเด็ดขาด "
        "แต่การติดตั้งกระจกหน้าต่างแซฟไฟร์ยังก่อให้เกิด Fresnel Reflection Drop ที่มุมตกกระทบสูง (> 55°) "
        "ซึ่งจำเป็นต้องนำหัววัดจริงเข้าทดสอบเก็บกราฟมุมตกกระทบในห้องปฏิบัติการฟิสิกส์ตามระเบียบวิธีในบทที่ 3 "
        "เพื่อใช้สร้างสมการชดเชยเชิงมุมร่วมกับตำแหน่งดวงอาทิตย์ในโครงงาน 2 ต่อไป"
    )

    add_section_heading(doc, "4.2", "ผลการสร้างบอร์ดต้นแบบ Perfboard และการแก้ปัญหาภาคจ่ายไฟจริง")
    add_thai_body_p(
        doc,
        "คณะผู้จัดทำได้บัดกรีประกอบวงจรลงบนแผงวงจรอเนกประสงค์ (Perfboard Prototype ขนาด 70×90 mm) ครบทั้งภาคจ่ายไฟ "
        "LM2596 + TPS7A4901, แหล่งจ่ายแรงดันอ้างอิง REF5045, วงจรขยายกระแสแสง OPA380 + BPW 34, เซนเซอร์ NTC (TTC3A103F39HAEY) "
        "และไอซี ADC 24 บิต ADS1220 เชื่อมต่อกับ ESP32-C6 Super Mini โดยจัดวางแยกด้านบนสำหรับรับแสง 180° และด้านล่างสำหรับวงจร (รูปที่ 4.3)"
    )

    pf_top = os.path.join(img_dir, "perfboard_top_view.jpg")
    pf_side = os.path.join(img_dir, "perfboard_side_ntc_view.jpg")
    pf_bot = os.path.join(img_dir, "perfboard_bottom_modules.jpg")
    if os.path.exists(pf_top) and os.path.exists(pf_bot):
        add_two_figure_images_side_by_side(
            doc,
            pf_top, "(ก) ด้านบนบอร์ด (70×90 mm): ติดตั้งเฉพาะหัววัดแสงและเทอร์มิสเตอร์ NTC เพื่อเปิดมุมรับแสง 180°",
            pf_bot, "(ข) ด้านล่างบอร์ด: จัดวาง LM2596, TPS7A4901, REF5045, OPA380, ADS1220 และ ESP32-C6",
            fig_no="4.3",
            main_title="ภาพถ่ายจริงของบอร์ดต้นแบบ Perfboard ที่บัดกรีสมบูรณ์แบบโครงสร้างแยกด้านรับแสงและด้านวงจร",
            col_width_inch=2.8
        )
        if os.path.exists(pf_side):
            add_figure_image(
                doc,
                pf_side,
                fig_no="4.3 (ต่อ)",
                title="มุมมองด้านข้างแสดงการดัดหัววัด NTC Thermistor (สีเขียว) แนบชิดใต้ฐานของตัวเซนเซอร์เพื่อวัดอุณหภูมิจริง",
                width_inches=4.6
            )

    add_subsection_heading(doc, "4.2.1", "การแก้ปัญหา High-Z Node Leakage ของ REF5045 และขา 9 Thermal Pad ของ TPS7A4901")
    add_thai_body_p(
        doc,
        "ประเด็นที่ 1 (High-Z Node Leakage บนขา Trim/NR ของ REF5045): ในการต่อวงจรช่วงแรก พบว่าแรงดันเอาต์พุตของ REF5045 ตกลงเหลือ 4.382 V "
        "(ต่ำกว่าพิกัด 4.500 V อยู่ 118 mV) เมื่อตรวจสอบพบว่าขา 5 (Trim/NR) เป็นโหนดความต้านทานสูงมาก (High-Z Node) ซึ่งไวต่อกระแสรั่วไหล "
        "จึงแก้ไขโดยการปลดการเชื่อมต่อขา 5 ให้ลอยตัว (No Connect / DNP) ร่วมกับการต่อ C10 (Low-ESR 10 µF) ชิดขาเอาต์พุต"
    )
    add_thai_body_p(
        doc,
        "ประเด็นที่ 2 (การค้นพบสาเหตุเรกูเลเตอร์ TPS7A4901DGNR ไม่ทำงานบนแผ่นแปลง MSOP-8): เมื่อป้อนไฟ +9V ถึง +12V เข้า TPS7A4901DGNR "
        "พบว่าแรงดันไฟออก +5.2VLDO ตกเหลือน้อยกว่า 0.8 V จากการตรวจสอบ Datasheet พบว่าแผ่นระบายความร้อนใต้ท้องชิป หรือ 'ขา 9 (Exposed Thermal Pad)' "
        "ต้องต่อลงกราวด์ (GND) เสมอเพื่อให้วงจรอ้างอิงภายในทำงานได้ครบวงจร เนื่องจากแผ่นแปลง MSOP-8 ทั่วไปไม่มีแผ่นทองแดงใต้ท้องชิป "
        "คณะผู้จัดทำจึงนำชิปไปบัดกรีทดสอบลงบนแผ่น PCB รุ่นแรกที่มี Footprint ขา 9 ลงกราวด์โดยตรงเพื่อพิสูจน์การทำงาน"
    )

    add_subsection_heading(doc, "4.2.2", "ผลวัดเสถียรภาพแรงดันภาคจ่ายไฟจริงด้วยมัลติมิเตอร์")
    add_thai_body_p(
        doc,
        "ภายหลังการแก้ไขวงจรและจ่ายไฟผ่านสถาปัตยกรรม 2 ชั้น (LM2596 ปรับตั้ง ~5.5 V ➔ LDO TPS7A4901 ระดับ +5.20 V เพื่อสร้าง Headroom 700 mV เหนือ 4.500 V) "
        "ผลการวัดแรงดันจริงด้วยมัลติมิเตอร์ยืนยันว่า TPS7A4901 บน PCB รุ่นแรกจ่ายแรงดันออกได้จริง 5.194 V (HI: 5.201 V) ตรงตามพิกัด +5.20 V "
        "และ REF5045 บนบอร์ด Perfboard จ่ายแรงดันอ้างอิงได้จริง 4.495 V (HI: 4.496 V) ตรงตามพิกัด +4.500 V โดยนิ่งสนิทแม้ขณะส่ง Wi-Fi ดังสรุปในตารางที่ 4.2 และรูปที่ 4.4"
    )

    add_table_caption(doc, "4.2", "การเปรียบเทียบเสถียรภาพแรงดันอ้างอิง REF5045 และการวัดแรงดันจริงภายใต้เงื่อนไขแหล่งจ่ายไฟต่างๆ")
    t_esr = doc.add_table(rows=4, cols=4)
    t_esr.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_esr, border_color="000000", border_sz="4")
    esr_headers = ["แหล่งจ่ายไฟ / เงื่อนไขวงจร", "แรงดันไฟเลี้ยง (V_IN)", "แรงดันอ้างอิง (V_REF)", "ผลการสังเกตและพฤติกรรมทางไฟฟ้า"]
    for j, h in enumerate(esr_headers):
        cell = t_esr.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    esr_data = [
        ["1. พอร์ต USB คอมพิวเตอร์โดยตรง", "4.55 - 4.68 V", "~4.445 - 4.450 V (แกว่ง)", "V_IN < 4.70 V ชิปติดสภาวะ Dropout เอาต์พุตแกว่งตามโหลด PC"],
        ["2. แหล่งจ่ายสวิตชิ่งปรับค่า 5.2 V", "5.200 V คงที่", "4.491 V นิ่งสนิท", "แรงดันตกคร่อม 700 mV (> V_DO 200 mV) คุมแรงดันนิ่งแม้ส่ง Wi-Fi"],
        ["3. TPS7A4901 (PCB รุ่นแรก) & REF5045 (Perfboard)", "5.194 V (HI: 5.201 V)", "4.495 V (HI: 4.496 V)", "ต่อขา 9 EP ลง GND บน PCB รุ่นแรก + ปลดขา 5 + C Low-ESR 10 µF เสถียรสมบูรณ์"]
    ]
    for i, row in enumerate(esr_data):
        for j, val in enumerate(row):
            cell = t_esr.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0 or j == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=13)

    meas_tps_p = os.path.join(img_dir, "meas_tps7a4901_5v194.jpg")
    meas_ref_p = os.path.join(img_dir, "meas_ref5045_4v495.jpg")
    if os.path.exists(meas_tps_p) and os.path.exists(meas_ref_p):
        add_two_figure_images_side_by_side(
            doc,
            meas_tps_p, "(ก) ทดสอบ TPS7A4901 บน PCB รุ่นแรก (มี Footprint ขา 9): วัดได้ 5.194 V",
            meas_ref_p, "(ข) ทดสอบ REF5045 บน Perfboard (ปลดขา 5 + C Low-ESR): วัดได้ 4.495 V",
            fig_no="4.4",
            main_title="ภาพถ่ายผลการวัดแรงดันภาคจ่ายไฟและแรงดันอ้างอิงจริงด้วยมัลติมิเตอร์",
            col_width_inch=2.8
        )

    add_section_heading(doc, "4.3", "ผลการทดสอบอ่านสัญญาณจริง 24 บิตใน 4 สภาวะแสง")
    add_thai_body_p(
        doc,
        "ภายหลังจากการแก้ไขภาคจ่ายไฟและประกอบวงจรขยาย TIA (ใช้ค่าเริ่มต้น R1 = 2,200 Ω, R3 = 100 kΩ, R4 = 10 kΩ) ลงบนบอร์ด Perfboard เรียบร้อยแล้ว "
        "คณะผู้จัดทำได้ทดสอบอ่านค่าจริงครบทุกช่องสัญญาณจากฝั่ง ESP32-C6 ผ่านพอร์ต Serial Monitor ภายใต้สภาวะแสง 4 ระดับ "
        "ตั้งแต่สภาวะมืดสนิท แสงสว่างในห้อง โคมไฟระยะใกล้ ไปจนถึงการส่องไฟฉายแรงสูงจ่อชิดหน้าเซนเซอร์ ดังแสดงผลการวัดจริงในตารางที่ 4.3"
    )

    add_table_caption(doc, "4.3", "ผลการอ่านค่าแรงดันและอุณหภูมิจริงจากบอร์ดต้นแบบ Perfboard (กรณี R1 = 2,200 Ω, R4 = 10 kΩ)")
    t_real = doc.add_table(rows=5, cols=7)
    t_real.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_real, border_color="000000", border_sz="4")
    real_headers = [
        "สภาวะแสงทดสอบจริง",
        "AIN0: V_op\n(V)",
        "AIN1: V_bias\n(V)",
        "AIN2: Temp_REF\n(V / อุณหภูมิ)",
        "AIN3: V_NTC\n(V / อุณหภูมิ)",
        "ΔV_sig สุทธิ\n(mV)",
        "I_P คำนวณ /\nความเข้มแสง G"
    ]
    for j, h in enumerate(real_headers):
        cell = t_real.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=55, bottom=55, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=12, bold=True)

    real_data = [
        [
            "1. ปิดทึบมืดสนิท\n(Dark Baseline)",
            "0.4185 V",
            "0.4180 V",
            "0.5940 V\n(32.20 °C)",
            "2.0942 V\n(28.48 °C)",
            "+0.50 mV\n(Zero Cal → 0.00)",
            "0.227 µA\n(~0.09 W/m²)"
        ],
        [
            "2. แสงสว่างภายในห้อง\n(Indoor Ambient)",
            "0.4405 V",
            "0.4180 V",
            "0.5942 V\n(32.27 °C)",
            "2.0940 V\n(28.49 °C)",
            "+22.50 mV",
            "10.227 µA\n(~3.93 W/m²)"
        ],
        [
            "3. โคมไฟส่องสว่างระยะใกล้\n(Bright Desk Lamp)",
            "1.5420 V",
            "0.4180 V",
            "0.5948 V\n(32.50 °C)",
            "2.0885 V\n(28.62 °C)",
            "+1,124.0 mV",
            "510.91 µA\n(~196.5 W/m²)"
        ],
        [
            "4. ส่องไฟฉายแรงสูงจ่อชิด\n(Flashlight Saturation)",
            "4.4998 V\n(ชนเพดาน V_REF)",
            "0.4181 V",
            "0.5951 V\n(32.61 °C)",
            "2.0810 V\n(28.80 °C)",
            "+4,081.7 mV\n(ADC Full-Scale)",
            "1,855.3 µA\n(อิ่มตัวที่ ~714 W/m²)"
        ]
    ]
    for i, row in enumerate(real_data):
        for j, val in enumerate(row):
            cell = t_real.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=40, right=40)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_thai_body_p(
        doc,
        "ข้อมูลจากการทดลองจริงในตารางที่ 4.3 พิสูจน์ความถูกต้องของการวิเคราะห์ทางทฤษฎีในบทที่ 2 ทุกประการ ได้แก่: "
        "(1) แรงดันฐาน V_bias ที่วัดได้จริงจากช่อง AIN1 มีค่านิ่งสนิทที่ 0.4180 V (ใกล้เคียงค่าคำนวณตามทฤษฎี 0.4091 V จากเกณฑ์ความคลาดเคลื่อนของตัวต้านทาน) "
        "และในสภาวะมืดสนิท แรงดันเอาต์พุต V_op ของ OPA380 อยู่ที่ 0.4185 V ซึ่งพ้นจากพื้นอิ่มตัว 100 mV ของออปแอมป์อย่างสมบูรณ์, "
        "(2) ช่องวัดอุณหภูมิชิป REF5045 (AIN2 = 0.5940 V คิดเป็น 32.20 °C ตามสมการ T_REF = (V_AIN2 - 0.509) / 0.00264) "
        "และวงจรวัดอุณหภูมิผิวหัววัด NTC (AIN3 = 2.0942 V → 28.48 °C และลดแรงดันลงเมื่อได้รับความร้อนจากโคมไฟ/ไฟฉายสู่ 28.80 °C) ทำงานได้อย่างแม่นยำ, "
        "และ (3) เมื่อส่องไฟฉายความเข้มสูงจ่อชิดหน้าเซนเซอร์ BPW 34 แรงดัน V_op พุ่งขึ้นไปแตะเพดาน 4.4998 V (เท่ากับ V_REF = 4.500 V) "
        "ยืนยันว่าการวัดแสงอาทิตย์กลางแจ้งผ่านกระจกแซฟไฟร์ใสจำเป็นต้องลดค่า R1 จาก 2,200 Ω ลงเป็น 1,100 Ω, 680 Ω หรือ 390 Ω (สำหรับพิกัดสูงสุด 4,000 W/m²) ตามที่คำนวณไว้ในตารางที่ 2.4"
    )

    add_section_heading(doc, "4.4", "ต้นแบบหน้าต่างแสดงผลบนเว็บ (Web Dashboard Mockup)")
    add_thai_body_p(
        doc,
        "ในส่วนซอฟต์แวร์และการแสดงผล คณะผู้จัดทำได้ออกแบบต้นแบบหน้าต่างแสดงผลบนเว็บ (Web Dashboard Mockup) "
        "เพื่อเตรียมพร้อมสำหรับเชื่อมต่อรับค่าดิบ (Raw Data) จาก ESP32-C6 และค่ากำลังไฟฟ้าจากอินเวอร์เตอร์ในโครงงาน 2 "
        "โดยครอบคลุมทั้งหน้าติดตามค่าความเข้มแสงและ Performance Ratio (PR) แบบเรียลไทม์, กราฟเทียบกำลังไฟฟ้าจริงกับค่าคาดหวัง, "
        "หน้าประวัติข้อมูล (Data Logs & CSV Export) และหน้าตั้งค่าพารามิเตอร์ระบบ (Settings & Calibration Constants) ดังแสดงในรูปที่ 4.5"
    )

    web_dir = os.path.join(img_dir, "web_mockup")
    w1 = os.path.join(web_dir, "slide15_pic01.jpg")
    w2 = os.path.join(web_dir, "slide16_pic02.jpg")
    w3 = os.path.join(web_dir, "slide17_pic01.jpg")
    if os.path.exists(w1):
        add_figure_image(
            doc,
            w1,
            fig_no="4.5",
            title="ต้นแบบหน้าต่างแสดงผล Web Dashboard Mockup (หน้าภาพรวมความเข้มรังสีอาทิตย์ ค่า PR และสถานะระบบ)",
            width_inches=5.6
        )
    if os.path.exists(w2) and os.path.exists(w3):
        add_two_figure_images_side_by_side(
            doc,
            w2, "(ก) หน้าวิเคราะห์ข้อมูลเปรียบเทียบกำลังไฟฟ้าและตาราง Log",
            w3, "(ข) หน้าตั้งค่าพารามิเตอร์ R_TIA, ค่าคงที่ชดเชย และการเชื่อมต่อ Modbus",
            fig_no="4.5 (ต่อ)",
            main_title="ต้นแบบหน้าต่าง Web Dashboard Mockup สำหรับการวิเคราะห์ Log และการตั้งค่าสอบเทียบ",
            col_width_inch=2.85
        )

    add_section_heading(doc, "4.5", "ขอบเขตงานทดสอบที่อยู่ระหว่างดำเนินการในโครงงาน 2")
    add_thai_body_p(
        doc,
        "สำหรับผลลัพธ์ในส่วนที่ยังไม่ได้ดำเนินการทดสอบจริง ณ สิ้นสุดภาคการศึกษาของโครงงานปริญญานิพนธ์ 1 "
        "(ได้แก่ ผลทดสอบทางไฟฟ้าหลังบัดกรีลงแผ่นวงจรพิมพ์ 4 เลเยอร์จริง V1.1 L1, ผลการวัดค่าการส่องผ่านสเปกตรัมของกระจกแซฟไฟร์ด้วย Vernier LabQuest 2, "
        "กราฟ Cosine Error จริงที่มุม 0°–85° จากห้องปฏิบัติการฟิสิกส์ และผลสอบเทียบกลางแจ้งเทียบกับไพราโนมิเตอร์มาตรฐาน) "
        "อยู่ระหว่างการเตรียมดำเนินการต่อเนื่องในโครงงานปริญญานิพนธ์ 2 (Capstone Project 2) จึงยังไม่บันทึกผลตัวเลขดังกล่าวในเล่มรายงานฉบับนี้"
    )

    # -------------------------------------------------------------------------
    # บทที่ 5: สรุปผล ความก้าวหน้า และแผนการดำเนินงานต่อยอด
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "5", "สรุปผล ความก้าวหน้า และแผนการดำเนินงานต่อยอด\n(Conclusion, Progress and Future Work)")

    add_section_heading(doc, "5.1", "สรุปสถานะความก้าวหน้าตามแผนงาน 6 ขั้นตอนหลัก")
    add_thai_body_p(
        doc,
        "สรุปสถานะความก้าวหน้าของการดำเนินงานโครงงานปริญญานิพนธ์ 1 ณ ปัจจุบัน แบ่งออกเป็น 6 ขั้นตอนหลักตามแผนการดำเนินงานตลอดปีการศึกษา (ตารางที่ 5.1):"
    )
    add_numbered_item(doc, "1.", "สร้างโมเดล 3D CAD ของเซนเซอร์ทั้ง BPX 61 และ BPW 34 บน FreeCAD และรันการจำลองแสง Ray-Tracing ด้วย Mitsuba 3 พิสูจน์จุดตัดแสงที่มุม 44° สำเร็จ (สำเร็จแล้ว 100%)", bold_prefix="ขั้นตอนที่ 1 — คัดเลือกหัววัดและจำลองแสง 3 มิติ: ")
    add_numbered_item(doc, "2.", "วิเคราะห์การเสื่อมสภาพของเรซิน Epoxy และขาโลหะเงิน แล้วออกแบบกล่องเคส 3D CAD ปิดผนึกพร้อมแหวนล็อกกระจกหน้าต่างแซฟไฟร์ (Al₂O₃ ความแข็ง 9 Mohs, ส่องผ่าน ~85% ช่วง 300–1100 nm) เรียบร้อยแล้ว (สำเร็จแล้ว 100%)", bold_prefix="ขั้นตอนที่ 2 — ออกแบบเคสปิดผนึกและกระจกแซฟไฟร์: ")
    add_numbered_item(doc, "3.", "ออกแบบวงจรขยายสัญญาณ OPA380 TIA (V_bias คำนวณ 0.409 V / วัดจริง 0.418 V), 24-bit ADC (ADS1220), เซนเซอร์อุณหภูมิ NTC (B25/85 = 3975 K) และภาคจ่ายไฟสัญญาณรบกวนต่ำ (TPS7A4901 + REF5045) ครบถ้วน (สำเร็จแล้ว 100%)", bold_prefix="ขั้นตอนที่ 3 — ออกแบบวงจร TIA 24 บิตและภาคจ่ายไฟ: ")
    add_numbered_item(doc, "4.", "บัดกรีวงจรภาคอ่านสัญญาณลงบนบอร์ด Perfboard (70×90 mm) ทดสอบอ่านค่า ADC 24 บิตผ่าน ESP32-C6 และทดสอบวัดแรงดันภาคจ่ายไฟจริงครบทุกตัว (TPS7A4901 บน PCB รุ่นแรก = 5.194 V, REF5045 บน Perfboard = 4.495 V) พร้อมออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนานรุ่น 'V1.1 L1 by Tachin' ผ่านการตรวจ DRC 100% โดยยังเหลือการสั่งผลิตและบัดกรีทดสอบจริงบนแผ่น PCB 4 เลเยอร์ในโครงงาน 2 (สำเร็จแล้ว 65%)", bold_prefix="ขั้นตอนที่ 4 — สร้างบอร์ดต้นแบบ Perfboard และออกแบบ 4-Layer PCB: ")
    add_numbered_item(doc, "5.", "นำหัววัดเข้าทดสอบวัดสเปกตรัมด้วย Vernier LabQuest 2 และเก็บกราฟ Cosine Error จริงที่มุมตกกระทบ 0°–85° บนแท่นหมุน Optical Rotation Stage ในห้องปฏิบัติการฟิสิกส์ (แผนงานโครงงาน 2 ระยะที่ 1)", bold_prefix="ขั้นตอนที่ 5 — เก็บกราฟ Cosine Error จริงในแล็บฟิสิกส์: ")
    add_numbered_item(doc, "6.", "พัฒนา Data Logger รับค่า Raw ทุกช่องผ่าน Wi-Fi (MQTT) / RS-485 (Modbus) เพื่อรันสมการชดเชยจากผลแล็บฟิสิกส์ร่วมกับพิกัดเวลาและอุณหภูมิ แสดงผลบน Web Dashboard และสอบเทียบกลางแจ้งเทียบกับเครื่องมาตรฐาน (แผนงานโครงงาน 2 ระยะที่ 2)", bold_prefix="ขั้นตอนที่ 6 — พัฒนา Data Logger, สมการชดเชย และ Web Dashboard: ")

    add_table_caption(doc, "5.1", "สรุปสถานะความก้าวหน้าการดำเนินงานโครงงาน (Roadmap 6 ขั้นตอนหลัก)")
    t_road = doc.add_table(rows=7, cols=3)
    t_road.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_road, border_color="000000", border_sz="4")
    road_headers = ["ขั้นตอนหลัก", "รายละเอียดผลสัมฤทธิ์และการดำเนินงาน", "สถานะปัจจุบัน"]
    for j, h in enumerate(road_headers):
        cell = t_road.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    road_data = [
        ["1. จำลองแสง 3 มิติ", "สร้าง 3D CAD (BPX 61 / BPW 34) และรัน Mitsuba 3 พิสูจน์จุดตัดแสง 44°", "สำเร็จแล้ว 100%"],
        ["2. เคสปิดผนึก & แซฟไฟร์", "ออกแบบเคส 3D CAD และแหวนล็อกกระจกแซฟไฟร์ (9 Mohs, T ≈ 85%)", "สำเร็จแล้ว 100%"],
        ["3. วงจร TIA & ภาคจ่ายไฟ", "ออกแบบ OPA380 TIA (V_bias ≈ 0.418 V), ADS1220, NTC, TPS7A4901, REF5045", "สำเร็จแล้ว 100%"],
        ["4. ต้นแบบ & 4-Layer PCB", "เทส Perfboard + วัดไฟจริงครบทุกตัว, ออกแบบ 4-Layer PCB ผ่าน DRC (รอบัดกรีลง PCB จริง)", "สำเร็จแล้ว 65%"],
        ["5. ทดสอบแล็บฟิสิกส์", "วัดสเปกตรัม (LabQuest 2) และเก็บกราฟ Cosine Error จริง 0°–85° บนแท่นหมุน", "โครงงาน 2 (ระยะที่ 1)"],
        ["6. Data Logger & Web", "ส่งค่า Raw (Wi-Fi / RS-485), รันสมการชดเชย & สอบเทียบกลางแจ้ง", "โครงงาน 2 (ระยะที่ 2)"]
    ]
    for i, row in enumerate(road_data):
        for j, val in enumerate(row):
            cell = t_road.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j == 2 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=13, bold=(j == 2))

    add_section_heading(doc, "5.2", "อภิปรายผลภาพรวมและแนวทางการดำเนินงานสู่โครงงาน 2")
    add_thai_body_p(
        doc,
        "จากการดำเนินงานโครงงานปริญญานิพนธ์ 1 สามารถสรุปผลสัมฤทธิ์หลักได้ 2 ด้าน ได้แก่: "
        "(1) ด้านฮาร์ดแวร์และทัศนศาสตร์ (ขั้นตอนที่ 1–4): พิสูจน์ได้ทั้งในแบบจำลอง Mitsuba 3 และต้นแบบฮาร์ดแวร์จริงว่า "
        "การเลือกใช้โฟโตไดโอดตัวถังแบนราบ BPW 34 ร่วมกับเคสปิดผนึกและกระจกแซฟไฟร์ช่วยขจัดปัญหาขอบกระป๋องบังแสงที่มุม 44° ของ BPX 61 "
        "และวงจร OPA380 TIA (V_bias = 0.418 V) คู่กับ ADS1220 24 บิต และภาคจ่ายไฟ TPS7A4901 (5.194 V) / REF5045 (4.495 V) สามารถอ่านสัญญาณแสงและอุณหภูมิได้เสถียรครบทุกช่อง "
        "และ (2) ด้านการชดเชยและระบบข้อมูล (ขั้นตอนที่ 5–6): ในโครงงานปริญญานิพนธ์ 2 จะดำเนินการสั่งผลิตและบัดกรีลงแผ่นวงจรพิมพ์ 4 เลเยอร์จริง "
        "ทดสอบเก็บกราฟ Cosine Error จริงในห้องปฏิบัติการฟิสิกส์ เพื่อนำไปสร้างสมการชดเชยบนตัว Data Logger ร่วมกับพิกัดเวลาและมุมดวงอาทิตย์ของประเทศไทย (มุมเอียงเฉลี่ยประมาณ 12° อ้อมทางทิศใต้) รวมถึงอุณหภูมิและค่ากำลังไฟฟ้าจากอินเวอร์เตอร์ "
        "เพื่อแสดงผลเปรียบเทียบค่าก่อน-หลังชดเชยบน Web Dashboard สอบเทียบกลางแจ้งเทียบกับไพราโนมิเตอร์มาตรฐาน และประยุกต์ใช้เป็นชุดเครื่องมือตรวจรับคุณภาพระบบโซลาร์เซลล์ตามบ้านพักอาศัย (งบติดตั้ง 50,000–150,000 บาท) ให้สมบูรณ์พร้อมใช้งานจริงต่อไป"
    )

    # -------------------------------------------------------------------------
    # เอกสารอ้างอิง (References)
    # -------------------------------------------------------------------------
    doc.add_page_break()
    p_ref_head = doc.add_paragraph()
    p_ref_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ref_head.paragraph_format.space_before = Pt(10)
    p_ref_head.paragraph_format.space_after = Pt(16)
    r_ref = p_ref_head.add_run("เอกสารอ้างอิง\n(References)")
    set_font_sarabun(r_ref, size_pt=18, bold=True)
    set_outline_level(p_ref_head, 0)

    references = [
        "International Organization for Standardization. (2018). Solar energy — Specification and classification of instruments for measuring hemispherical solar and direct solar radiation (ISO Standard No. 9060:2018).",
        "Graeme, J. G. (1996). Photodiode Amplifiers: OP AMP Solutions. McGraw-Hill Education.",
        "Hamamatsu Photonics K.K. (2022). Opto-Semiconductor Handbook — Chapter 2: Si Photodiodes (Characteristics, Equivalent Circuit, and Upper Limit of Linearity).",
        "ams-OSRAM AG. (2016). BPW 34 DIL Silicon PIN Photodiode [Datasheet]. Version 1.4 / 1.6, Ordering Code Q62702P0073.",
        "ams-OSRAM AG. (2023). BPX 61 Silicon PIN Photodiode in Metal Can TO-39 Package [Datasheet]. Version 1.6, Document No. Q62705P0025.",
        "Texas Instruments. (2007). OPA380, OPA2380 Precision, High-Speed Transimpedance Operational Amplifier [Datasheet]. SBOS291G.",
        "Texas Instruments. (2016). ADS1220 4-Channel, 2-kSPS, Low-Power, 24-Bit ADC with Integrated PGA and Reference [Datasheet]. SBAS501B.",
        "Texas Instruments. (2018). REF50xx Low-Noise, Very Low Drift, Precision Voltage Reference [Datasheet]. SBOS410I.",
        "Texas Instruments. (2015). TPS7A49 36-V, 150-mA, Ultra-Low-Noise, Positive Linear Regulator [Datasheet]. SBVS098B.",
        "Thinking Electronic Industrial Co., Ltd. (2022). TTC03 Series Epoxy Coated Type NTC Thermistor (TTC3A103F39HAEY) Specifications [Datasheet].",
        "Hecht, E. (2017). Optics (5th ed.). Pearson Education. (Chapters on Fresnel Reflection and Snell's Law).",
        "World Meteorological Organization (WMO). (2021). Guide to Instruments and Methods of Observation (WMO-No. 8), Volume I — Measurement of Radiation."
    ]
    for idx, ref_text in enumerate(references, start=1):
        add_numbered_item(
            doc,
            num_str=f"[{idx}]",
            text=ref_text,
            space_after=4,
            left_indent=0.55,
            hanging_indent=0.35
        )

    # -------------------------------------------------------------------------
    # ภาคผนวก (Appendices)
    # -------------------------------------------------------------------------
    doc.add_page_break()
    p_appx_head = doc.add_paragraph()
    p_appx_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_appx_head.paragraph_format.space_before = Pt(10)
    p_appx_head.paragraph_format.space_after = Pt(16)
    r_appx = p_appx_head.add_run("ภาคผนวก\n(Appendices)")
    set_font_sarabun(r_appx, size_pt=18, bold=True)
    set_outline_level(p_appx_head, 0)

    add_section_heading(doc, "ภาคผนวก ก", "รายการอุปกรณ์ต้นแบบและตารางค่าปรับตั้ง (Bill of Materials - BOM)")
    add_thai_body_p(
        doc,
        "รายการอุปกรณ์และชิ้นส่วนอิเล็กทรอนิกส์ทั้งหมดที่ใช้ในวงจรต้นแบบบน Perfboard และแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom PCB) "
        "พร้อมตัวเลือกค่าความต้านทานป้อนกลับ R1 (สำหรับพิกัดสูงสุดถึง 4,000 W/m²) และตัวต้านทานกำหนดแรงดันไบแอส R4:"
    )

    t_bom = doc.add_table(rows=12, cols=4)
    t_bom.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_bom, border_color="000000", border_sz="4")
    bom_headers = ["อุปกรณ์ / ไอซี", "เบอร์อุปกรณ์ / ค่าพารามิเตอร์", "คุณสมบัติสำคัญทางวิศวกรรม", "สถานะการใช้งาน"]
    for j, h in enumerate(bom_headers):
        cell = t_bom.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    bom_data = [
        ["Silicon Photodiode", "ams-OSRAM BPW 34", "Flat DIL Package, 7.02 mm², P_tot=150 mW, R_S ≈ 4.5 Ω", "ใช้งานหลักบน Perfboard / PCB"],
        ["Alternative Detector", "ams-OSRAM BPX 61", "TO-39 Metal Can, ชิป 7.02 mm² เดียวกัน", "สำหรับเปรียบเทียบผลเชิงมุม"],
        ["Optical Window", "Synthetic Sapphire (Al₂O₃)", "ความแข็ง 9 Mohs, หนา 1 mm, T ≈ 85% (ไม่มี Diffuser)", "พร้อมใช้งานครอบเซนเซอร์"],
        ["Transimpedance Op-Amp", "TI OPA380AIDR", "High-Speed Precision TIA, GBW 90 MHz, I_SC 150 mA", "ใช้งานจริงในภาคขยาย TIA"],
        ["TIA Feedback Network", "R1: 2.2k / 1.1k / 680 / 390 Ω\nC1: 10 pF (C0G/NP0)", "ปรับพิกัดได้ถึง 4,000 W/m² (Class A) พร้อม Butterworth Cap", "ติดตั้งปรับค่าได้บนบอร์ด"],
        ["Bias Divider Network", "R3: 100 kΩ, C3: 1 µF\nR4: 10 kΩ (หรือ 4.7k / 3.3k)", "สร้าง V_bias = 0.409 V (วัดจริง 0.418 V; หรือลดเป็น 0.202 V / 0.144 V)", "ติดตั้งปรับค่าได้บนบอร์ด"],
        ["24-bit Delta-Sigma ADC", "TI ADS1220IPWR", "TSSOP-16, 4-ch MUX, 20 SPS (50/60 Hz Rejection)", "ใช้งานจริงบน Perfboard / PCB"],
        ["Precision Voltage Ref.", "TI REF5045AIDR", "SOIC-8, V_REF = 4.500 V, 3 ppm/°C, ขา 3 วัดอุณหภูมิชิป", "ใช้งานจริงบน Perfboard / PCB"],
        ["Ultra-Low-Noise LDO", "TI TPS7A4901DGNR", "MSOP-8 PowerPAD (ต่อขา 9 EP ลง GND), จ่ายไฟ +5.20 V", "ใช้งานจริงบน Perfboard / PCB"],
        ["NTC Thermistor", "Thinking TTC3A103F39HAEY", "10 kΩ ±1% ที่ 25°C, B25/85 = 3975 K ±1.5% (รหัส 39H)", "ใช้งานจริงแนบชิดเซนเซอร์"],
        ["Microcontroller Unit", "ESP32-C6 Super Mini", "RISC-V 32-bit, Wi-Fi 6, ประมวลผลสมการชดเชย 4 ขั้นตอน", "ใช้งานจริงพร้อม Web Dashboard"]
    ]
    for i, row in enumerate(bom_data):
        for j, val in enumerate(row):
            cell = t_bom.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0 or j == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=13)

    add_section_heading(doc, "ภาคผนวก ข", "เอกสารมาตรฐานสากลอ้างอิง ISO 9060:2018 (Table 1)")
    add_thai_body_p(
        doc,
        "ภาพถ่ายตารางเกณฑ์การจำแนกประเภทไพราโนมิเตอร์ (Table 1 — Pyranometer classification list) "
        "จากเอกสารมาตรฐานสากล ISO 9060:2018 ฉบับจริง ซึ่งใช้อ้างอิงพารามิเตอร์ Class A, Class B และ Class C ในตารางที่ 2.1:"
    )
    iso_t1_p1 = os.path.join(img_dir, "iso9060_crops", "iso_p12_table1_part1.png")
    iso_t1_p2 = os.path.join(img_dir, "iso9060_crops", "iso_p13_table1_part2.png")
    if os.path.exists(iso_t1_p1) and os.path.exists(iso_t1_p2):
        add_two_figure_images_stacked(
            doc,
            iso_t1_p1, "(ก) ตารางที่ 1 ส่วนที่ 1: Response time (a), Zero off-set (b), Non-stability (c1), Non-linearity (c2) และ Directional response (c3)",
            iso_t1_p2, "(ข) ตารางที่ 1 ส่วนที่ 2: Spectral error (c4), Temperature response (c5), Tilt response (c6) และ Additional errors (c7)",
            fig_no="ข.1",
            main_title="ตารางเกณฑ์การจำแนกประเภท Pyranometer จากเอกสารมาตรฐานสากล ISO 9060:2018 (Table 1)",
            width_inches=5.4
        )

    # บันทึกไฟล์เอกสารฉบับสมบูรณ์ (พร้อม Fallback หากไฟล์เปิดค้างอยู่ใน Word)
    out_dir = os.path.dirname(__file__)
    candidates = [
        "69_IoT_Project1_Report_V2.docx",
        "69_IoT_Project1_Report_Progress.docx",
        "69_IoT_Project1_Report_Latest.docx"
    ]
    saved_path = None
    for fname in candidates:
        candidate_path = os.path.join(out_dir, fname)
        try:
            doc.save(candidate_path)
            print(f"[OK] Report Document successfully generated at: {candidate_path}")
            saved_path = candidate_path
            break
        except PermissionError:
            print(f"[NOTE] '{fname}' is currently open/locked in Word. Trying next candidate...")

    return saved_path

if __name__ == "__main__":
    create_report_v2()
