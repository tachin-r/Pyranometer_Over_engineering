"""
========================================================================================
สคริปต์คอมไพล์เล่มรายงานความก้าวหน้าโครงงานปริญญานิพนธ์ 1 (Capstone Project 1 Progress Report)
โครงงาน: การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (High-Precision Pyranometer)
หลักสูตรสองปริญญา: วศ.บ. (วิศวกรรมระบบไอโอทีและสารสนเทศ) และ วท.บ. (ฟิสิกส์อุตสาหกรรม)
คณะวิศวกรรมศาสตร์ สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง (สจล.)

ขอบเขตรายงาน: 3 บทหลัก (บทที่ 1 บทนำ, บทที่ 2 ทฤษฎีและวรรณกรรม, บทที่ 3 การออกแบบ การสร้างต้นแบบ การจำลอง และระเบียบวิธีวิจัย)
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
    set_section_numbering,
    add_thesis_chapter_heading,
    add_section_heading,
    add_subsection_heading,
    add_thai_body_p,
    add_numbered_item,
    add_bullet_item,
    add_toc_line,
    set_table_clean_borders,
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
    # SECTION 1: ปกนอก ปกใน และหน้ารับรองโครงงาน (ไม่มีเลขหน้า)
    # =========================================================================
    sec1 = doc.sections[0]
    setup_kmitl_page(sec1)

    # 1.1 หน้าปกภาษาไทย
    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(36)

    p_th_title = doc.add_paragraph()
    p_th_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_th_title.paragraph_format.space_after = Pt(8)
    p_th_title.paragraph_format.line_spacing = 1.15
    r_th = p_th_title.add_run(
        "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"
    )
    set_font_sarabun(r_th, size_pt=20, bold=True)

    p_en_title = doc.add_paragraph()
    p_en_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_title.paragraph_format.space_after = Pt(60)
    p_en_title.paragraph_format.line_spacing = 1.15
    r_en = p_en_title.add_run(
        "DEVELOPMENT OF SOLAR IRRADIANCE MEASUREMENT SYSTEM"
    )
    set_font_sarabun(r_en, size_pt=18, bold=True)

    p_authors = doc.add_paragraph()
    p_authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_authors.paragraph_format.space_after = Pt(70)
    p_authors.paragraph_format.line_spacing = 1.2
    r_auth = p_authors.add_run(
        "นายเตชินท์  รัฐวงษา     รหัสนักศึกษา 66010296\n"
        "นายณัฐพัชร์ รัตนนวงศ์    รหัสนักศึกษา 66010256\n"
        "นายพีรวิช  เฉลิมเกียรติ  รหัสนักศึกษา 66010591"
    )
    set_font_sarabun(r_auth, size_pt=16, bold=True)

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_footer.paragraph_format.line_spacing = 1.15
    r_foot = p_footer.add_run(
        "รายงานความก้าวหน้าโครงงานปริญญานิพนธ์ 1 นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตร\n"
        "หลักสูตรสองปริญญา (Dual Degree):\n"
        "วิศวกรรมศาสตรบัณฑิต (วิศวกรรมระบบไอโอทีและสารสนเทศ)\n"
        "และ วิทยาศาสตรบัณฑิต (ฟิสิกส์อุตสาหกรรม)\n"
        "คณะวิศวกรรมศาสตร์\n"
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง\n"
        "ปีการศึกษา 2569"
    )
    set_font_sarabun(r_foot, size_pt=16, bold=True)

    doc.add_page_break()

    # 1.2 หน้าปกภาษาอังกฤษ
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(40)

    p_en_cov = doc.add_paragraph()
    p_en_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_cov.paragraph_format.space_after = Pt(70)
    p_en_cov.paragraph_format.line_spacing = 1.15
    r_encov = p_en_cov.add_run(
        "DEVELOPMENT OF SOLAR IRRADIANCE\n"
        "MEASUREMENT SYSTEM"
    )
    set_font_sarabun(r_encov, size_pt=20, bold=True)

    p_auth_en = doc.add_paragraph()
    p_auth_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth_en.paragraph_format.space_after = Pt(70)
    p_auth_en.paragraph_format.line_spacing = 1.2
    r_authen = p_auth_en.add_run(
        "MR. TECHIN RATTAWONGSA          STUDENT ID 66010296\n"
        "MR. NATTAPAT RATANANWONG        STUDENT ID 66010256\n"
        "MR. PEERAWICH CHALOEMKIAT       STUDENT ID 66010591"
    )
    set_font_sarabun(r_authen, size_pt=16, bold=True)

    p_en_sub = doc.add_paragraph()
    p_en_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_sub.paragraph_format.line_spacing = 1.15
    r_ensub = p_en_sub.add_run(
        "A CAPSTONE PROJECT 1 PROGRESS REPORT SUBMITTED IN PARTIAL FULFILLMENT\n"
        "OF THE REQUIREMENTS FOR THE DUAL DEGREE OF\n"
        "BACHELOR OF ENGINEERING IN IOT SYSTEM AND INFORMATION ENGINEERING AND\n"
        "BACHELOR OF SCIENCE IN INDUSTRIAL PHYSICS\n"
        "SCHOOL OF ENGINEERING\n"
        "KING MONGKUT'S INSTITUTE OF TECHNOLOGY LADKRABANG\n"
        "ACADEMIC YEAR 2026"
    )
    set_font_sarabun(r_ensub, size_pt=15, bold=True)

    doc.add_page_break()

    # 1.3 หน้ารับรองโครงงาน (Approval Sheet)
    p_app_head = doc.add_paragraph()
    p_app_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_app_head.paragraph_format.space_before = Pt(10)
    p_app_head.paragraph_format.space_after = Pt(14)
    r_app_h = p_app_head.add_run("ใบรับรองรายงานความก้าวหน้าโครงงานปริญญานิพนธ์ 1")
    set_font_sarabun(r_app_h, size_pt=18, bold=True)

    t_app_info = doc.add_table(rows=7, cols=2)
    t_app_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_app_info, border_color="000000", border_sz="4")
    col_widths = [Inches(1.8), Inches(4.5)]
    for row in t_app_info.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    app_data = [
        ("หัวข้อปริญญานิพนธ์", "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"),
        ("Thesis Title", "Development of Solar Irradiance Measurement System"),
        ("ชื่อนักศึกษา", "นาย เตชินท์  รัฐวงษา    รหัสนักศึกษา 66010296\nนาย ณัฐพัชร์ รัตนนวงศ์   รหัสนักศึกษา 66010256\nนาย พีรวิช  เฉลิมเกียรติ รหัสนักศึกษา 66010591"),
        ("ระดับปริญญา", "วิศวกรรมศาสตรบัณฑิตและวิทยาศาสตรบัณฑิต (หลักสูตรสองปริญญา)"),
        ("สาขาวิชา", "วิศวกรรมระบบไอโอทีและสารสนเทศ และ ฟิสิกส์อุตสาหกรรม"),
        ("ปีการศึกษา", "2569"),
        ("อาจารย์ที่ปรึกษา", "ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ\nผศ.ดร.ประธาน บูรณศิริ")
    ]

    for idx, (label, val) in enumerate(app_data):
        c0 = t_app_info.cell(idx, 0)
        c1 = t_app_info.cell(idx, 1)
        set_cell_margins(c0, top=50, bottom=50, left=80, right=80)
        set_cell_margins(c1, top=50, bottom=50, left=80, right=80)
        r0 = c0.paragraphs[0].add_run(label)
        set_font_sarabun(r0, size_pt=15, bold=True)
        r1 = c1.paragraphs[0].add_run(val)
        set_font_sarabun(r1, size_pt=15)

    p_app_body = doc.add_paragraph()
    p_app_body.paragraph_format.space_before = Pt(14)
    p_app_body.paragraph_format.space_after = Pt(12)
    p_app_body.paragraph_format.line_spacing = 1.15
    r_app_txt = p_app_body.add_run(
        "คณะวิศวกรรมศาสตร์ สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง "
        "อนุมัติให้รายงานความก้าวหน้านี้เป็นส่วนหนึ่งของรายวิชา 01236269 - IoT AND INFORMATION CAPSTONE PROJECT 1 "
        "ตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต (วิศวกรรมระบบไอโอทีและสารสนเทศ) และวิทยาศาสตรบัณฑิต (ฟิสิกส์อุตสาหกรรม) ประจำปีการศึกษา 2569"
    )
    set_font_sarabun(r_app_txt, size_pt=15)

    # Committee Table
    t_comm = doc.add_table(rows=6, cols=2)
    t_comm.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_comm, border_color="000000", border_sz="4")
    for r in t_comm.rows:
        r.cells[0].width = Inches(3.5)
        r.cells[1].width = Inches(2.8)

    c_th0 = t_comm.cell(0, 0)
    c_th1 = t_comm.cell(0, 1)
    set_cell_shading(c_th0, "F1F5F9")
    set_cell_shading(c_th1, "F1F5F9")
    set_cell_margins(c_th0, top=60, bottom=60, left=80, right=80)
    set_cell_margins(c_th1, top=60, bottom=60, left=80, right=80)
    r_th0 = c_th0.paragraphs[0].add_run("คณะกรรมการสอบ")
    set_font_sarabun(r_th0, size_pt=15, bold=True)
    r_th1 = c_th1.paragraphs[0].add_run("ลายมือชื่อ")
    set_font_sarabun(r_th1, size_pt=15, bold=True)

    comm_members = [
        "ประธานกรรมการสอบ",
        "กรรมการ",
        "กรรมการ",
        "กรรมการและอาจารย์ที่ปรึกษา (ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ)",
        "กรรมการและอาจารย์ที่ปรึกษาร่วม (ผศ.ดร.ประธาน บูรณศิริ)"
    ]

    for idx, role in enumerate(comm_members, start=1):
        cell_role = t_comm.cell(idx, 0)
        cell_sig = t_comm.cell(idx, 1)
        set_cell_margins(cell_role, top=60, bottom=60, left=80, right=80)
        set_cell_margins(cell_sig, top=60, bottom=60, left=80, right=80)
        r_role = cell_role.paragraphs[0].add_run(role)
        set_font_sarabun(r_role, size_pt=15)
        r_sig = cell_sig.paragraphs[0].add_run("....................................................")
        set_font_sarabun(r_sig, size_pt=15)

    # =========================================================================
    # SECTION 2: ส่วนนำ (บทคัดย่อไทย-อังกฤษ, กิตติกรรมประกาศ, สารบัญ) -> เลขหน้าโรมัน i, ii, iii...
    # =========================================================================
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec2)
    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False
    set_section_numbering(sec2, fmt="lowerRoman", start=1)
    add_page_number_to_footer(sec2.footer)

    # 2.1 บทคัดย่อภาษาไทย (หน้า i)
    p_abs_title = doc.add_paragraph()
    p_abs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs_title.paragraph_format.space_before = Pt(10)
    p_abs_title.paragraph_format.space_after = Pt(12)
    r_abst = p_abs_title.add_run("บทคัดย่อ")
    set_font_sarabun(r_abst, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "รายงานความก้าวหน้าโครงงานปริญญานิพนธ์ 1 ฉบับนี้นำเสนอการวิจัย ออกแบบ และสร้างต้นแบบระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง "
        "(High-Precision Solar Irradiance Measurement System) โดยมุ่งเน้นการแก้ปัญหาความคลาดเคลื่อนทางทัศนศาสตร์เชิงมุม (Cosine Error) "
        "และเสถียรภาพของวงจรตรวจวัดสัญญาณแอนะล็อกความละเอียดสูง ในขั้นแรกได้ทำการวิเคราะห์เปรียบเทียบเซนเซอร์โฟโตไดโอดชนิดซิลิคอนระหว่างรุ่น BPX 61 "
        "(ตัวถังโลหะ TO-39) และรุ่น BPW 34 (ตัวถังพลาสติกแบนราบ) ซึ่งพบว่าทั้งสองรุ่นใช้ชิปซิลิคอน (Silicon Die) ขนาดเดียวกันคือ 2.65 x 2.65 mm² "
        "(พื้นที่รับแสง 7.02 mm²) และมีคุณสมบัติทางไฟฟ้าเหมือนกันทุกประการ แต่ผลจากการจำลองทางทัศนศาสตร์ (Optical Ray-tracing Simulation) "
        "พบข้อเท็จจริงสำคัญว่า เมื่อจำลองเฉพาะกระป๋องเปล่าของ BPX 61 ขอบกระป๋องจะตัดแสงที่มุม 62 องศาตรงตามสเปก แต่เมื่อใส่กระจกหน้าต่างกระจก Schott 8250 "
        "การหักเหของแสงทำให้ขอบกระป๋องบังแสงเร็วขึ้นอย่างรุนแรงจนเกิดจุดตัดแสงตั้งแต่ 44 องศา โครงงานจึงเลือกใช้เซนเซอร์ BPW 34 ที่มีระนาบรับแสงแบนราบ "
        "ร่วมกับการติดตั้งกระจกแซฟไฟร์ (Sapphire Optical Window) เพื่อป้องกันรังสี UV และทนทานต่อรอยขีดข่วนกลางแจ้ง"
    )
    add_thai_body_p(
        doc,
        "ในภาคฮาร์ดแวร์ ได้ทำการต่อวงจรทดสอบขั้นแรกบน Breadboard ร่วมกับชิปแปลงสัญญาณ 24-bit Delta-Sigma ADC (ADS1220), "
        "แหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูง (REF5045) และหัววัดอุณหภูมิ NTC Thermistor ซึ่งสามารถอ่านค่าสัญญาณจริงจากสภาวะแสงต่าง ๆ "
        "(ปิดไฟ, แสงขาวสว่างสุด, แสงส้มสว่างสุด, และแสงส้ม+ขาว) ผ่านพอร์ต Serial ได้สำเร็จ แต่ตรวจพบประเด็นเสถียรภาพของแรงดันอ้างอิง REF5045 "
        "อันเนื่องมาจากตัวเก็บประจุโหลดที่ใช้อยู่มีค่าความต้านทานอนุกรมสมมูล (ESR) สูงเกินเกณฑ์มาตรฐาน คณะผู้จัดทำจึงตัดสินใจรอดูตัวเก็บประจุชนิด Low-ESR Ceramic "
        "ที่สั่งซื้อไป (กำหนดส่งมอบ 30 ต.ค. 2569) ก่อนทำการบัดกรีลงแผงวงจรอเนกประสงค์ (Perfboard) ในคราวเดียว นอกจากนี้ ผลซิมที่พบว่าการมีกระจกทำให้ค่าเพี้ยนอย่างมีนัยสำคัญ "
        "ทำให้คณะผู้จัดทำได้ออกแบบระเบียบวิธีการทดลองจริงร่วมกับภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ ได้แก่ การทดสอบการส่องผ่านสเปกตรัม (UV-Vis-NIR Spectrophotometer) "
        "และการทดสอบการตอบสนองเชิงมุมบนแท่นหมุนโกนิโอมิเตอร์ เพื่อนำไปปฏิบัติการจริงในการสอบเทียบระยะถัดไป"
    )

    doc.add_page_break()

    # 2.2 บทคัดย่อภาษาอังกฤษ (หน้า ii)
    p_absen_title = doc.add_paragraph()
    p_absen_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_absen_title.paragraph_format.space_before = Pt(10)
    p_absen_title.paragraph_format.space_after = Pt(12)
    r_absen = p_absen_title.add_run("ABSTRACT")
    set_font_sarabun(r_absen, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "This Capstone Project 1 progress report presents the research, design, and prototyping of a high-precision solar irradiance measurement system, "
        "focusing on mitigating directional cosine error and front-end instrumentation noise. An in-depth comparative analysis between the BPX 61 (TO-39 metal can) "
        "and BPW 34 (flat clear plastic package) silicon photodiodes revealed that both devices share the exact same silicon die (2.65 x 2.65 mm², 7.02 mm² active area) "
        "with identical optoelectronic characteristics. However, optical ray-tracing simulations revealed that while the bare TO-39 can exhibits an optical cutoff at ~62°, "
        "introducing the Schott 8250 glass window severely accelerates rim shadowing due to refractive bending, causing a premature optical cutoff at 44°. "
        "Consequently, the flat BPW 34 package was selected and coupled with an optical-grade sapphire window for enhanced outdoor durability and UV protection."
    )
    add_thai_body_p(
        doc,
        "The signal acquisition circuit was prototyped and validated on a breadboard, integrating a 24-bit delta-sigma ADC (TI ADS1220), "
        "a precision voltage reference (REF5045), and an NTC thermistor. Successful live data logging was demonstrated across various lighting conditions "
        "(dark, full white LED, full orange LED, and combined spectrum). Debugging identified reference voltage instability caused by excessive Equivalent Series Resistance (ESR) "
        "in the output decoupling capacitor, leading to the strategic decision to await low-ESR ceramic capacitors before executing the permanent perfboard assembly. "
        "Furthermore, because simulations demonstrated that optical windows introduce non-ideal refractive and reflective shifts, two rigorous experimental protocols "
        "were formulated in collaboration with the Faculty of Science: (1) UV-Vis-NIR spectrophotometer transmission testing, and (2) goniometric angular response verification, "
        "establishing a robust empirical foundation for Capstone Project 2."
    )

    doc.add_page_break()

    # 2.3 กิตติกรรมประกาศ (หน้า iii)
    p_ack_title = doc.add_paragraph()
    p_ack_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ack_title.paragraph_format.space_before = Pt(10)
    p_ack_title.paragraph_format.space_after = Pt(12)
    r_ack = p_ack_title.add_run("กิตติกรรมประกาศ")
    set_font_sarabun(r_ack, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "คณะผู้ดำเนินโครงงานขอกราบขอบพระคุณ ผศ.ดร.ธนวิชญ์ อนวงค์พินิจ อาจารย์ที่ปรึกษาหลัก และ ผศ.ดร.ประธาน บูรณศิริ "
        "อาจารย์ที่ปรึกษาร่วม ตลอดจนคณาจารย์ภาควิชาวิศวกรรมระบบไอโอทีและสารสนเทศ คณะวิศวกรรมศาสตร์ และภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ "
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง ที่ได้กรุณาถ่ายทอดองค์ความรู้ ให้คำแนะนำเชิงวิชาการอย่างลึกซึ้ง "
        "ทั้งในด้านการออกแบบวงจรอิเล็กทรอนิกส์ความแม่นยำสูงและการสนับสนุนเครื่องมือวัดทางทัศนศาสตร์ระดับห้องปฏิบัติการ"
    )
    add_thai_body_p(
        doc,
        "ขอขอบคุณคณาจารย์และเจ้าหน้าที่ห้องปฏิบัติการทัศนศาสตร์ ภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล. "
        "ที่ได้ให้ความอนุเคราะห์ในการเปิดให้ยืมใช้เครื่องตรวจวัดสเปกตรัม Vernier LabQuest 2 และอุปกรณ์ชุดแท่นหมุนวัดองศา "
        "ตลอดจนให้คำปรึกษาในการออกแบบระเบียบวิธีการทดลองเพื่อประเมินคุณลักษณะทางแสงของเซนเซอร์และกระจกแซฟไฟร์ในโครงงานนี้"
    )

    p_ack_sign = doc.add_paragraph()
    p_ack_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_ack_sign.paragraph_format.space_before = Pt(16)
    r_asign = p_ack_sign.add_run(
        "คณะผู้จัดทำโครงงาน\n"
        "นายเตชินท์  รัฐวงษา\n"
        "นายณัฐพัชร์ รัตนนวงศ์\n"
        "นายพีรวิช  เฉลิมเกียรติ"
    )
    set_font_sarabun(r_asign, size_pt=16, bold=True)

    doc.add_page_break()

    # 2.4 สารบัญ (Table of Contents - หน้า iv)
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_head.paragraph_format.space_before = Pt(10)
    p_toc_head.paragraph_format.space_after = Pt(12)
    r_toc = p_toc_head.add_run("สารบัญ")
    set_font_sarabun(r_toc, size_pt=18, bold=True)

    p_col = doc.add_paragraph()
    p_col.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), WD_TAB_ALIGNMENT.RIGHT)
    r_c1 = p_col.add_run("เรื่อง")
    set_font_sarabun(r_c1, size_pt=16, bold=True)
    r_c2 = p_col.add_run("\tหน้า")
    set_font_sarabun(r_c2, size_pt=16, bold=True)

    add_toc_line(doc, "บทคัดย่อภาษาไทย", "i")
    add_toc_line(doc, "บทคัดย่อภาษาอังกฤษ (ABSTRACT)", "ii")
    add_toc_line(doc, "กิตติกรรมประกาศ", "iii")
    add_toc_line(doc, "สารบัญ", "iv", is_bold=True)
    add_toc_line(doc, "สารบัญตาราง", "vi")
    add_toc_line(doc, "สารบัญรูป", "vii")
    add_toc_line(doc, "บทที่ 1 บทนำ", "1", is_bold=True)
    add_toc_line(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา", "1", level=1)
    add_toc_line(doc, "1.2 วัตถุประสงค์ของโครงงาน", "2", level=1)
    add_toc_line(doc, "1.3 ขอบเขตของโครงงาน (Project 1)", "3", level=1)
    add_toc_line(doc, "1.4 ประโยชน์ที่คาดว่าจะได้รับ", "3", level=1)
    add_toc_line(doc, "1.5 แผนการดำเนินงานโครงงาน", "4", level=1)
    add_toc_line(doc, "บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "5", is_bold=True)
    add_toc_line(doc, "2.1 มาตรฐานการวัดรังสีดวงอาทิตย์ (ISO 9060:2018)", "5", level=1)
    add_toc_line(doc, "2.2 การวิเคราะห์เปรียบเทียบชิปเซนเซอร์ BPX 61 และ BPW 34", "7", level=1)
    add_toc_line(doc, "2.3 ทฤษฎีโฟโตไดโอดและวงจรขยายกระแสแสง (Photodiode & TIA Theory)", "10", level=1)
    add_toc_line(doc, "2.4 คุณสมบัติทางแสงของกระจกแซฟไฟร์ (Sapphire Optical Window)", "13", level=1)
    add_toc_line(doc, "2.5 ไอซีแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)", "15", level=1)
    add_toc_line(doc, "2.6 แหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูง (REF5045) และข้อกำหนดตัวเก็บประจุ", "17", level=1)
    add_toc_line(doc, "บทที่ 3 การออกแบบระบบ การสร้างต้นแบบ การจำลอง และระเบียบวิธีวิจัย", "16", is_bold=True)
    add_toc_line(doc, "3.1 สถาปัตยกรรมระบบวัดและผังวงจรรวม", "16", level=1)
    add_toc_line(doc, "3.2 การสร้างต้นแบบและการทดสอบจริงบน Breadboard", "18", level=1)
    add_toc_line(doc, "3.3 ผลการจำลองทางทัศนศาสตร์และการค้นพบความผิดเพี้ยนจากกระจกหน้าต่าง", "22", level=1)
    add_toc_line(doc, "3.4 การออกแบบข้อเสนอระเบียบวิธีการทดลองจริงทางทัศนศาสตร์", "26", level=1)
    add_toc_line(doc, "3.5 สถานะปัจจุบันและการออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์สู่การผลิตจริง", "29", level=1)
    add_toc_line(doc, "เอกสารอ้างอิง", "31", is_bold=True)
    add_toc_line(doc, "ภาคผนวก", "33", is_bold=True)

    doc.add_page_break()

    # 2.5 สารบัญตาราง (List of Tables - หน้า vi)
    p_lot_head = doc.add_paragraph()
    p_lot_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lot_head.paragraph_format.space_before = Pt(10)
    p_lot_head.paragraph_format.space_after = Pt(12)
    r_lot = p_lot_head.add_run("สารบัญตาราง")
    set_font_sarabun(r_lot, size_pt=18, bold=True)

    p_col_t = doc.add_paragraph()
    p_col_t.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), WD_TAB_ALIGNMENT.RIGHT)
    r_ct1 = p_col_t.add_run("ตารางที่")
    set_font_sarabun(r_ct1, size_pt=16, bold=True)
    r_ct2 = p_col_t.add_run("\tหน้า")
    set_font_sarabun(r_ct2, size_pt=16, bold=True)

    add_toc_line(doc, "ตารางที่ 1.1 แผนการดำเนินงานและกิจกรรมโครงงานตลอดปีการศึกษา", "4")
    add_toc_line(doc, "ตารางที่ 2.1 เกณฑ์มาตรฐานการจำแนกประเภท Pyranometer ตาม ISO 9060:2018", "6")
    add_toc_line(doc, "ตารางที่ 2.2 การเปรียบเทียบพารามิเตอร์ทางเทคนิคระหว่าง OSRAM BPX 61 กับ BPW 34", "8")
    add_toc_line(doc, "ตารางที่ 3.1 ตารางกำหนดช่องสัญญาณอินพุตมัลติเพล็กเซอร์ (Input MUX) ของไอซี ADS1220", "17")
    add_toc_line(doc, "ตารางที่ 3.2 ตารางกำหนดการเชื่อมต่อขาสัญญาณระหว่างไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini กับ ADS1220", "18")
    add_toc_line(doc, "ตารางที่ 3.3 ข้อมูลการอ่านค่าแรงดันจริงบน Breadboard ภายใต้สภาวะแสงต่างๆ", "20")
    add_toc_line(doc, "ตารางที่ 3.4 ตารางเปรียบเทียบแรงดันอ้างอิง REF5045 ระหว่างแหล่งจ่าย USB และแหล่งจ่ายสวิตชิ่ง 5.2 V", "22")
    add_toc_line(doc, "ตารางที่ 3.5 ตารางเปรียบเทียบผลการจำลองเชิงมุม: BPX 61 vs BPW 34", "26")

    doc.add_page_break()

    # 2.6 สารบัญรูป (List of Figures - หน้า vii)
    p_lof_head = doc.add_paragraph()
    p_lof_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lof_head.paragraph_format.space_before = Pt(10)
    p_lof_head.paragraph_format.space_after = Pt(12)
    r_lof = p_lof_head.add_run("สารบัญรูป")
    set_font_sarabun(r_lof, size_pt=18, bold=True)

    p_col_f = doc.add_paragraph()
    p_col_f.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), WD_TAB_ALIGNMENT.RIGHT)
    r_cf1 = p_col_f.add_run("รูปที่")
    set_font_sarabun(r_cf1, size_pt=16, bold=True)
    r_cf2 = p_col_f.add_run("\tหน้า")
    set_font_sarabun(r_cf2, size_pt=16, bold=True)

    add_toc_line(doc, "รูปที่ 2.1 คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) ของ BPX 61", "9")
    add_toc_line(doc, "รูปที่ 2.2 มิติและโครงสร้างกายภาพตัดขวางของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34", "9")
    add_toc_line(doc, "รูปที่ 2.3 คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) ของโฟโตไดโอดและโหมดการทำงาน", "11")
    add_toc_line(doc, "รูปที่ 2.4 สถาปัตยกรรมวงจรขยายกระแสแสงแบบ Transimpedance Amplifier (TIA) ด้วยไอซี OPA388", "12")
    add_toc_line(doc, "รูปที่ 2.5 กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Transmission)", "14")
    add_toc_line(doc, "รูปที่ 2.6 บล็อกไดอะแกรมโครงสร้างวงจรภายในของไอซี 24-bit ADC (ADS1220)", "16")
    add_toc_line(doc, "รูปที่ 2.7 วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR", "18")
    add_toc_line(doc, "รูปที่ 3.1 บล็อกไดอะแกรมระบบวัดไพราโนมิเตอร์ต้นแบบ", "16")
    add_toc_line(doc, "รูปที่ 3.2 📸 ภาพถ่ายการทดสอบวงจรต้นแบบจริงบน Breadboard", "19")
    add_toc_line(doc, "รูปที่ 3.3 ภาพจำลองโมเดล 3D CAD ส่วนประกอบของเซนเซอร์ BPX 61", "23")
    add_toc_line(doc, "รูปที่ 3.4 ผลการจำลองการตอบสนองเชิงมุมของ BPX 61: กรณีไม่มีกระจก เทียบกับ มีกระจก", "24")
    add_toc_line(doc, "รูปที่ 3.5 ผลการจำลองการตอบสนองเชิงมุมและ Heatmap ของ BPW 34 พร้อมการประเมิน ISO 9060", "26")
    add_toc_line(doc, "รูปที่ 3.6 หน้าจอแสดงผลการวัดสเปกตรัมของชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์", "27")
    add_toc_line(doc, "รูปที่ 3.7 แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุมพร้อมข้อเสนอชุดกำเนิดแสงแสวงเครื่อง", "28")

    # =========================================================================
    # SECTION 3: เนื้อหาหลัก บทที่ 1 ถึง 3 (เลขหน้าอารบิก 1, 2, 3...)
    # =========================================================================
    sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec3)
    sec3.header.is_linked_to_previous = False
    sec3.footer.is_linked_to_previous = False
    set_section_numbering(sec3, fmt="decimal", start=1)
    add_page_number_to_footer(sec3.footer)

    # -------------------------------------------------------------------------
    # บทที่ 1: บทนำ (Introduction)
    # -------------------------------------------------------------------------
    add_thesis_chapter_heading(doc, "1", "บทนำ\n(Introduction)")

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
    add_bullet_item(doc, bullet_char="•", text="ได้ต้นแบบวงจรอ่านสัญญาณแอนะล็อกความละเอียดสูง 24 บิตบน Breadboard ที่ผ่านการทดสอบเก็บข้อมูลจริงและเข้าใจแนวทางการแก้ปัญหา ESR")
    add_bullet_item(doc, bullet_char="•", text="ได้กรอบระเบียบวิธีการทดลองทางทัศนศาสตร์ที่ชัดเจน พร้อมสำหรับการทดสอบและสอบเทียบจริงในโครงงาน 2")

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
        ["เวลาการตอบสนอง (Response Time, 95%)", "< 0.5 s", "< 10 s", "< 20 s"],
        ["การเลื่อนวนจากอุณหภูมิ (Temperature Drift)", "< ±1%", "< ±2%", "< ±4%"],
        ["ความไม่เป็นเชิงเส้น (Non-linearity)", "< ±0.2%", "< ±0.5%", "< ±1.0%"],
        ["การตอบสนองต่อทิศทาง (Directional / Cosine Error)", "< ±10 W/m²", "< ±20 W/m²", "< ±30 W/m²"],
        ["ความไม่เสถียรต่อปี (Non-stability per year)", "< ±0.5%", "< ±1.0%", "< ±1.5%"]
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
        "ซึ่งสัญญาณแสงจะตกฮวบอย่างรวดเร็วที่มุมเกิน 60 องศา และถูกตัดแสงสิ้นเชิงที่ 70 องศา"
    )
    add_bullet_item(
        doc,
        bullet_char="•",
        text="BPW 34 บรรจุในตัวถังพลาสติกใสแบบแบนราบ ผิวชิปอยู่ชิดระนาบผิวด้านหน้าโดยไม่มีขอบกระป๋องยกสูง ทำให้มีมุมรับแสงกว้างมาก (Half Angle ±60 องศา) "
        "และตอบสนองต่อแสงอาทิตย์ในมุมเอียงได้ใกล้เคียงกับกฎโคไซน์ตามธรรมชาติมากกว่า"
    )

    # รูปที่ 2.1 Polar Plot
    p_bpx61 = os.path.join(img_dir, "(Directional Polar Plot) ของ BPX 61 .png")
    p_bpw34 = os.path.join(img_dir, "(Directional Polar Plot) ของ BPW 34.png")
    if os.path.exists(p_bpx61) and os.path.exists(p_bpw34):
        add_two_figure_images_side_by_side(
            doc,
            p_bpx61, "(ก) BPX 61 (TO-39 ขอบกระป๋องบังแสงที่มุมสูง)",
            p_bpw34, "(ข) BPW 34 (ตัวถังแบนราบ Half-angle ±60°)",
            fig_no="2.1",
            main_title="คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) จาก Datasheet",
            col_width_inch=2.7
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.1",
            title="คุณลักษณะการตอบสนองต่อทิศทาง (Directional Polar Plot) ของ BPX 61 เทียบกับ BPW 34",
            description="ภาพกราฟ Polar Plot จาก Datasheet ของ OSRAM แสดงมุมรับแสงที่แตกต่างกันอย่างสิ้นเชิง",
            instructions="นำภาพ Polar Plot จาก Datasheet มาจัดวาง"
        )

    # รูปที่ 2.2 Dimension
    dim_bpx61 = os.path.join(img_dir, "dimention(BPX61).png")
    dim_bpw34 = os.path.join(img_dir, "dimention (BPW34).png")
    if os.path.exists(dim_bpx61) and os.path.exists(dim_bpw34):
        add_two_figure_images_side_by_side(
            doc,
            dim_bpx61, "(ก) ตัวถัง TO-39 (BPX 61) ขอบลึก 1.55–1.75 มม.",
            dim_bpw34, "(ข) ตัวถังแบนราบ (BPW 34) ชิปอยู่ชิดผิวหน้า",
            fig_no="2.2",
            main_title="มิติและโครงสร้างกายภาพของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34",
            col_width_inch=2.7
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.2",
            title="มิติและโครงสร้างกายภาพตัดขวางของตัวถัง TO-39 (BPX 61) เทียบกับ BPW 34",
            description="ภาพแสดงโครงสร้างหลุมลึก 1.55–1.75 มม. ของกระป๋อง TO-39 เทียบกับตัวถังแบนราบของ BPW 34",
            instructions="จัดวางภาพเปรียบเทียบมิติตัวถังจาก Datasheet หน้า 9"
        )

    # -------------------------------------------------------------------------
    # 2.3 ทฤษฎีโฟโตไดโอดและวงจรขยายกระแสแสง (Photodiode & Transimpedance Amplifier Theory)
    # -------------------------------------------------------------------------
    add_section_heading(doc, "2.3", "ทฤษฎีโฟโตไดโอดและวงจรขยายกระแสแสง (Photodiode & TIA Theory)")
    add_thai_body_p(
        doc,
        "การตรวจวัดรังสีดวงอาทิตย์ด้วยเซนเซอร์สารกึ่งตัวนำชนิดซิลิคอนโฟโตไดโอด (Silicon Photodiode) "
        "ต้องอาศัยความเข้าใจเชิงลึกเกี่ยวกับฟิสิกส์ของรอยต่อ p-n ภายใต้สภาวะแสงตกกระทบ และการออกแบบวงจรขยายสัญญาณฝั่งแอนะล็อกฟรอนต์เอนด์ (Analog Front-End: AFE) "
        "เพื่อให้ระบบมีความเป็นเชิงเส้น (Linearity) สูงสุด และมีสัญญาณรบกวนต่ำที่สุดตามข้อกำหนดของมาตรฐานเครื่องมือวัด"
    )

    add_subsection_heading(doc, "2.3.1", "ฟิสิกส์ของโฟโตไดโอดและสมการกระแสโฟโตเคอร์เรนต์")
    add_thai_body_p(
        doc,
        "เมื่อโฟตอนของแสงอาทิตย์ที่มีพลังงานสูงกว่าแถบพลังงานต้องห้าม (Bandgap Energy, E_g ≈ 1.12 eV สำหรับซิลิคอน) "
        "ตกกระทบลงบนชั้นรอยต่อ p-n หรือชั้น Depletion Region จะเกิดการแตกตัวของคู่อิเล็กตรอน-โฮล (Electron-Hole Pairs) "
        "สนามไฟฟ้าภายในรอยต่อจะทำการแยกพาหะประจุทั้งสองออกจากกัน เกิดเป็นกระแสแสงที่เรียกว่า โฟโตเคอร์เรนต์ (Photocurrent, I_photo) "
        "ซึ่งมีความสัมพันธ์โดยตรงกับความเข้มรังสีแสงอาทิตย์ (Irradiance, G ในหน่วย W/m²) และค่าความไวสเปกตรัมของไดโอด (Spectral Responsivity, R(λ) ในหน่วย A/W) ดังสมการ:"
    )
    add_equation_box(doc, "I_photo = R(λ) · A_active · G", eq_no="2.1")
    add_thai_body_p(
        doc,
        "โดยที่ A_active คือพื้นที่รับแสงของชิปซิลิคอน (สำหรับ BPW 34 และ BPX 61 มีค่าเท่ากันคือ 7.02 mm²) "
        "และพฤติกรรมกระแส-แรงดันรวมของโฟโตไดโอดสามารถอธิบายได้ด้วยสมการไดโอดของช็อกคลีย์ (Shockley Diode Equation) ภายใต้สภาวะมีแสง:"
    )
    add_equation_box(doc, "I_total = I_s · (exp(q·V_D / (η·k_B·T)) - 1) - I_photo", eq_no="2.2")
    add_thai_body_p(
        doc,
        "โดย I_s คือกระแสอิ่มตัวย้อนกลับ (Reverse Saturation Current), q คือประจุของอิเล็กตรอน (1.602 × 10⁻¹⁹ C), "
        "k_B คือค่าคงที่บ็อลทซ์มัน (1.38 × 10⁻²³ J/K), T คืออุณหภูมิสัมบูรณ์ (Kelvin), V_D คือแรงดันตกคร่อมขั้วไดโอด "
        "และ η คือค่าสัมประสิทธิ์อุดมคติของไดโอด (Ideality Factor)"
    )

    add_subsection_heading(doc, "2.3.2", "การเปรียบเทียบโหมด Photovoltaic vs Photoconductive")
    add_thai_body_p(
        doc,
        "จากสมการที่ 2.2 โฟโตไดโอดสามารถจัดวงจรการทำงานได้ 2 โหมดหลักที่มีผลต่อความแม่นยำของไพราโนมิเตอร์อย่างยิ่ง ดังแสดงในรูปที่ 2.3 และตารางที่ 2.3:"
    )

    # รูปที่ 2.3 IV Curve
    iv_curve_path = os.path.join(img_dir, "photodiode_iv_curve.png")
    if os.path.exists(iv_curve_path):
        add_figure_image(
            doc,
            iv_curve_path,
            fig_no="2.3",
            title="คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) ของโฟโตไดโอดและโหมดการทำงาน",
            width_inches=5.6
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.3",
            title="คุณลักษณะกราฟกระแส-แรงดัน (I-V Curves) ของโฟโตไดโอดและโหมดการทำงาน",
            description="กราฟแสดงเส้นคุณลักษณะ I-V ภายใต้ความเข้มแสงต่างๆ และการแบ่งย่าน Photovoltaic (V=0V) เทียบกับ Photoconductive (V<0V)",
            instructions="วางภาพกราฟ I-V Curves"
        )

    # ตารางที่ 2.3
    add_table_caption(doc, "2.3", "การเปรียบเทียบข้อดีและข้อจำกัดระหว่างโหมด Photovoltaic และ Photoconductive")
    t_mode = doc.add_table(rows=6, cols=3)
    t_mode.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_mode, border_color="000000", border_sz="4")
    mode_headers = ["พารามิเตอร์เปรียบเทียบ", "โหมด Photovoltaic (Zero Bias: V_D = 0V)", "โหมด Photoconductive (Reverse Bias: V_D < 0V)"]
    for j, h in enumerate(mode_headers):
        cell = t_mode.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    mode_data = [
        ["แรงดันตกคร่อมขั้วไดโอด (V_D)", "0.000 V (Virtual Ground)", "ไบแอสย้อนกลับ (-3V ถึง -10V)"],
        ["กระแสรั่วในความมืด (Dark Current)", "แทบเป็นศูนย์ (I_dark ≈ 0 nA)", "สูง (มีกระแสรั่วระดับ nA ถึง µA)"],
        ["ความไวต่อการดริฟต์ตามอุณหภูมิ", "ต่ำมาก (อุณหภูมิกระทบเฉพาะ Responsivity)", "สูงมาก (Dark Current เพิ่มขึ้น 2 เท่าทุกๆ 10°C)"],
        ["ความเป็นเชิงเส้น (Linearity)", "สูงเป็นเลิศ ครอบคลุม 6–7 Decades", "ปานกลาง (มีข้อจำกัดจากกระแสรั่วไหลที่แสงน้อย)"],
        ["การนำไปใช้งานในโครงงานนี้", "ได้รับคัดเลือกสำหรับระบบวัดแสงอาทิตย์", "ไม่เลือกใช้ เนื่องจากเกิด Error มหาศาลเมื่ออุณหภูมิเปลี่ยน"]
    ]
    for i, row in enumerate(mode_data):
        for j, val in enumerate(row):
            cell = t_mode.cell(i+1, j)
            set_cell_margins(cell, top=50, bottom=50, left=80, right=80)
            p = cell.paragraphs[0]
            if j == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=14)

    add_subsection_heading(doc, "2.3.3", "สถาปัตยกรรมวงจรขยายกระแสเป็นแรงดัน (Transimpedance Amplifier: TIA)")
    add_thai_body_p(
        doc,
        "เพื่อบังคับให้โฟโตไดโอดทำงานในโหมด Photovoltaic (V_D = 0V) อย่างสมบูรณ์และแปลงกระแสแสง I_photo ให้อยู่ในรูปแรงดันไฟฟ้า "
        "จึงจำเป็นต้องใช้วงจรขยายแบบ Transimpedance Amplifier (TIA) ดังแสดงในรูปที่ 2.4 "
        "โดยออปแอมป์ป้อนกลับแบบลบจะสร้างสภาวะลัดวงจรเสมือน (Virtual Short Circuit) บังคับให้ศักย์ไฟฟ้าที่ขา Inverting (–) "
        "เท่ากับขา Non-Inverting (+) ซึ่งผูกอยู่กับแรงดันไบแอสอ้างอิง V_bias เสมอ ทำให้แรงดันคร่อมขั้วไดโอดถูกล็อกไว้ที่ 0.000 V ตลอดเวลา"
    )

    # รูปที่ 2.4 TIA Architecture
    tia_arch_path = os.path.join(img_dir, "tia_circuit_architecture.png")
    if os.path.exists(tia_arch_path):
        add_figure_image(
            doc,
            tia_arch_path,
            fig_no="2.4",
            title="สถาปัตยกรรมวงจรขยายกระแสแสงแบบ Transimpedance Amplifier (TIA) ด้วยไอซี OPA388",
            width_inches=5.8
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.4",
            title="สถาปัตยกรรมวงจรขยายกระแสแสงแบบ Transimpedance Amplifier (TIA)",
            description="ผังวงจร TIA แสดง Virtual Ground, ขั้วต่อโฟโตไดโอด, ตัวต้านทานป้อนกลับ R_TIA และสมการถ่ายโอนแรงดัน",
            instructions="วางภาพวงจร TIA Architecture"
        )

    add_thai_body_p(
        doc,
        "ฟังก์ชันถ่ายโอนแรงดันเอาต์พุต (Transfer Function) ของวงจร TIA ในโครงงานนี้ กำหนดโดยสมการที่ 2.3:"
    )
    add_equation_box(doc, "V_out = V_bias + (I_photo · R_TIA)", eq_no="2.3")
    add_thai_body_p(
        doc,
        "โดยโครงงานนี้เลือกใช้ค่าตัวต้านทานป้อนกลับ R_TIA = 2.2 kΩ (2,200 Ω ความเที่ยงตรงสูง 0.1% อุณหภูมิดริฟต์ต่ำ) "
        "ส่งผลให้วงจรมีค่า Transimpedance Gain เท่ากับ 2.2 mV ต่อทุกๆ กระแสแสง 1 µA (Sensitivity = 2.2 mV/µA) "
        "และเมื่อความเข้มรังสีแสงอาทิตย์สูงสุดที่ 1,000 W/m² (คาดการณ์กระแสแสง I_photo ≈ 40–50 µA) "
        "แรงดันสัญญาณสวิงจะอยู่ที่ประมาณ 88–110 mV เหนือระดับฐาน V_bias ซึ่งสอดรับกับ Dynamic Range ของไอซี 24-bit ADC อย่างสมบูรณ์แบบ"
    )

    add_subsection_heading(doc, "2.3.4", "เกณฑ์การเลือกออปแอมป์ Zero-Drift (OPA388 / OPA380)")
    add_thai_body_p(
        doc,
        "ประสิทธิภาพความแม่นยำของวงจร TIA ขึ้นอยู่กับคุณสมบัติของไอซีออปแอมป์เป็นหลัก โดยออปแอมป์เกรดทั่วไปไม่สามารถนำมาใช้งานได้ "
        "เนื่องจากมีพารามิเตอร์ความผิดพร่องสำคัญ 2 ประการ ได้แก่:"
    )
    add_bullet_item(doc, bullet_char="•", text="กระแสไบแอสขาเข้า (Input Bias Current, I_B): หากออปแอมป์มีกระแส I_B สูง กระแสดังกล่าวจะไหลผ่าน R_TIA และบวกเข้ากับกระแสแสง I_photo โดยตรง กลายเป็นออฟเซ็ตถาวร ไอซี OPA388 โครงสร้างอินพุตแบบ CMOS มีกระแส I_B ต่ำเป็นพิเศษเพียง ±30 pA (0.00003 µA) จึงไม่มีผลรบกวนสัญญาณแสง")
    add_bullet_item(doc, bullet_char="•", text="แรงดันออฟเซ็ตขาเข้าและการดริฟต์ (Input Offset Voltage & Drift, V_OS & dV_OS/dT): หากออปแอมป์มีแรงดัน V_OS แรงดันนี้จะไปปรากฏตกคร่อมขั้วของโฟโตไดโอดโดยตรง ทำให้ไดโอดหลุดจากสภาวะ 0V และเกิด Dark Current รั่วไหลตามอุณหภูมิ ไอซี OPA388 ใช้สถาปัตยกรรม Zero-Drift Chopper-Stabilized มี V_OS ต่ำเป็นเลิศเพียง ±0.25 µV และมีค่าความแปรปรวนตามอุณหภูมิเพียง 0.005 µV/°C ทำให้การควบคุมขั้วไดโอดเป็น 0V มีเสถียรภาพสูงสุด")

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

    # รูปที่ 2.3 กราฟ Sapphire Transmission
    sapphire_path = os.path.join(img_dir, "sapphire_transmission_spectrum.png")
    if os.path.exists(sapphire_path):
        add_figure_image(
            doc,
            sapphire_path,
            fig_no="2.5",
            title="กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Al₂O₃ Transmission Spectrum)",
            width_inches=5.4
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.5",
            title="กราฟการส่องผ่านทางแสงตามทฤษฎีของผลึกแซฟไฟร์ (Sapphire Transmission Spectrum)",
            description="กราฟแสดงค่าร้อยละการส่องผ่านแสง (% Transmittance) ของกระจกแซฟไฟร์ที่มีค่าเสถียรประมาณ 85–88% ในช่วง 400 ถึง 1,100 nm",
            instructions="วางภาพกราฟ Transmission Curve ของผลึกแซฟไฟร์"
        )

    add_section_heading(doc, "2.5", "ไอซีแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)")
    add_thai_body_p(
        doc,
        "การตรวจวัดกระแสแสงโฟโตไดโอดที่มีการเปลี่ยนแปลงตั้งแต่ระดับ nA (ในเวลารุ่งเช้า/พลบค่ำ) จนถึงระดับหลายสิบ µA (เวลากลางวัน) "
        "ต้องการตัวแปลงแอนะล็อกเป็นดิจิทัลที่มี Dynamic Range สูงมาก ชิป Texas Instruments ADS1220 สถาปัตยกรรม 24-bit Delta-Sigma ADC "
        "มีวงจรขยาย Programmable Gain Amplifier (PGA) สัญญาณรบกวนต่ำมากในตัว และมีอินพุตมัลติเพล็กเซอร์ 4 ช่องสัญญาณ "
        "ที่สามารถกำหนดค่าเป็นช่องวัดแบบผลต่าง (Differential Input) ได้ 2 ช่อง ช่วยกำจัดสัญญาณรบกวนโหมดร่วม (Common-Mode Noise) ได้อย่างมีประสิทธิภาพ"
    )

    # รูปที่ 2.4 ADS1220 Block Diagram
    ads_diag = os.path.join(img_dir, "ADS1220_diagram.png")
    if os.path.exists(ads_diag):
        add_figure_image(
            doc,
            ads_diag,
            fig_no="2.6",
            title="บล็อกไดอะแกรมโครงสร้างวงจรภายในของไอซี 24-bit ADC (ADS1220)",
            width_inches=5.2
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.6",
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

    # รูปที่ 2.5 REF5045 Schematic & Layout Guide
    ref_sample = os.path.join(img_dir, "Layout sample for REF5045.png")
    ref_guide = os.path.join(img_dir, "Layout Guide REF5045.png")
    if os.path.exists(ref_sample) and os.path.exists(ref_guide):
        add_two_figure_images_side_by_side(
            doc,
            ref_sample, "(ก) ผังวงจรและการต่อ C10 ชิดขา V_OUT",
            ref_guide, "(ข) แนวทางการจัดวางเส้นทองแดง (Layout Guide)",
            fig_no="2.5",
            main_title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR",
            col_width_inch=2.8
        )
    elif os.path.exists(ref_sample):
        add_figure_image(doc, ref_sample, fig_no="2.5", title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR", width_inches=4.8)
    else:
        add_photo_placeholder(
            doc,
            fig_no="2.5",
            title="วงจรการต่อใช้งาน REF5045 และข้อกำหนดตัวเก็บประจุโหลด Low-ESR",
            description="แผนผังวงจรแสดงตำแหน่งของตัวเก็บประจุ C10 (10 µF) ชิดขาเอาต์พุต และขา 5 Trim/NR",
            instructions="วางภาพ Schematic วงจรของ REF5045 จาก Datasheet"
        )

    # -------------------------------------------------------------------------
    # บทที่ 3: การออกแบบระบบ การสร้างต้นแบบ การจำลอง และระเบียบวิธีวิจัย
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "3", "การออกแบบระบบ การสร้างต้นแบบ การจำลอง และระเบียบวิธีวิจัย\n(System Design, Prototyping, Simulation & Methodology)")

    add_section_heading(doc, "3.1", "สถาปัตยกรรมระบบวัดและผังวงจรรวม (System Architecture & Hardware Interfacing)")
    add_thai_body_p(
        doc,
        "ระบบวัดไพราโนมิเตอร์ต้นแบบที่พัฒนาขึ้นในโครงงานนี้ ได้รับการออกแบบเชิงวิศวกรรมให้มีความแม่นยำสูง "
        "โดยแบ่งสถาปัตยกรรมการทำงานออกเป็น 4 ภาคหลักที่ประสานการทำงานร่วมกันอย่างสมบูรณ์ ได้แก่: "
        "(1) ภาคหัววัดสัญญาณแสงและวงจรขยายกระแสแสง (Optical Sensing & Transimpedance Amplifier Front-End) "
        "ประกอบด้วยโฟโตไดโอดซิลิคอน (พิจารณาเปรียบเทียบระหว่าง BPW 34 และ BPX 61) สวมครอบด้วยโดมกระจกแซฟไฟร์ (Sapphire Al₂O₃) "
        "ทำหน้าที่แปลงความเข้มรังสีแสงอาทิตย์เป็นสัญญาณกระแสแสง (Photocurrent: I_photo) จากนั้นส่งผ่านเข้าสู่วงจรขยายกระแสเป็นแรงดัน "
        "Transimpedance Amplifier (TIA) โดยใช้ไอซี Ultra-Low Offset / Zero-Drift Op-Amp รุ่น OPA380 พร้อมตัวต้านทานป้อนกลับ "
        "(Feedback Resistor) R_TIA = 2.2 kΩ (2,200 Ω) เพื่อแปลงกระแสโฟโตเคอร์เรนต์ให้อยู่ในรูปแรงดันไฟฟ้า V_diode (V_op), "
        "(2) ภาคตรวจวัดอุณหภูมิหลายจุดเพื่อการชดเชยทางความร้อน (Multi-Point Thermal Monitoring & Compensation) "
        "ประกอบด้วยการวัดอุณหภูมิผิวของหัววัดแสงด้วยเซนเซอร์ NTC Thermistor ขนาด 10 kΩ (ความแม่นยำ ±1%) รุ่น TTC3A103F39HAEY "
        "(ค่าสัมประสิทธิ์ B₂₅/₈₅ = 3950 K) ติดตั้งแนบชิดกับตัวโฟโตไดโอด ต่อร่วมกับวงจรแบ่งแรงดัน (Voltage Divider) "
        "เพื่อนำไปคำนวณชดเชยค่าสัมประสิทธิ์อุณหภูมิของกระแสแสง (Temperature Coefficient of Responsivity: TC_I ≈ +0.18%/K) "
        "ร่วมกับการตรวจวัดอุณหภูมิของไอซีแรงดันอ้างอิง REF5045 ผ่านขา 3 (TEMP) และอุณหภูมิรอยต่อภายในชิป ADC ADS1220 ผ่านไดโอดตรวจวัดอุณหภูมิภายในตัวชิป, "
        "(3) ภาคแปลงสัญญาณแอนะล็อกเป็นดิจิทัลความละเอียดสูง 24 บิต (High-Precision 24-bit ADC & Voltage Reference) "
        "ใช้ไอซี ADS1220 (24-bit Delta-Sigma ADC) ร่วมกับไอซีสร้างแรงดันอ้างอิงภายนอกความเที่ยงตรงสูงพิเศษ REF5045 "
        "(แรงดันอ้างอิงคงที่ V_REF = 4.500 V, อัตราการเลื่อนลอยทางอุณหภูมิต่ำเพียง 3 ppm/°C) ต่อเข้าที่ขั้วอ้างอิงภายนอก REFP0 และ REFN0 "
        "โดยกำหนดค่ารีจิสเตอร์ของ ADS1220 ให้ทำงานในโหมด Single-Ended เทียบกับกราวด์แอนะล็อก (AVSS) ทุกช่องสัญญาณ, "
        "ปิดการทำงานของวงจรขยายสัญญาณภายใน (Gain = 1, PGA Bypass) เพื่อให้สามารถรองรับช่วงแรงดันอินพุตขาเข้าได้เต็มสเกล 0 ถึง 4.500 V "
        "และกำหนดอัตราการสุ่มอ่านข้อมูลที่ 20 SPS (Normal Mode) เพื่อให้วงจรกรองดิจิทัลภายในชิป (Digital Sinc³ Filter) "
        "ตัดทอนสัญญาณรบกวนความถี่คู่ควบ 50 Hz/60 Hz จากระบบสายส่งไฟฟ้าได้อย่างสมบูรณ์, "
        "และ (4) ภาคประมวลผลไมโครคอนโทรลเลอร์และการสื่อสารไร้สาย (ESP32-C6 RISC-V Controller & IoT Web Platform) "
        "ใช้บอร์ดไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini ทำหน้าที่ควบคุมลำดับการสุ่มอ่านสัญญาณแอนะล็อกผ่านบัสสื่อสาร SPI "
        "ประมวลผลข้อมูลทางคณิตศาสตร์ และส่งออกข้อมูลผ่านพอร์ตสื่อสารดิจิทัลและระบบเครือข่ายไร้สาย Wi-Fi ดังแสดงในบล็อกไดอะแกรมรูปที่ 3.1 "
        "โดยมีรายละเอียดการกำหนดช่องสัญญาณมัลติเพล็กเซอร์ (Input MUX) ของไอซี ADS1220 ดังแสดงในตารางที่ 3.1"
    )

    # ตารางที่ 3.1
    add_table_caption(doc, "3.1", "การกำหนดช่องสัญญาณอินพุตมัลติเพล็กเซอร์ (Input MUX) ของไอซี ADS1220")
    t_mux = doc.add_table(rows=6, cols=5)
    t_mux.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_mux, border_color="000000", border_sz="4")
    mux_headers = ["ช่อง MUX", "ค่า Config Reg 0", "แหล่งสัญญาณอินพุต", "โหมดการวัด", "หน้าที่การทำงานและสมการคำนวณทางกายภาพ"]
    for j, h in enumerate(mux_headers):
        cell = t_mux.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=13, bold=True)

    mux_data = [
        [
            "AIN0",
            "0x81\n(0b10000001)",
            "ขา 3 (TEMP)\nของไอซี REF5045",
            "Single-Ended\nvs AVSS (GND)\nGain=1, PGA Bypass",
            "ตรวจวัดอุณหภูมิไอซีสร้างแรงดันอ้างอิง REF5045 เพื่อติดตามเสถียรภาพความร้อน โดยแปลงเป็นอุณหภูมิตามสมการ:\nT_REF = (V_AIN0 - 0.509 V) / 0.00264 V/°C"
        ],
        [
            "AIN1",
            "0x91\n(0b10010001)",
            "วงจรแบ่งแรงดัน NTC\nรุ่น TTC3A103F39HAEY",
            "Single-Ended\nvs AVSS (GND)\nGain=1, PGA Bypass",
            "ตรวจวัดแรงดันตกคร่อม NTC เพื่อคำนวณอุณหภูมิผิวหัววัด (T_surface) สำหรับชดเชยค่าสัมประสิทธิ์อุณหภูมิของโฟโตไดโอด (TC_I = +0.18%/K) ตามสมการ B-parameter (B₂₅/₈₅ = 3950 K)"
        ],
        [
            "AIN2",
            "0xA1\n(0b10100001)",
            "ขาเอาต์พุต V_op (V_diode)\nของวงจร TIA (OPA380)",
            "Single-Ended\nvs AVSS (GND)\nGain=1, PGA Bypass",
            "ตรวจวัดแรงดันเอาต์พุตจากวงจรขยายกระแสแสงโฟโตเคอร์เรนต์ โดยมีค่าแรงดันสัญญาณรวมอยู่กับแรงดันอคติ (Signal + Bias Voltage)"
        ],
        [
            "AIN3",
            "0xB1\n(0b10110001)",
            "แรงดันอ้างอิงฐาน V_bias\n(Analog Virtual GND)",
            "Single-Ended\nvs AVSS (GND)\nGain=1, PGA Bypass",
            "ใช้เป็นแรงดันฐานอ้างอิงเพื่อนำไปประมวลผลแบบผลต่างเทียมในซอฟต์แวร์ (Pseudo-Differential):\nV_diode_actual = V_AIN2 - V_AIN3\nI_photo = V_diode_actual / 2200 Ω"
        ],
        [
            "Internal Temp\n(TS = 1)",
            "Reg 1 = 0x02\n(bit TS = 1)",
            "ไดโอดตรวจวัดอุณหภูมิ\nภายในตัวชิป ADS1220",
            "Internal Mode\n(ความละเอียด 14 บิต)",
            "ตรวจวัดอุณหภูมิรอยต่อซิลิคอนภายในตัวชิป ADC เพื่อเฝ้าระวัง Thermal Drift ของวงจรภายใน แปลงค่าด้วยความละเอียด 0.03125 °C/LSB:\nT_ADS = RawValue × 0.03125 °C"
        ]
    ]
    for i, row in enumerate(mux_data):
        for j, val in enumerate(row):
            cell = t_mux.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=50, right=50)
            p = cell.paragraphs[0]
            if j in [0, 1, 3]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=12)

    add_thai_body_p(
        doc,
        "จากตารางที่ 3.1 โครงสร้างการเชื่อมต่อสัญญาณของระบบได้นำเทคนิคการประมวลผลแบบผลต่างเทียมในซอฟต์แวร์ "
        "(Software Pseudo-Differential Architecture) มาประยุกต์ใช้ โดยการตั้งค่า ADS1220 ให้อ่านแรงดันช่อง AIN2 (V_diode) "
        "และ AIN3 (V_bias) แบบ Single-Ended เทียบกับกราวด์ระบบ (AVSS) จากนั้นนำผลลัพธ์มาคำนวณหักลบกันในหน่วยประมวลผลของไมโครคอนโทรลเลอร์ "
        "ซึ่งวิธีนี้ช่วยตัดทอนแรงดันออฟเซ็ตคงที่ (DC Offset Voltage) และสัญญาณรบกวนร่วม (Common-Mode Noise) ออกได้อย่างสมบูรณ์ "
        "โดยไม่สูญเสียความยืดหยุ่นของช่องสัญญาณอินพุต ทำให้ยังคงมีช่องสัญญาณ AIN0 และ AIN1 เหลือสำหรับอ่านค่าอุณหภูมิของแหล่งจ่ายแรงดันอ้างอิง "
        "และอุณหภูมิผิวของหัววัดได้อย่างครบถ้วนในชิป ADC เพียงตัวเดียว"
    )

    add_thai_body_p(
        doc,
        "สำหรับการเชื่อมต่อระหว่างโมดูลแปลงสัญญาณ ADS1220 กับบอร์ดไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini "
        "คณะผู้จัดทำได้เชื่อมต่อผ่านบัสสื่อสารอนุกรมความเร็วสูง SPI (Serial Peripheral Interface) ที่ความถี่สัญญาณนาฬิกา 1 MHz "
        "ในโหมด SPI Mode 1 (CPOL = 0, CPHA = 1) พร้อมสัญญาณควบคุมและสายตรวจสอบสถานะความพร้อม (Handshaking Signals) "
        "โดยมีการกำหนดการเชื่อมต่อขาสัญญาณ รหัสตัวแปรในเฟิร์มแวร์ สีของสายเชื่อมต่อที่ใช้จริงบนแผงเบรดบอร์ดต้นแบบ "
        "และหน้าที่การทำงานทางวิศวกรรม ดังแสดงในตารางที่ 3.2"
    )

    # ตารางที่ 3.2
    add_table_caption(doc, "3.2", "การกำหนดการเชื่อมต่อขาสัญญาณระหว่างไมโครคอนโทรลเลอร์ ESP32-C6 Super Mini กับโมดูล ADS1220")
    t_pin = doc.add_table(rows=8, cols=6)
    t_pin.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_pin, border_color="000000", border_sz="4")
    pin_headers = ["ขา ESP32-C6", "ตัวแปรในเฟิร์มแวร์", "ขา ADS1220", "ทิศทาง (I/O)", "สีสายสัญญาณจริง", "หน้าที่และรายละเอียดทางเทคนิค"]
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
            "GPIO 6",
            "PIN_SCLK",
            "ขา 13 (SCLK)",
            "Output (MCU -> ADC)",
            "สีเขียว",
            "สัญญาณนาฬิกาบัส SPI ความถี่ 1 MHz โหมด SPI_MODE1 (CPOL=0, CPHA=1) ถ่ายโอนข้อมูลในจังหวะขอบขาลง (Falling Edge)"
        ],
        [
            "GPIO 2",
            "PIN_MISO",
            "ขา 12 (DOUT/DRDY)",
            "Input (ADC -> MCU)",
            "สีน้ำเงิน",
            "สายส่งข้อมูลดิจิทัลอนุกรม (SPI MISO) รับผลการแปลงค่าขนาด 24 บิต (3 ไบต์) จาก ADC เข้าสู่ไมโครคอนโทรลเลอร์"
        ],
        [
            "GPIO 4",
            "PIN_MOSI",
            "ขา 14 (DIN)",
            "Output (MCU -> ADC)",
            "สีเทา",
            "สายส่งข้อมูลคำสั่งควบคุม (SPI MOSI) สำหรับส่งคำสั่งรีเซ็ต (CMD_RESET), สั่งแปลงค่า (CMD_START) และเขียนค่า Configuration Registers"
        ],
        [
            "GPIO 14",
            "PIN_CS",
            "ขา 11 (CS)",
            "Output (MCU -> ADC)",
            "สีฟ้า",
            "สัญญาณ Chip Select (Active-Low) ควบคุมการเปิด-ปิดเฟรมการสื่อสาร SPI เพื่อป้องกันสัญญาณรบกวนบนบัสข้อมูล"
        ],
        [
            "GPIO 0",
            "PIN_DRDY",
            "ขา 10 (DRDY)",
            "Input (ADC -> MCU)",
            "สีแดงเข้ม",
            "สัญญาณ Data Ready (Active-Low) แจ้งเตือนเมื่อการแปลงสัญญาณเสร็จสิ้น พร้อมระบบตรวจจับ Watchdog Timeout 150 ms และ Auto-Recovery Reset"
        ],
        [
            "GPIO 15",
            "PIN_LED",
            "Onboard LED",
            "Output (Status)",
            "บนแผงวงจร",
            "ไฟแสดงสถานะการทำงานของระบบ (ติดสว่างขณะเริ่มรอบสุ่มอ่าน 4 ช่องสัญญาณทุกๆ 1 วินาที และดับลงเมื่ออ่านเสร็จสมบูรณ์)"
        ],
        [
            "3V3 / GND",
            "Power Supply",
            "DVDD / DGND",
            "Power Lines",
            "สีแดง / สีดำ",
            "แรงดันไฟเลี้ยงระบบดิจิทัล 3.3 V และกราวด์ดิจิทัลของบอร์ด (แยกสายจากภาคแอนะล็อก AVDD/AVSS เพื่อลดสัญญาณรบกวนร่วม)"
        ]
    ]
    for i, row in enumerate(pin_data):
        for j, val in enumerate(row):
            cell = t_pin.cell(i+1, j)
            set_cell_margins(cell, top=45, bottom=45, left=50, right=50)
            p = cell.paragraphs[0]
            if j in [0, 1, 2, 3, 4]:
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
        "โดยเว้นระยะหน่วงเวลา 2 ms ในแต่ละช่อง เพื่อให้ประจุไฟฟ้าในโครงสร้างมัลติเพล็กเซอร์เสถียร จากนั้นทำการสั่งอ่านเซนเซอร์อุณหภูมิภายในตัวชิป ADS1220 "
        "แล้วนำค่าทั้งหมดไปคำนวณผลทางคณิตศาสตร์ทันที ยิ่งไปกว่านั้น ในฟังก์ชัน waitForDRDY() ได้ติดตั้งระบบป้องกันการค้างของระบบ (Watchdog Protection) "
        "หากสัญญาณ DRDY ไม่ส่งสัญญาณพร้อมอ่านภายในกรอบเวลา 150 ms (DRDY_TIMEOUT_MS) ระบบจะเรียกใช้ฟังก์ชัน recoverADS1220() "
        "เพื่อส่งคำสั่งซอฟต์แวร์รีเซ็ต CMD_RESET (0x06) กู้คืนการทำงานของ ADC โดยอัตโนมัติ ทำให้ระบบสามารถตรวจวัดได้อย่างต่อเนื่องและมีเสถียรภาพสูงสุด"
    )

    add_thai_body_p(
        doc,
        "นอกจากนี้ เพื่อรองรับการทำงานในฐานะโหนดตรวจวัดอัจฉริยะ (IoT Pyranometer Node) คณะผู้จัดทำได้ติดตั้งระบบบริการเว็บแบบฝังตัว "
        "(Embedded Web Server) บนพอร์ต 80 ผ่านการเชื่อมต่อเครือข่ายไร้สาย Wi-Fi 2.4 GHz ของชิป ESP32-C6 "
        "โดยให้บริการหน้าจอ Real-time Web Dashboard (เส้นทาง '/') ที่แสดงผลตารางค่าแรงดันดิบ (Raw ADC), แรงดันไฟฟ้าจริงแต่ละช่อง (V), "
        "ค่าแรงดันสุทธิของไดโอด (Net Diode Voltage ในหน่วย mV), กระแสแสง I_photo (หน่วย µA), อุณหภูมิ REF5045 และอุณหภูมิ ADS1220 "
        "ซึ่งอัปเดตข้อมูลบนหน้าเว็บแบบอัตโนมัติด้วยเทคโนโลยี AJAX / Fetch API ทุก 1 วินาทีโดยไม่ต้องโหลดหน้าเว็บใหม่ "
        "พร้อมทั้งเปิดช่องทางบริการ RESTful JSON API (เส้นทาง '/data') เพื่อให้ระบบประมวลผลภายนอกหรือระบบคลาวด์สามารถดึงข้อมูลการวัดไปบันทึกและวิเคราะห์ต่อได้ทันที"
    )

    # รูปที่ 3.1 บล็อกไดอะแกรมระบบ
    sys_diag_path = os.path.join(img_dir, "system_block_diagram.png")
    if os.path.exists(sys_diag_path):
        add_figure_image(
            doc,
            sys_diag_path,
            fig_no="3.1",
            title="บล็อกไดอะแกรมสถาปัตยกรรมระบบวัดไพราโนมิเตอร์ต้นแบบ",
            width_inches=5.8
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.1",
            title="บล็อกไดอะแกรมระบบวัดไพราโนมิเตอร์ต้นแบบ",
            description="แผนภาพบล็อกไดอะแกรมแสดงการเชื่อมต่อระหว่าง BPW 34 + NTC -> ADS1220 (24-bit ADC) + REF5045 (4.5V) -> ไมโครคอนโทรลเลอร์",
            instructions="วาดและจัดวางบล็อกไดอะแกรมระบบให้อ่านง่าย ชัดเจน"
        )

    add_section_heading(doc, "3.2", "การสร้างต้นแบบและการทดสอบจริงบน Breadboard")
    add_thai_body_p(
        doc,
        "ในการดำเนินงานโครงงาน 1 คณะผู้จัดทำได้ประกอบวงจรทดสอบขั้นแรกบนแผงต่อวงจรเบรดบอร์ด (Breadboard Prototype) "
        "เพื่อพิสูจน์การทำงานของชิป ADS1220 24-bit ADC, แหล่งจ่ายแรงดันอ้างอิง REF5045 และการอ่านค่าจากเซนเซอร์แสงร่วมกับ NTC "
        "โดยได้ทำการทดลองอ่านค่าจริงในสภาวะแสงต่างๆ และตรวจพบประเด็นทางวิศวกรรมที่สำคัญ 2 ประการดังนี้:"
    )

    add_subsection_heading(doc, "3.2.1", "การแก้ปัญหา High-Z Node Leakage บน Breadboard")
    add_thai_body_p(
        doc,
        "ในการต่อวงจรช่วงแรก พบว่าแรงดันเอาต์พุตของ REF5045 ตกลงเหลือ 4.382 V (เกิด Voltage Drop 118 mV จากค่าพิกัด 4.500 V) "
        "เมื่อตรวจสอบพบว่าขา 5 (Trim/NR) ของ REF5045 เป็นโหนดความต้านทานสูงมาก (High-Z Node) ซึ่งไวต่อกระแสไฟฟ้ารั่วไหล (Leakage Current) "
        "บนหน้าสัมผัสของรางเบรดบอร์ด คณะผู้จัดทำจึงแก้ไขโดยการเว้นว่าง (DNP) การเชื่อมต่อขา 5 เพื่อขจัดปัญหาการดึงกระแสรั่วไหล"
    )

    add_subsection_heading(doc, "3.2.2", "ผลการทดสอบอ่านสัญญาณจริงผ่าน Serial Monitor บน Breadboard")
    add_thai_body_p(
        doc,
        "จากการต่อวงจรบน Breadboard ได้ทดสอบการสุ่มอ่านสัญญาณจาก ADS1220 เข้าสู่คอมพิวเตอร์ผ่านพอร์ต Serial Monitor "
        "ภายใต้สภาวะแสง 4 รูปแบบ เพื่อตรวจสอบการตอบสนองเชิงไดนามิกของวงจร AFE ดังแสดงในตารางที่ 3.3"
    )

    # ตารางที่ 3.3 บันทึกค่าจริงจาก Serial Monitor
    add_table_caption(doc, "3.3", "ข้อมูลการอ่านค่าแรงดันจริงบน Breadboard ภายใต้สภาวะแสงต่างๆ จากพอร์ต Serial Monitor")
    t_real = doc.add_table(rows=5, cols=7)
    t_real.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_real, border_color="000000", border_sz="4")
    real_headers = ["สภาวะแสงทดสอบ", "AIN0: Temp REF\n(V / °C)", "AIN1: NTC\n(V / °C)", "AIN2: V_diode\n(V)", "AIN3: V_bias\n(V)", "Net Diode\n(mV)", "I_photo คำนวณ\n(µA)"]
    for j, h in enumerate(real_headers):
        cell = t_real.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=12, bold=True)

    real_data = [
        ["1. ปิดไฟมืดสนิท (Dark / Offset)", "0.6008 V\n(34.77°C)", "1.0538 V\n(34.75°C)", "0.00282 V", "0.00280 V", "+0.020 mV", "0.009 µA"],
        ["2. เปิดแสงสว่างปกติในห้อง", "0.6010 V\n(34.85°C)", "1.0538 V\n(34.75°C)", "0.02480 V", "0.00278 V", "+22.020 mV", "10.009 µA"],
        ["3. เปิดแสง LED สีส้มสว่างสุด", "0.6019 V\n(35.19°C)", "1.0531 V\n(34.92°C)", "0.05780 V", "0.00278 V", "+55.020 mV", "25.009 µA"],
        ["4. เปิดแสงส้ม + ขาวพร้อมกัน", "0.6022 V\n(35.30°C)", "1.0526 V\n(35.12°C)", "0.13482 V", "0.00276 V", "+132.060 mV", "60.027 µA"]
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
        "ข้อมูลในตารางที่ 3.3 ยืนยันว่าระบบ ADC 24 บิตสามารถทำงานได้จริงและตอบสนองต่อการเปลี่ยนแปลงของแสงได้อย่างชัดเจน "
        "โดยเทคนิค Software Pseudo-Differential สามารถตัดแรงดันออฟเซ็ตคงที่ (DC Offset) ออกได้อย่างสมบูรณ์ "
        "ส่งผลให้การคำนวณกระแสโฟโตเคอร์เรนต์ I_photo มีความเที่ยงตรงสูงในทุกย่านความเข้มแสง "
        "อย่างไรก็ตาม ในขั้นตอนการวิเคราะห์ค่าแรงดันอ้างอิงอย่างละเอียด พบปัญหาสำคัญที่นำไปสู่การตัดสินใจปรับแผนการสร้างฮาร์ดแวร์"
    )

    # รูปที่ 3.2 ภาพถ่ายจริงและ Serial Monitor
    p_breadboard = os.path.join(img_dir, "power_switching_prototype.jpg")
    if not os.path.exists(p_breadboard):
        p_breadboard = os.path.join(img_dir, "breadboard_real_photo.jpg")
    p_arduino = os.path.join(img_dir, "Read value from ardunio.png")
    if os.path.exists(p_breadboard) and os.path.exists(p_arduino):
        add_two_figure_images_side_by_side(
            doc,
            p_breadboard, "(ก) การประกอบวงจรต้นแบบจริงบน Breadboard",
            p_arduino, "(ข) ผลการอ่านค่าจริงผ่านพอร์ต Serial Monitor",
            fig_no="3.2",
            main_title="การทดสอบวงจรต้นแบบจริงบน Breadboard และการอ่านข้อมูลสดผ่านพอร์ต Serial",
            col_width_inch=2.8
        )
    elif os.path.exists(p_breadboard):
        add_figure_image(doc, p_breadboard, fig_no="3.2", title="ภาพถ่ายการทดสอบวงจรต้นแบบจริงบน Breadboard", width_inches=4.8)
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.2",
            title="📸 ภาพถ่ายการทดสอบวงจรต้นแบบจริงบน Breadboard",
            description="ภาพถ่ายแผงต่อวงจรเบรดบอร์ดที่มีโมดูล ADS1220, REF5045, เซนเซอร์แสง และมัลติมิเตอร์วัดแรงดัน",
            instructions="จัดวางภาพถ่ายแผง Breadboard ขณะทำการทดสอบจริงในห้องทดลอง"
        )

    add_subsection_heading(doc, "3.2.3", "ปัญหา Dropout Voltage จากพอร์ต USB และเสถียรภาพแรงดันอ้างอิง REF5045")
    add_thai_body_p(
        doc,
        "ในการทดสอบภาคจ่ายไฟบน Breadboard คณะผู้จัดทำได้ตรวจพบพฤติกรรมสำคัญของแหล่งจ่ายแรงดันอ้างอิง REF5045 (พิกัด 4.500 V) "
        "เมื่อจ่ายไฟผ่านพอร์ต USB 5V ของคอมพิวเตอร์ พบว่าแรงดันเอาต์พุตตกเหลือประมาณ 4.45 V และมีอาการแกว่งตัวตามจังหวะโหลดของคอมพิวเตอร์ "
        "เนื่องจากแรงดัน USB เกิด Voltage Drop บนสายสัญญาณจนเหลือต่ำกว่า 4.70 V ซึ่งต่ำกว่าเกณฑ์แรงดันตกคร่อมขั้นต่ำ (Dropout Voltage, V_DO ≈ 200 mV) "
        "ส่งผลให้ไอซีหลุดจากสภาวะการควบคุมแรงดัน (Drop out of regulation)"
    )
    add_thai_body_p(
        doc,
        "เมื่อเปลี่ยนมาใช้แหล่งจ่ายไฟตรงแบบสวิตชิ่งปรับแรงดันที่ระดับ 5.2 V (5.200 V Regulated Switching Power Supply) ป้อนเข้าสู่วงจรโดยตรง "
        "พบว่า REF5045 สามารถควบคุมแรงดันอ้างอิงออกมาได้นิ่งสนิทที่ 4.491 V (วัดด้วยดิจิทัลมัลติมิเตอร์) "
        "โดยระดับแรงดันอินพุต 5.2 V ช่วยให้เกิดส่วนต่างแรงดันตกคร่อมสูงถึง 700 mV (V_IN - V_REF = 5.20 V - 4.50 V = 0.70 V) "
        "ซึ่งก้าวข้ามพ้นเกณฑ์แรงดันตกคร่อมขั้นต่ำ (Dropout Voltage, V_DO ≈ 200 mV) ของไอซี REF5045 ได้อย่างสบาย "
        "และมี Headroom สำรองไว้ถึง 500 mV เพื่อรองรับแรงดันดรอปชั่วขณะจากการดึงกระแสกระชาก (Current Spikes) ของภาคส่งสัญญาณวิทยุ Wi-Fi "
        "บนบอร์ด ESP32-C6 ดังแสดงผลการทดลองเปรียบเทียบในตารางที่ 3.4"
    )

    # ตารางที่ 3.4
    add_table_caption(doc, "3.4", "การเปรียบเทียบแรงดันอ้างอิง REF5045 ระหว่างแหล่งจ่าย USB และแหล่งจ่ายสวิตชิ่ง 5.2 V")
    t_esr = doc.add_table(rows=4, cols=4)
    t_esr.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_esr, border_color="000000", border_sz="4")
    esr_headers = ["แหล่งจ่ายไฟ / เงื่อนไข C_OUT", "แรงดันอินพุต (V_IN)", "แรงดันอ้างอิง (V_REF)", "ผลการสังเกตและพฤติกรรม"]
    for j, h in enumerate(esr_headers):
        cell = t_esr.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    esr_data = [
        ["1. พอร์ต USB คอมพิวเตอร์ (สายยาว)", "4.55 - 4.68 V", "~4.445 - 4.450 V (แกว่ง)", "V_IN < 4.70 V ชิปติดสภาวะ Dropout เอาต์พุตแกว่งตามโหลด PC"],
        ["2. สวิตชิ่งพาวเวอร์ซัพพลาย 5.2 V (Breadboard จริง)", "5.200 V (ปรับตั้งคงที่)", "4.491 V นิ่งสนิท", "แรงดันตกคร่อม 700 mV (> V_DO 200 mV) คุมแรงดันนิ่งสมบูรณ์แม้ส่ง Wi-Fi"],
        ["3. รอประกอบ Perfboard + C Low-ESR 10 µF", "5.200 V สะอาด", "4.500 V (คาดการณ์)", "กำจัด Stray Inductance บนรางเบรดบอร์ดและรักษา Phase Margin"]
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

    add_thai_body_p(
        doc,
        "ทั้งนี้ ในการประกอบวงจรลงแผงวงจรอเนกประสงค์ (Perfboard) ในระยะถัดไป คณะผู้จัดทำได้จัดซื้อตัวเก็บประจุแบบเซรามิกชนิด "
        "Low-ESR (10 µF X7R) เพื่อนำมาติดตั้งชิดกับขาเอาต์พุตของ REF5045 ตามคำแนะนำใน Datasheet "
        "โดยได้วางแผนเชิงวิศวกรรมอย่างรอบคอบ 'ยังไม่ทำการบัดกรีวงจรลงแผง Perfboard ในตอนนี้' "
        "เพื่อหลีกเลี่ยงการบัดกรีซ้ำซ้อนและลดความเสียหายต่อแผ่นวงจร โดยอะไหล่มีกำหนดส่งมอบไม่เกินวันที่ 30 ตุลาคม 2569 นี้ "
        "และจะทำการบัดกรีลงบอร์ด Perfboard ในคราวเดียวทันทีที่อุปกรณ์มาถึงครบถ้วน"
    )
    add_thai_body_p(
        doc,
        "ในส่วนของสถาปัตยกรรมภาคจ่ายไฟของวงจรต้นแบบ (Prototype Power Architecture) นั้น คณะผู้จัดทำได้เลือกใช้แหล่งจ่ายไฟตรงแบบสวิตชิ่ง "
        "(Regulated Switching Power Supply 5.2 V) ป้อนเข้าสู่วงจรโดยตรง เพื่อให้ได้ระดับแรงดันอินพุตที่ข้ามพ้นสภาวะ Dropout ของ REF5045 "
        "ซึ่งส่งผลให้วงจรสามารถเริ่มทำงานและอ่านข้อมูลจริงได้ทันทีตามวัตถุประสงค์ของการสร้างต้นแบบขั้นแรก อย่างไรก็ตาม ในเชิงมาตรวิทยาการวัดความละเอียดสูง 24 บิต "
        "สัญญาณรบกวนจากการสลับความถี่ (Switching Ripple / Noise Spikes) เป็นปัจจัยที่อาจส่งผลกระทบต่อระดับสัญญาณรบกวนพื้นฐาน (Noise Floor) "
        "ดังนั้น ในแผนการปรับปรุงฮาร์ดแวร์ระยะถัดไป (Capstone 2) เพื่อยกระดับความเที่ยงตรงสู่มาตรฐานสากล ISO 9060 Class A "
        "คณะผู้จัดทำจึงได้วางแผนการเพิ่มวงจรควบคุมแรงดันตกคร่อมต่ำสัญญาณรบกวนต่ำพิเศษ (Ultra-Low-Noise LDO Post-Regulator เช่น ตระกูล TPS7A หรือ LP5907) "
        "เข้ามาต่อคั่นเพื่อกรองระลอกคลื่นความถี่สูงออกจากภาคจ่ายไฟแอนะล็อกอย่างสมบูรณ์แบบ"
    )

    add_section_heading(doc, "3.3", "ผลการจำลองทางทัศนศาสตร์และการค้นพบความผิดเพี้ยนจากกระจกหน้าต่าง")
    add_thai_body_p(
        doc,
        "คณะผู้จัดทำได้สร้างแบบจำลอง 3 มิติ (CAD Model) ของชิ้นส่วนเซนเซอร์อย่างละเอียดในโปรแกรม CAD ดังแสดงในรูปที่ 3.3 "
        "ประกอบด้วยส่วน Header, โลหะ Can, Glass Window, แผ่นชิป Active Die ขนาด 2.65 x 2.65 mm² และ Cathode Post "
        "จากนั้นนำไปจำลองการเดินทางของลำแสง (Optical Ray-tracing Simulation) ในซอฟต์แวร์จำลองเชิงแสง "
        "เพื่อวิเคราะห์พฤติกรรมเชิงมุมตกกระทบ (0 ถึง 90 องศา) และได้ค้นพบข้อเท็จจริงทางทัศนศาสตร์ที่สำคัญยิ่งดังนี้:"
    )

    # รูปที่ 3.3 โมเดล CAD
    cad1_p = os.path.join(img_dir, "CAD1.png")
    cad2_p = os.path.join(img_dir, "CAD2.png")
    if os.path.exists(cad1_p) and os.path.exists(cad2_p):
        add_two_figure_images_side_by_side(
            doc,
            cad1_p, "(ก) โครงสร้าง 3D CAD ชิ้นส่วนและ Active Die",
            cad2_p, "(ข) มิติเรขาคณิตและมุมมองตัดขวางของตัวถัง",
            fig_no="3.3",
            main_title="ภาพจำลองโมเดล 3D CAD ส่วนประกอบของเซนเซอร์ BPX 61",
            col_width_inch=2.8
        )
    elif os.path.exists(cad1_p):
        add_figure_image(doc, cad1_p, fig_no="3.3", title="ภาพจำลองโมเดล 3D CAD ส่วนประกอบของเซนเซอร์ BPX 61", width_inches=5.0)
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.3",
            title="ภาพจำลองโมเดล 3D CAD ส่วนประกอบของเซนเซอร์ BPX 61",
            description="ภาพแสดงชิ้นส่วน Body_BPX61_Header, Body_BPX61_Can, Body_BPX61_Glass Window, และ Body_BPX61_Active_Die",
            instructions="นำภาพ CAD Model จาก PDF หน้า 2 มาจัดวางแสดงรายละเอียดโครงสร้าง"
        )

    add_subsection_heading(doc, "3.3.1", "การจำลอง BPX 61: ความผิดเพี้ยนจากการหักเหของกระจกหน้าต่าง Schott 8250")
    add_thai_body_p(
        doc,
        "เมื่อทำการจำลองพฤติกรรมเชิงมุมของ BPX 61 เปรียบเทียบกับเส้นกราฟใน Datasheet ของ ams-OSRAM พบว่า (รูปที่ 3.4):"
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

    # แทรกรูปจริง BPX 61 simulation
    img_bpx_path = os.path.join(script_dir, "sim_bpx61_with_without_glass.png")
    if os.path.exists(img_bpx_path):
        add_figure_image(
            doc,
            img_bpx_path,
            fig_no="3.4",
            title="ผลการจำลองการตอบสนองเชิงมุมของ BPX 61: กรณีไม่มีกระจก เทียบกับ กรณีมีกระจก SCHOTT 8250",
            width_inches=5.8
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.4",
            title="ผลการจำลองการตอบสนองเชิงมุมของ BPX 61: กรณีไม่มีกระจก เทียบกับ กรณีมีกระจก",
            description="กราฟเปรียบเทียบ Cut-off ที่ 62° เมื่อไม่มีกระจก เทียบกับ Cut-off ที่ร่นมาเป็น 44° เมื่อใส่กระจก Schott 8250"
        )

    add_subsection_heading(doc, "3.3.2", "การจำลอง BPW 34: การตอบสนองเชิงมุมแบบแบนราบและการประเมิน ISO 9060")
    add_thai_body_p(
        doc,
        "เมื่อเปลี่ยนมาจำลองโฟโตไดโอดแบบแบนราบ BPW 34 ซึ่งใช้ชิปซิลิคอนขนาด 7.02 mm² เดียวกัน ผลการจำลองแสดงข้อได้เปรียบที่เด่นชัด (รูปที่ 3.5):"
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

    # แทรกรูปจริง BPW 34 simulation
    img_bpw_path = os.path.join(script_dir, "sim_bpw34_with_without_glass.png")
    if os.path.exists(img_bpw_path):
        add_figure_image(
            doc,
            img_bpw_path,
            fig_no="3.5",
            title="ผลการจำลองการตอบสนองเชิงมุมและ Heatmap ของ BPW 34 พร้อมการประเมินความคลาดเคลื่อน ISO 9060",
            width_inches=5.8
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.5",
            title="ผลการจำลองการตอบสนองเชิงมุมและ Heatmap ของ BPW 34 พร้อมการประเมิน ISO 9060",
            description="ภาพ Die Heatmap 10 ช่อง และกราฟ ISO 9060 Directional Error เทียบระหว่างมีและไม่มีกระจกแซฟไฟร์"
        )

    add_thai_body_p(
        doc,
        "จากการเปรียบเทียบพฤติกรรมเชิงมุมของทั้ง BPX 61 และ BPW 34 ภายใต้เงื่อนไขที่มีและไม่มีกระจกหน้าต่าง "
        "สามารถสรุปผลการตอบสนองเชิงมุมสัมพัทธ์ (Relative Angular Response) ที่มุมตกกระทบสำคัญต่างๆ ได้ดังแสดงในตารางที่ 3.5"
    )

    # ตารางที่ 3.5
    add_table_caption(doc, "3.5", "การเปรียบเทียบผลการจำลองเชิงมุม: BPX 61 vs BPW 34")
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

    add_subsection_heading(doc, "3.3.3", "ข้อสรุปทางวิศวกรรม: เหตุผลความจำเป็นที่ต้องทำการทดสอบจริงในห้องปฏิบัติการ")
    add_thai_body_p(
        doc,
        "ผลจากการจำลองทางทัศนศาสตร์ชี้ให้เห็นประเด็นสำคัญอย่างยิ่งว่า: 'การเพิ่มกระจกครอบไม่ว่าจะในตัวถัง TO-39 หรือตัวถังแบน "
        "ทำให้พฤติกรรมของแสงในแบบจำลองคอมพิวเตอร์เกิดความคลาดเคลื่อนเพี้ยนไปจากค่าสเปกชีตหรือทฤษฎีอุดมคติเป็นอย่างมาก' "
        "การจำลองทางคณิตศาสตร์มีข้อจำกัดในการจำลองการสะท้อนหลายชั้น (Multiple Internal Reflections) "
        "และการกระเจิงของแสงที่ผิวสัมผัสจริง ความจริงทางวิศวกรรมจึงไม่สามารถสรุปได้ด้วยโปรแกรมจำลองเพียงอย่างเดียว "
        "นี่จึงเป็นเหตุผลสำคัญที่คณะผู้จัดทำจำเป็นต้องออกแบบระเบียบวิธีการทดลองเพื่อนำอุปกรณ์จริงเข้าทดสอบในห้องปฏิบัติการทางทัศนศาสตร์จริง "
        "(Empirical Laboratory Validation) เพื่อให้ได้ข้อมูลจริงที่ถูกต้องที่สุด"
    )

    add_section_heading(doc, "3.4", "การออกแบบข้อเสนอระเบียบวิธีการทดลองจริงทางทัศนศาสตร์")
    add_thai_body_p(
        doc,
        "จากผลการจำลองทางทัศนศาสตร์ในหัวข้อ 3.3 ที่ค้นพบความผิดเพี้ยนของการตอบสนองเชิงมุมเมื่อเพิ่มกระจกหน้าต่าง "
        "คณะผู้จัดทำตระหนักดีว่าการยืนยันพฤติกรรมทางกายภาพที่แท้จริงไม่สามารถสรุปได้ด้วยโปรแกรมจำลองคอมพิวเตอร์เพียงอย่างเดียว "
        "แต่จำเป็นต้องนำชิ้นงานจริงเข้าทดสอบในห้องปฏิบัติการทางทัศนศาสตร์ อย่างไรก็ดี เพื่อให้การดำเนินงานสอดคล้องกับความเป็นไปได้"
        "เชิงปฏิบัติและความพร้อมของเครื่องมือในสถาบัน คณะผู้จัดทำจึงได้จัดทำ 'ข้อเสนอระเบียบวิธีการทดลอง (Proposed Experimental Methodology)' "
        "เพื่อนำเข้าขอรับคำปรึกษาและข้อชี้แนะจากอาจารย์ที่ปรึกษาและคณาจารย์ภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล. "
        "โดยได้วางแผนครอบคลุมทั้งแนวทางในอุดมคติ (Ideal Setup) และแนวทางสำรองเชิงปฏิบัติหรือชุดแสวงเครื่อง (Practical Contingency Setup) "
        "ที่สามารถจัดหาและสร้างขึ้นได้เองอย่างสมเหตุสมผลทางฟิสิกส์ ดังรายละเอียดต่อไปนี้"
    )

    add_subsection_heading(doc, "3.4.1", "ชุดการทดลองที่ 1: การวัดค่าการส่องผ่านสเปกตรัมด้วย Vernier LabQuest 2 Spectrometer (350–900 nm)")
    add_thai_body_p(
        doc,
        "การทดสอบนี้มีวัตถุประสงค์เพื่อตรวจสอบคุณสมบัติการส่องผ่านแสง (Spectral Transmittance: %T) ของแผ่นกระจกแซฟไฟร์สังเคราะห์ "
        "(Synthetic Sapphire Optical Window) ความหนา 1.0 mm เทียบกับอากาศเปล่า (Baseline) ว่ามีค่าการส่องผ่านคงที่และราบเรียบ (Flat) "
        "ครอบคลุมย่านความไวแสงของซิลิคอนไดโอด BPW 34 หรือไม่ โดยตรวจสอบว่ามีระลอกคลื่นการแทรกสอด (Fabry-Pérot Fringes) "
        "หรือการดูดกลืนแสงที่ไม่คาดคิดเกิดขึ้นหรือไม่"
    )
    add_thai_body_p(
        doc,
        "จากการสำรวจเครื่องมือในห้องปฏิบัติการภาควิชาฟิสิกส์ คณะวิทยาศาสตร์ สจล. พบว่ามีชุดตรวจวัดสเปกตรัม 'Vernier LabQuest 2 "
        "Connected Science System' ร่วมกับหัววัดสเปกโตรมิเตอร์ (Vernier Spectrometer พร้อมโพรบใยแก้วนำแสง Optical Fiber) "
        "ซึ่งมีขอบเขตการตรวจวัดครอบคลุมช่วงความยาวคลื่น 350.1 nm ถึง 899.7 nm ดังแสดงในรูปที่ 3.6 โดยครอบคลุมทั้งย่านแสงที่ตามองเห็น (Visible) "
        "และเข้าสู่ย่านอินฟราเรดย่านใกล้ (Near-Infrared: NIR) ซึ่งเป็นช่วงตอบสนองหลักของซิลิคอนไดโอด BPW 34 ได้อย่างสมบูรณ์ "
        "ทำให้ข้อเสนอนี้มีความเป็นไปได้จริง 100% ในเชิงทรัพยากรห้องปฏิบัติการ โดยมีระเบียบวิธีทดสอบดังนี้:"
    )
    add_numbered_item(doc, "1.", "เปิดเครื่อง Vernier LabQuest 2 เชื่อมต่อหัววัดสเปกโตรมิเตอร์และโพรบใยแก้วนำแสง (Optical Fiber Probe)", bold_prefix="การเตรียมเครื่องมือ: ")
    add_numbered_item(doc, "2.", "จัดแนวหัวโพรบใยแก้วนำแสงรับแสงคงที่จากหลอดกำเนิดแสง และทำการบันทึกสเปกตรัมเส้นฐานอ้างอิง I_0(λ) (100% Baseline) ในช่วง 350.1 ถึง 899.7 nm", bold_prefix="การวัดเส้นอ้างอิง (Baseline): ")
    add_numbered_item(doc, "3.", "นำแผ่นกระจกแซฟไฟร์สังเคราะห์มาติดตั้งกั้นขวางหน้าปลายโพรบใยแก้วนำแสงในแนวตั้งฉาก แล้วบันทึกสเปกตรัมแสงที่ส่องผ่าน I(λ)", bold_prefix="การวัดชิ้นงานจริง: ")
    add_numbered_item(doc, "4.", "ประมวลผลคำนวณค่าร้อยละการส่องผ่าน %T(λ) = [I(λ) / I_0(λ)] × 100% และส่งออกไฟล์ข้อมูลตาราง (.csv) ผ่านพอร์ต USB ของ LabQuest 2 เพื่อนำมาพล็อตเปรียบเทียบกับค่า 85–88% ในแบบจำลองทางทัศนศาสตร์", bold_prefix="การประมวลผลข้อมูล: ")

    # แทรกภาพถ่ายจริงของหน้าจอ LabQuest 2
    labquest_img_path = os.path.join(img_dir, "labquest2_vernier_spectrometer.png")
    if not os.path.exists(labquest_img_path):
        labquest_img_path = os.path.join(script_dir, "labquest2_vernier_spectrometer.png")
    if os.path.exists(labquest_img_path):
        add_figure_image(
            doc,
            img_path=labquest_img_path,
            fig_no="3.6",
            title="หน้าจอแสดงผลการวัดสเปกตรัมความยาวคลื่น 350.1–899.7 nm ของชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์",
            width_inches=5.2
        )
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.6",
            title="หน้าจอแสดงผลการวัดสเปกตรัมของชุด Vernier LabQuest 2 ในห้องปฏิบัติการฟิสิกส์",
            description="ภาพถ่ายหน้าจอ LabQuest 2 แสดงกราฟ Intensity vs Wavelength (350.1–899.7 nm)",
            instructions="วางภาพถ่ายเครื่อง Vernier LabQuest 2"
        )

    add_subsection_heading(doc, "3.4.2", "ชุดการทดลองที่ 2: การวัดการตอบสนองเชิงมุมบนแท่นหมุน (Directional Cosine Response Test)")
    add_thai_body_p(
        doc,
        "วัตถุประสงค์เพื่อตรวจสอบพฤติกรรมการตอบสนองต่อมุมตกกระทบของแสง (Cosine Response) ของหัววัด BPW 34 ทั้งกรณีมีและไม่มีกระจกหน้าต่าง "
        "เพื่อพิสูจน์การดรอปของกระแสที่มุมชันเนื่องจากปรากฏการณ์ Fresnel Reflection ดังที่พบในการจำลองทางทัศนศาสตร์ "
        "ระบบการทดสอบได้รับการออกแบบโดยแบ่งเป็นส่วนประกอบหลักดังนี้ (รูปที่ 3.7):"
    )
    add_thai_body_p(
        doc,
        "1) แท่นหมุนปรับมุมองศา (Rotation Stage): อาศัยแท่นหมุนเชิงแสงแบบมือหมุน (Manual Optical Rotation Stage พร้อมสเกล Vernier) "
        "ซึ่งเป็นอุปกรณ์มาตรฐานบนโต๊ะเชิงแสง (Optical Table) ของห้องปฏิบัติการฟิสิกส์ โดยจะจัดทำชุดยึดหัววัด (Test Jig) "
        "เพื่อให้ระนาบผิวรับแสงของชิปไดโอดและกระจกแซฟไฟร์วางอยู่ในแนวแกนหมุนกึ่งกลางพอดี (Co-axial Alignment) "
        "เพื่อป้องกันการเยื้องศูนย์ขณะปรับมุมตั้งแต่ 0° ถึง 85°"
    )
    add_thai_body_p(
        doc,
        "2) การจัดเตรียมแหล่งกำเนิดแสงและข้อเสนอลำแสงขนาน (Light Source & Collimation Strategy): คณะผู้จัดทำได้วางแนวทางไว้ 2 รูปแบบ "
        "เพื่อนำเสนอขอรับความเห็นชอบจากอาจารย์ที่ปรึกษาตามความเหมาะสมของอุปกรณ์ในภาควิชา:"
    )
    add_bullet_item(doc, bullet_char="•", text="แนวทางตามอุดมคติ (Ideal Scenario): หากห้องปฏิบัติการมีชุดเครื่องจำลองแสงอาทิตย์ (Solar Simulator) หรือชุดแหล่งกำเนิดแสงซีนอนร่วมกับเลนส์คอลลิเมเตอร์ (Collimating Lens System) พร้อมใช้งาน จะใช้ลำแสงขนานมาตรฐานนี้ส่องเข้าหาแท่นหมุนโดยตรง")
    add_bullet_item(doc, bullet_char="•", text="แนวทางสำรองเชิงปฏิบัติหรือชุดแสวงเครื่อง (Practical Contingency / DIY Setup): ในกรณีที่ไม่มีชุดกำเนิดลำแสงขนานเฉพาะทาง คณะผู้จัดทำเสนอให้ใช้ 'หลอดทังสเตน-ฮาโลเจน (Tungsten-Halogen Lamp ขนาด 12V 50W หรือ 100W)' จ่ายไฟด้วย Bench Power Supply กระแสตรงคงที่ เพื่อให้ความเข้มแสงนิ่งสนิท ไร้การกระพริบตามความถี่ไฟบ้าน 50 Hz โดยมีข้อได้เปรียบทางฟิสิกส์คือ หลอดทังสเตนเปล่งแสงแบบการแผ่รังสีวัตถุดำ (Blackbody Radiation ที่อุณหภูมิ 2,800–3,000 K) ซึ่งให้สเปกตรัมที่ต่อเนื่องราบเรียบตั้งแต่ช่วงคลื่นตามองเห็นจนถึงอินฟราเรดย่านใกล้ (350–2,000+ nm) สอดคล้องกับย่านการทำงานของซิลิคอนไดโอดได้เป็นอย่างดี ปราศจากเส้นสเปกตรัมแหลมแบบหลอดฟลูออเรสเซนต์หรือ LED")
    add_bullet_item(doc, bullet_char="•", text="การสร้างลำแสงขนานเสมือนด้วยระยะห่าง (Quasi-parallel Beam via Distance): คณะผู้จัดทำประยุกต์ใช้หลักการทางเรขาคณิต โดยติดตั้งหลอดทังสเตนไว้ที่ระยะห่าง d ≥ 1.5 ถึง 2.0 เมตรจากแท่นหมุน ด้วยขนาดชิป BPW 34 ที่เล็กเพียง 2.65 × 2.65 mm มุมกวาดเบี่ยงเบนของลำแสง (Subtended Divergence Angle) จะมีค่าเพียง arctan(2.65 / 2000) ≈ 0.076° ซึ่งน้อยกว่า 0.1° มาก ทำให้ลำแสงที่เดินทางมาตกกระทบหน้าสัมผัสเซนเซอร์มีความขนานเสมือนจริงตามทฤษฎี โดยไม่จำเป็นต้องพึ่งพาเลนส์รวมแสงราคาแพง")
    add_bullet_item(doc, bullet_char="•", text="การป้องกันแสงสะท้อนและแสงกระเจิง (Stray Light & Baffle System): ทำการทดสอบในห้องมืดสนิท และติดตั้งท่อบังแสงทรงกระบอก (Baffle Tube) บุผิวด้านในด้วยวัสดุสีดำด้านหน้าหลอดไฟและเซนเซอร์ เพื่อกำจัดแสงสะท้อนจากผนังห้องหรือพื้นโต๊ะทดลองไม่ให้เล็ดลอดเข้ามาส่งผลต่อการวัดที่มุมเฉียง")

    # รูปที่ 3.7 แบบจำลอง Blender สำหรับชุดแท่นหมุน Goniometer
    blender_persp = os.path.join(img_dir, "goniometer_blender_perspective.png")
    blender_top = os.path.join(img_dir, "goniometer_blender_topdown.jpg")
    if os.path.exists(blender_persp) and os.path.exists(blender_top):
        add_two_figure_images_stacked(
            doc,
            blender_persp, "(ก) แบบจำลองสามมิติ (Perspective View) แสดงการจัดวางโคมหลอดทังสเตน ท่อ Baffle และแท่นหมุนบน Optical Table",
            blender_top, "(ข) แบบจำลองมุมมองระนาบบน (Top-down Plan View) แสดงแนวลำแสงขนานและการวัดมุมตกกระทบ θ",
            fig_no="3.7",
            main_title="แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุม (Goniometer Cosine Test Jig) พร้อมข้อเสนอชุดกำเนิดแสงแสวงเครื่อง",
            width_inches=5.6
        )
    elif os.path.exists(blender_persp):
        add_figure_image(doc, blender_persp, fig_no="3.7", title="แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุม (Goniometer Cosine Test Jig)", width_inches=5.6)
    else:
        add_photo_placeholder(
            doc,
            fig_no="3.7",
            title="แบบจำลองชุดแท่นหมุนวัดการตอบสนองเชิงมุม (Goniometer Cosine Test Jig) พร้อมข้อเสนอชุดกำเนิดแสงแสวงเครื่อง",
            description="ภาพไดอะแกรมแสดงแท่นหมุน Co-axial การจัดวางหลอดทังสเตนระยะ 2 เมตร และระบบท่อ Baffle ป้องกันแสงสะท้อน"
        )

    add_thai_body_p(
        doc,
        "ข้อเสนอระเบียบวิธีการทดลองทั้งสองชุดนี้ ได้รับการออกแบบขึ้นเพื่อเป็นกรอบการดำเนินงานอย่างรอบคอบ "
        "โดยคณะผู้จัดทำจะนำแผนงานดังกล่าวเข้าปรึกษาอาจารย์ที่ปรึกษาและอาจารย์ประจำห้องปฏิบัติการฟิสิกส์ เพื่อร่วมกันประเมิน "
        "ความพร้อมของเครื่องมือ และเลือกแนวทางการทดลองที่เหมาะสมและเป็นไปได้จริงมากที่สุดสำหรับการดำเนินงานโครงงานในภาคการศึกษาถัดไป"
    )

    add_section_heading(doc, "3.5", "สถานะปัจจุบันและการออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom 4-Layer PCB) สู่การผลิตจริง")
    add_thai_body_p(
        doc,
        "สรุปสถานะความก้าวหน้าและการพัฒนาฮาร์ดแวร์ของโครงงาน Capstone Project 1 ณ ปัจจุบัน:"
    )
    add_numbered_item(doc, "1.", "วงจรภาคอ่านสัญญาณได้รับการพิสูจน์การทำงานจริงบน Breadboard และตัวต้นแบบบัดกรีมือเรียบร้อยแล้ว สามารถอ่านสัญญาณ ADC 24 บิตและตอบสนองต่อสเปกตรัมแสงต่างๆ ได้อย่างแม่นยำ พร้อมทั้งค้นพบปัญหา Parasitic Trace Offset และ USB Dropout ซึ่งได้รับการแก้ไขด้วยสวิตชิ่ง 5.2 V", bold_prefix="การพิสูจน์วงจรต้นแบบ: ")
    add_numbered_item(doc, "2.", "ค้นพบปัญหาความไม่เสถียรของ REF5045 อันเนื่องมาจากค่า ESR ของตัวเก็บประจุ และได้วางแผนแก้ปัญหาอย่างถูกต้องโดยการสั่งซื้อ Low-ESR Ceramic Cap (10 µF X7R) เพื่อรักษา Phase Margin ของลูปควบคุมแรงดัน", bold_prefix="การแก้ไขด้านความเสถียร: ")
    add_numbered_item(doc, "3.", "ผลการจำลองทัศนศาสตร์พิสูจน์ชัดเจนว่า BPW 34 ให้การตอบสนองเชิงมุมเหนือกว่า BPX 61 และค้นพบประเด็นความผิดเพี้ยนจากการใส่กระจกหน้าต่าง ซึ่งเป็นที่มาของการออกแบบระเบียบวิธีวิจัยเพื่อนำเข้าทดสอบจริงในห้องปฏิบัติการของคณะวิทยาศาสตร์ในโครงงาน 2", bold_prefix="กรอบการวิจัยทางแสงสู่ Capstone 2: ")
    add_numbered_item(doc, "4.", "เพื่อขจัดปัญหาความต้านทานแฝงและสัญญาณรบกวนจากการต่อสายลอย คณะผู้จัดทำได้ออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์ (Custom 4-Layer Circular Sensor Head PCB) รุ่น 'V1.1 L1 by Tachin' บนโปรแกรม KiCad 10.0 โดยจัดสรรโครงสร้างแบบคู่ขนาน 2 รูปแบบ ได้แก่ รุ่น L1_Analog (สำหรับ BPX 61 ในตัวถัง TO-39 Can) และรุ่น L1_Analog_BPW34 (สำหรับ BPW 34 ในตัวถัง Flat DIL Plastic) ทั้งนี้ จากการตรวจสอบเอกสารข้อมูลทางเทคนิคของ ams OSRAM ยืนยันว่าเซนเซอร์ทั้งสองรุ่นใช้เนื้อชิปซิลิคอน (Active Chip Die) ขนาดเดียวกันเป๊ะ คือ 2.65 × 2.65 mm² (พื้นที่รับแสง 7.02 mm²) จึงมีคุณสมบัติทางไฟฟ้าและสเปกตรัมเท่ากันทุกประการ ทำให้การออกแบบแผ่นวงจรคู่ขนานนี้เป็นการทดลองแบบควบคุมตัวแปร (Controlled Experiment) ที่สมบูรณ์แบบเพื่อเปรียบเทียบผลของตัวถังและหน้าต่างทางทัศนศาสตร์อย่างแท้จริง โดยทั้งสองรุ่นผ่านการตรวจ DRC 100% ไร้ข้อผิดพลาด (0 Violations / 0 Unconnected items) พร้อมส่งผลิตโรงงาน (PCBWay) สำหรับใช้ใน Capstone 2 ต่อไป", bold_prefix="การออกแบบแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน (Custom Dual-Variant PCB): ")

    # ภาพเรนเดอร์ 3D PCB V1.1 L1 (BPX 61 vs BPW 34)
    p_bpx = os.path.join(img_dir, "pcb_3d_render_iso.png")
    p_bpw = os.path.join(img_dir, "pcb_bpw34_3d_render_iso.png")
    if os.path.exists(p_bpx) and os.path.exists(p_bpw):
        add_two_figure_images_side_by_side(
            doc,
            p_bpx, "(ก) รุ่นที่ 1: ตัวถังโลหะ TO-39 Can (BPX 61)",
            p_bpw, "(ข) รุ่นที่ 2: ตัวถังพลาสติกแบน Flat DIL (BPW 34)",
            fig_no="3.13",
            main_title="ภาพจำลอง 3 มิติของแผ่นวงจรพิมพ์ 4 เลเยอร์คู่ขนาน (Custom 4-Layer PCB: BPX 61 vs BPW 34)",
            col_width_inch=2.85
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

    references = [
        "International Organization for Standardization. (2018). Solar energy — Specification and classification of instruments for measuring hemispherical solar and direct solar radiation (ISO Standard No. 9060:2018).",
        "ams-OSRAM AG. (2023). BPX 61 Silicon PIN Photodiode in Metal Can TO39 Package [Datasheet]. Version 1.6, Document No. Q62705P0025.",
        "ams-OSRAM AG. (2023). BPW 34 Silicon PIN Photodiode in Flat Clear Plastic Package [Datasheet]. Version 1.6.",
        "Texas Instruments. (2016). ADS1220 4-Channel, 2-kSPS, Low-Power, 24-Bit ADC with Integrated PGA and Reference [Datasheet]. SBAS501B.",
        "Texas Instruments. (2018). REF50xx Low-Noise, Very Low Drift, Precision Voltage Reference [Datasheet]. SBOS410I.",
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

    add_section_heading(doc, "ภาคผนวก ก", "รายการอุปกรณ์ต้นแบบ (Bill of Materials - BOM)")
    add_thai_body_p(
        doc,
        "รายการอุปกรณ์และชิ้นส่วนอิเล็กทรอนิกส์ทั้งหมดที่ใช้ในการทดสอบวงจรต้นแบบบน Breadboard "
        "และเตรียมประกอบลงบนแผงวงจรอเนกประสงค์ (Perfboard) สำหรับโครงงาน Project 1:"
    )

    t_bom = doc.add_table(rows=11, cols=4)
    t_bom.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(t_bom, border_color="000000", border_sz="4")
    bom_headers = ["อุปกรณ์ / ไอซี", "เบอร์อุปกรณ์ / Part Number", "คุณสมบัติสำคัญ", "สถานะการใช้งาน"]
    for j, h in enumerate(bom_headers):
        cell = t_bom.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    bom_data = [
        ["Silicon Photodiode", "OSRAM BPW 34", "Flat Clear Mold, 7.02 mm² Active Area", "ใช้งานจริงบน Breadboard"],
        ["Alternative Detector", "OSRAM BPX 61", "TO-39 Metal Can, แผ่นชิป 7.02 mm²", "จัดหาแล้วสำหรับทดสอบเปรียบเทียบ"],
        ["Optical Window", "Synthetic Sapphire", "ความแข็งระดับ 9 Mohs, หนา 1 mm", "พร้อมสำหรับการทดสอบแล็บ"],
        ["Transimpedance Op-Amp", "TI OPA380AID", "Zero-Drift, Offset ต่ำมาก, R_TIA = 2.2 kΩ", "วงจรขยายกระแสแสง TIA"],
        ["24-bit Delta-Sigma ADC", "TI ADS1220IPW", "TSSOP-16 พร้อมบอร์ดแปลง, 4-ch MUX, Low-noise PGA", "ใช้งานจริงบน Breadboard"],
        ["Precision Reference", "TI REF5045AID", "SOIC-8 พร้อมบอร์ดแปลง, 4.500 V, 3 ppm/°C", "ใช้งานจริงบน Breadboard"],
        ["Microcontroller Unit", "ESP32-C6 Super Mini", "RISC-V 32-bit, Wi-Fi 6, SPI Master, Web Server", "ใช้งานจริงบน Breadboard"],
        ["NTC Thermistor", "Thinking TTC3A103F39HAEY", "10 kΩ ±1% ที่ 25°C, ค่า B25/85 = 3950 K", "ใช้งานจริงบน Breadboard (จัดหาในไทย)"],
        ["Decoupling Capacitor C10", "10 µF 16V X7R", "Low-ESR Ceramic Capacitor", "สั่งซื้อแล้ว (กำหนดส่งมอบ 30 ต.ค.)"],
        ["Regulated Power Supply", "DC Switching 5.2V", "แรงดันคงที่ 5.20 V, ก้าวข้าม Dropout ของ REF5045", "ใช้งานจริงบน Breadboard"]
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
