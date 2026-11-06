import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

sys.path.append(os.path.dirname(__file__))
from helpers import (
    set_cell_shading, set_cell_margins, add_photo_placeholder,
    add_heading_1, add_heading_2, add_heading_3, add_body_p, add_bullet_p
)

def create_report():
    doc = docx.Document()

    # Page setup (A4 standard: Top 1.5 in, Left 1.5 in, Right 1.0 in, Bottom 1.0 in per KMITL thesis guidelines)
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.5)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.5)
    section.right_margin = Inches(1.0)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'TH SarabunPSK'
    normal_style.font.size = Pt(15)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)

    # =========================================================================
    # 1. THAI COVER PAGE (หน้าปกภาษาไทย)
    # =========================================================================
    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(40)

    p_th_title = doc.add_paragraph()
    p_th_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_th = p_th_title.add_run("การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง\n")
    r_th.font.name = 'TH SarabunPSK'
    r_th.font.size = Pt(22)
    r_th.font.bold = True

    p_en_title = doc.add_paragraph()
    p_en_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_title.paragraph_format.space_after = Pt(80)
    r_en = p_en_title.add_run("DEVELOPMENT OF A HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM\n")
    r_en.font.name = 'TH SarabunPSK'
    r_en.font.size = Pt(20)
    r_en.font.bold = True

    p_authors = doc.add_paragraph()
    p_authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_authors.paragraph_format.space_after = Pt(90)
    r_auth = p_authors.add_run(
        "ทชินทร์ รักษาวงศ์          รหัสนักศึกษา 6601xxxx\n"
        "(ชื่อนักศึกษาคนที่ 2)       รหัสนักศึกษา 6601xxxx\n"
    )
    r_auth.font.name = 'TH SarabunPSK'
    r_auth.font.size = Pt(16)
    r_auth.font.bold = True

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_footer.add_run(
        "รายงานความก้าวหน้าปริญญานิพนธ์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตร\n"
        "ปริญญาวิศวกรรมศาสตรบัณฑิต สาขาวิชาวิศวกรรมระบบไอโอทีและสารสนเทศ\n"
        "และปริญญาวิทยาศาสตรบัณฑิต สาขาวิชาฟิสิกส์อุตสาหกรรม\n"
        "ภาควิชาวิศวกรรมระบบไอโอทีและสารสนเทศ\n"
        "คณะวิศวกรรมศาสตร์\n"
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง\n"
        "ปีการศึกษา 2569"
    )
    r_foot.font.name = 'TH SarabunPSK'
    r_foot.font.size = Pt(15)

    doc.add_page_break()

    # =========================================================================
    # 2. ENGLISH COVER PAGE (หน้าปกภาษาอังกฤษ)
    # =========================================================================
    p_sp2 = doc.add_paragraph()
    p_sp2.paragraph_format.space_before = Pt(45)

    p_en_cov = doc.add_paragraph()
    p_en_cov.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_cov.paragraph_format.space_after = Pt(90)
    r_encov = p_en_cov.add_run(
        "DEVELOPMENT OF A HIGH-PRECISION SOLAR IRRADIANCE\nMEASUREMENT SYSTEM\n"
    )
    r_encov.font.name = 'TH SarabunPSK'
    r_encov.font.size = Pt(22)
    r_encov.font.bold = True

    p_auth_en = doc.add_paragraph()
    p_auth_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth_en.paragraph_format.space_after = Pt(90)
    r_authen = p_auth_en.add_run(
        "TACHIN RAKSAWONG          STUDENT ID 6601xxxx\n"
        "CO-AUTHOR NAME            STUDENT ID 6601xxxx\n"
    )
    r_authen.font.name = 'TH SarabunPSK'
    r_authen.font.size = Pt(16)
    r_authen.font.bold = True

    p_en_sub = doc.add_paragraph()
    p_en_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ensub = p_en_sub.add_run(
        "A CAPSTONE PROJECT 1 PROGRESS REPORT SUBMITTED IN PARTIAL FULFILLMENT\n"
        "OF THE REQUIREMENTS FOR THE DUAL DEGREE OF\n"
        "BACHELOR OF ENGINEERING IN IOT SYSTEM AND INFORMATION ENGINEERING AND\n"
        "BACHELOR OF SCIENCE IN INDUSTRIAL PHYSICS\n"
        "DEPARTMENT OF IoT AND INFORMATION ENGINEERING\n"
        "SCHOOL OF ENGINEERING\n"
        "KING MONGKUT'S INSTITUTE OF TECHNOLOGY LADKRABANG\n"
        "ACADEMIC YEAR 2026"
    )
    r_ensub.font.name = 'TH SarabunPSK'
    r_ensub.font.size = Pt(15)

    doc.add_page_break()

    # =========================================================================
    # 3. APPROVAL SHEET (หน้าอนุมัติปริญญานิพนธ์)
    # =========================================================================
    p_app_head = doc.add_paragraph()
    p_app_head.paragraph_format.space_before = Pt(10)
    p_app_head.paragraph_format.space_after = Pt(16)
    
    t_app_info = doc.add_table(rows=7, cols=2)
    t_app_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_app_info.autofit = False
    col_widths = [Inches(1.8), Inches(4.5)]
    for row in t_app_info.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    app_data = [
        ("หัวข้อปริญญานิพนธ์", "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง"),
        ("Thesis Title", "Development of a High-Precision Solar Irradiance Measurement System"),
        ("ชื่อนักศึกษา", "นายทชินทร์ รักษาวงศ์   รหัสนักศึกษา 6601xxxx\n(ชื่อนักศึกษาคนที่ 2)  รหัสนักศึกษา 6601xxxx"),
        ("ระดับปริญญา", "วิศวกรรมศาสตรบัณฑิตและวิทยาศาสตรบัณฑิต (หลักสูตรสองปริญญา)"),
        ("สาขาวิชา", "วิศวกรรมระบบไอโอทีและสารสนเทศ และฟิสิกส์อุตสาหกรรม"),
        ("ปีการศึกษา", "2569"),
        ("อาจารย์ที่ปรึกษา", "รศ.ดร. (ระบุชื่ออาจารย์ที่ปรึกษาหลัก)\nอาจารย์ที่ปรึกษาร่วม (ถ้ามี)")
    ]

    for idx, (label, val) in enumerate(app_data):
        c0 = t_app_info.cell(idx, 0)
        c1 = t_app_info.cell(idx, 1)
        r0 = c0.paragraphs[0].add_run(label)
        r0.font.name = 'TH SarabunPSK'
        r0.font.size = Pt(15)
        r0.font.bold = True
        r1 = c1.paragraphs[0].add_run(val)
        r1.font.name = 'TH SarabunPSK'
        r1.font.size = Pt(15)

    p_app_body = doc.add_paragraph()
    p_app_body.paragraph_format.space_before = Pt(16)
    p_app_body.paragraph_format.space_after = Pt(16)
    r_app_txt = p_app_body.add_run(
        "คณะวิศวกรรมศาสตร์ สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง อนุมัติให้รายงานความก้าวหน้านี้เป็นส่วนหนึ่งของรายวิชา 01236269-IoT AND INFORMATION CAPSTONE PROJECT 1 ตามหลักสูตรปริญญาวิศวกรรมศาสตรบัณฑิต (วิศวกรรมระบบไอโอทีและสารสนเทศ) และวิทยาศาสตรบัณฑิต (ฟิสิกส์อุตสาหกรรม) ประจำปีการศึกษา 2569"
    )
    r_app_txt.font.name = 'TH SarabunPSK'
    r_app_txt.font.size = Pt(15)

    # Committee Table
    t_comm = doc.add_table(rows=6, cols=2)
    t_comm.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_comm.autofit = False
    for r in t_comm.rows:
        r.cells[0].width = Inches(3.5)
        r.cells[1].width = Inches(2.8)

    c_th0 = t_comm.cell(0, 0)
    c_th1 = t_comm.cell(0, 1)
    set_cell_shading(c_th0, "E2E8F0")
    set_cell_shading(c_th1, "E2E8F0")
    c_th0.paragraphs[0].add_run("คณะกรรมการสอบ").font.bold = True
    c_th1.paragraphs[0].add_run("ลายมือชื่อ").font.bold = True
    c_th0.paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    c_th1.paragraphs[0].runs[0].font.name = 'TH SarabunPSK'

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
        r_role = cell_role.paragraphs[0].add_run(role)
        r_role.font.name = 'TH SarabunPSK'
        r_role.font.size = Pt(15)
        cell_role.paragraphs[0].paragraph_format.space_before = Pt(8)
        cell_role.paragraphs[0].paragraph_format.space_after = Pt(8)
        r_sig = cell_sig.paragraphs[0].add_run("....................................................")
        r_sig.font.name = 'TH SarabunPSK'
        r_sig.font.size = Pt(15)
        cell_sig.paragraphs[0].paragraph_format.space_before = Pt(8)
        cell_sig.paragraphs[0].paragraph_format.space_after = Pt(8)

    doc.add_page_break()

    # =========================================================================
    # 4. ABSTRACT THAI (บทคัดย่อภาษาไทย) - หน้า I
    # =========================================================================
    p_abs_title = doc.add_paragraph()
    p_abs_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_abs_title.paragraph_format.space_before = Pt(10)
    p_abs_title.paragraph_format.space_after = Pt(12)
    r_abst = p_abs_title.add_run("บทคัดย่อ")
    r_abst.font.name = 'TH SarabunPSK'
    r_abst.font.size = Pt(18)
    r_abst.font.bold = True

    add_body_p(
        doc,
        "รายงานความก้าวหน้าปริญญานิพนธ์ฉบับนี้นำเสนอการออกแบบและพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (High-Precision Solar Irradiance Measurement System) "
        "โดยมุ่งเน้นการแก้ปัญหาความคลาดเคลื่อนทางกายภาพของเซนเซอร์ตรวจวัดแสงด้วยวิธีการคำนวณเชิงอัลกอริทึม (Computational Sensing & Algorithmic Calibration) "
        "ระบบประกอบด้วยภาคปรับสภาพสัญญาณแอนะล็อกความไวสูง (L1_Analog Front-End) โดยใช้วงจรขยาย Transimpedance ร่วมกับไอซีแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต "
        "(Texas Instruments ADS1220), แหล่งจ่ายแรงดันอ้างอิงความเสถียรสูง (Texas Instruments REF5045) และหัววัดอุณหภูมิ NTC Thermistor 10 kΩ (Semitec 103AT) "
        "เพื่อใช้ในการชดเชยการเปลี่ยนแปลงค่าสัมประสิทธิ์การตอบสนองต่ออุณหภูมิ (Temperature Coefficient of Responsivity) ของหัววัดซิลิกอนโฟโตไดโอด (BPX61)"
    )
    add_body_p(
        doc,
        "ในการศึกษาความคลาดเคลื่อนเชิงมุมตามกฎของโคไซน์ (Cosine Error) สมมติฐานเริ่มต้นของโครงงานได้เลือกใช้เซนเซอร์ BPX61 ตัวถัง TO-39 ร่วมกับแนวคิดการขึ้นรูปโดมกระจายแสง PTFE (Teflon) "
        "อย่างไรก็ตาม จากการจำลองเชิงแสง (Ray-tracing Optical Simulation) และการวิเคราะห์คุณสมบัติวัสดุ พบว่าแบบจำลองการกระเจิงแสงของ PTFE มีพารามิเตอร์ซับซ้อนสูงยิ่ง "
        "(Bulk Scattering Coefficient, Absorption Coefficient และ Henyey-Greenstein Anisotropy Factor) อีกทั้งยังมีข้อจำกัด Trade-off ระหว่างความหนาและการสูญเสียสัญญาณ "
        "ประกอบกับตัวถัง TO-39 มีขอบกระป๋องโลหะยกสูงทำให้เกิดการบดบังผิวรับแสงภายใน (Aperture Cavity Shadowing) ที่มุมเกิน 40 องศา และตัดแสงสิ้นเชิงที่ 70 องศา "
        "การขึ้นรูปโดม PTFE เองจึงมีความเสี่ยงสูงต่อคุณภาพและความเสร็จสมบูรณ์ของโครงงาน"
    )
    add_body_p(
        doc,
        "คณะผู้จัดทำจึงได้ดำเนินกลยุทธ์แก้ปัญหา 2 ระยะอย่างเป็นระบบ: ในโครงงานที่ 1 (Project 1) แก้ปัญหาเฉพาะหน้าด้วยการพัฒนาซอฟต์แวร์ชดเชยค่าเชิงคณิตศาสตร์ (Software Cosine Correction) "
        "ผ่านการคำนวณตำแหน่งดวงอาทิตย์ (Solar Position Algorithm) อิงพิกัด GPS และเวลาจริงเพื่อหักล้างเงาของกระป๋อง พร้อมทั้งประกอบและทดสอบวงจรต้นแบบจริงบนแผงวงจรอเนกประสงค์ (Perfboard) "
        "และทดสอบการสื่อสารข้อมูลเบื้องต้น และในโครงงานที่ 2 (Project 2) ได้วางแผนการแก้ปัญหาที่ต้นตอกายภาพด้วยการย้ายแพ็กเกจเซนเซอร์ไปสู่ Flat SMD (Vishay TEMD5010X01) บน Custom PCB 4 เลเยอร์ "
        "พร้อมทั้งดำเนินงานทดสอบเปรียบเทียบ A/B Testing ในห้องปฏิบัติการทัศนศาสตร์ (Optics Lab) และเครื่องจำลองแสงอาทิตย์ Solar Light XPS-300 ต่อไป"
    )

    doc.add_page_break()

    # =========================================================================
    # 5. ABSTRACT ENGLISH - หน้า II
    # =========================================================================
    p_absen_title = doc.add_paragraph()
    p_absen_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_absen_title.paragraph_format.space_before = Pt(10)
    p_absen_title.paragraph_format.space_after = Pt(12)
    r_absen = p_absen_title.add_run("ABSTRACT")
    r_absen.font.name = 'TH SarabunPSK'
    r_absen.font.size = Pt(18)
    r_absen.font.bold = True

    add_body_p(
        doc,
        "This progress report presents the design and development of a high-precision solar irradiance measurement system, "
        "combining precision analog instrumentation with computational sensing and algorithmic calibration. The analog front-end incorporates "
        "a 24-bit delta-sigma analog-to-digital converter (ADS1220), an ultra-low-drift voltage reference (REF5045, 3 ppm/°C), and a 10 kΩ NTC thermistor (Semitec 103AT) "
        "for real-time thermal compensation of the silicon photodiode responsivity drift."
    )
    add_body_p(
        doc,
        "The project initially selected the BPX61 photodiode (TO-39 metal can) with the intention of fabricating a custom PTFE (Teflon) diffuser dome. "
        "However, rigorous optical ray-tracing simulations and material scattering modeling revealed extreme parameter complexities, including bulk scattering coefficients, "
        "absorption trade-offs, and critical internal aperture vignetting caused by the raised rim of the TO-39 package at incident angles θ > 40°, leading to total optical cutoff at 70°. "
        "Attempting manual PTFE sintering was identified as an engineering risk that could jeopardize prototype repeatability and project timelines."
    )
    add_body_p(
        doc,
        "Consequently, a two-phase pragmatic engineering roadmap was executed: In Project 1, circuit integrity and stability were proven via a fully functional perfboard prototype, "
        "and an algorithmic Software Cosine Correction was established using the Solar Position Algorithm (SPA) with GPS coordinates and timestamps. "
        "In Project 2, the root-cause physical limitation will be resolved by migrating to an ultra-thin Flat SMD photodiode (Vishay TEMD5010X01, 1.12 mm height) on a custom 4-layer PCB, "
        "accompanied by comprehensive A/B comparative testing using the KMITL Physics Optics Lab goniometer bench and the Solar Light XPS-300 Solar Simulator."
    )

    doc.add_page_break()

    # =========================================================================
    # 6. ACKNOWLEDGEMENTS (กิตติกรรมประกาศ) - หน้า III
    # =========================================================================
    p_ack_title = doc.add_paragraph()
    p_ack_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ack_title.paragraph_format.space_before = Pt(10)
    p_ack_title.paragraph_format.space_after = Pt(12)
    r_ack = p_ack_title.add_run("กิตติกรรมประกาศ")
    r_ack.font.name = 'TH SarabunPSK'
    r_ack.font.size = Pt(18)
    r_ack.font.bold = True

    add_body_p(
        doc,
        "รายงานความก้าวหน้าปริญญานิพนธ์ฉบับนี้สำเร็จลุล่วงได้ด้วยความกรุณาและการสนับสนุนเป็นอย่างดียิ่งจากอาจารย์ที่ปรึกษาโครงงาน "
        "ที่ได้กรุณาสละเวลาให้คำปรึกษา แนะนำแนวทางการวิเคราะห์ปัญหาเชิงวิศวกรรม ตลอดจนการถ่ายทอดองค์ความรู้ด้านการออกแบบฮาร์ดแวร์สัญญาณแอนะล็อกความแม่นยำสูง "
        "การจำลองคุณลักษณะทางทัศนศาสตร์ และระเบียบวิธีทางมาตรวิทยาอันเป็นประโยชน์อย่างยิ่งต่อการดำเนินโครงงาน"
    )
    add_body_p(
        doc,
        "ขอขอบพระคุณคณาจารย์ประจำภาควิชาวิศวกรรมระบบไอโอทีและสารสนเทศ และสาขาวิชาฟิสิกส์อุตสาหกรรม คณะวิศวกรรมศาสตร์ "
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง สำหรับการประสิทธิ์ประสาทวิชาความรู้ และขอขอบคุณเจ้าหน้าที่ห้องปฏิบัติการทัศนศาสตร์ (Optics Lab) "
        "และศูนย์เครื่องมือวิทยาศาสตร์และเทคโนโลยี (CMIT) สำหรับการให้คำปรึกษาด้านการเข้าใช้เครื่องจำลองแสงอาทิตย์ (Solar Simulator XPS-300)"
    )
    add_body_p(
        doc,
        "สุดท้ายนี้ ขอขอบคุณเพื่อนร่วมทีมโครงงานสำหรับการทุ่มเททำงานร่วมกันอย่างเป็นระบบ และขอขอบพระคุณครอบครัวที่ให้กำลังใจและสนับสนุนการศึกษาตลอดมา"
    )

    doc.add_page_break()

    # =========================================================================
    # 7. TABLE OF CONTENTS (สารบัญ) - หน้า IV
    # =========================================================================
    p_toc_title = doc.add_paragraph()
    p_toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_title.paragraph_format.space_after = Pt(12)
    r_toct = p_toc_title.add_run("สารบัญ")
    r_toct.font.name = 'TH SarabunPSK'
    r_toct.font.size = Pt(18)
    r_toct.font.bold = True

    t_toc = doc.add_table(rows=1, cols=2)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_toc.rows[0].cells[0].paragraphs[0].add_run("เรื่อง").font.bold = True
    t_toc.rows[0].cells[1].paragraphs[0].add_run("หน้า").font.bold = True
    t_toc.rows[0].cells[0].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_toc.rows[0].cells[1].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_toc.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    toc_items = [
        ("บทคัดย่อภาษาไทย", "I"),
        ("บทคัดย่อภาษาอังกฤษ", "II"),
        ("กิตติกรรมประกาศ", "III"),
        ("สารบัญ", "IV"),
        ("สารบัญตาราง", "VI"),
        ("สารบัญรูป", "VII"),
        ("บทที่ 1 บทนำ", "1"),
        ("    1.1 ความเป็นมาและความสำคัญของปัญหา", "1"),
        ("    1.2 วัตถุประสงค์ของโครงงาน", "2"),
        ("    1.3 ขอบเขตของโครงงาน (Project 1 และ Project 2)", "3"),
        ("    1.4 การแบ่งบทบาทและความรับผิดชอบในทีม (Division of Responsibilities)", "4"),
        ("    1.5 ประโยชน์ที่คาดว่าจะได้รับ", "5"),
        ("    1.6 แผนการดำเนินงานตลอดโครงงาน", "5"),
        ("บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "7"),
        ("    2.1 หลักการตรวจวัดการแผ่รังสีแสงอาทิตย์และกฎของโคไซน์ (Cosine Law)", "7"),
        ("    2.2 คุณสมบัติของ Silicon Photodiode และผลกระทบของอุณหภูมิ", "9"),
        ("    2.3 วงจรแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)", "12"),
        ("    2.4 วงจรแรงดันอ้างอิงความเสถียรสูง (REF5045) และปัญหา High-Z Leakage", "14"),
        ("    2.5 ข้อจำกัดของวัสดุ PTFE และการเกิดเงาของตัวถัง TO-39 (Cosine Simulation)", "16"),
        ("    2.6 หลักการมาตรวิทยา การสอบเทียบ และห่วงโซ่การสอบกลับได้ (Traceability)", "19"),
        ("บทที่ 3 การออกแบบระบบและการดำเนินงาน (System Design & Implementation)", "23"),
        ("    3.1 ภาพรวมสถาปัตยกรรมระบบแบบ Computational Sensing", "23"),
        ("    3.2 การออกแบบฮาร์ดแวร์ต้นแบบ (Hardware Prototype Implementation)", "25"),
        ("        3.2.1 วงจรแอนะล็อกส่วนหน้า (L1_Analog Front-End)", "25"),
        ("        3.2.2 การสร้างและแก้ปัญหาบน Breadboard สู่การประกอบลง Perfboard", "27"),
        ("        3.2.3 บอร์ดประมวลผลและการส่งข้อมูล (L2_MCU Board)", "30"),
        ("        3.2.4 บอร์ดจัดการพลังงานลดสัญญาณรบกวน (L3_Switching Board)", "31"),
        ("    3.3 สถาปัตยกรรมซอฟต์แวร์และอัลกอริทึมการชดเชยค่า (Software Calibration)", "33"),
        ("        3.3.1 ผังการทำงานของเฟิร์มแวร์ไมโครคอนโทรลเลอร์", "33"),
        ("        3.3.2 อัลกอริทึมคำนวณตำแหน่งดวงอาทิตย์และการชดเชยเงา (SPA & Shadow Inversion)", "35"),
        ("        3.3.3 สถาปัตยกรรมเว็บแอปพลิเคชันและโครงร่างหน้าจอแสดงผล (Web UI Prototype)", "37"),
        ("    3.4 แผนงานและแนวทางแก้ปัญหาฮาร์ดแวร์ใน Project 2 (Hardware Evolution Roadmap)", "39"),
        ("บทที่ 4 ผลการดำเนินงานและการวิเคราะห์ผล (Results & Discussion)", "42"),
        ("    4.1 ผลการทดสอบวงจร Analog AFE บน Breadboard และ Perfboard", "42"),
        ("    4.2 ผลการตรวจวัดและชดเชยอุณหภูมิด้วย NTC Thermistor", "44"),
        ("    4.3 ผลการจำลองพฤติกรรมเชิงแสง (Optical Ray-Tracing Analysis)", "45"),
        ("    4.4 ผลการทดสอบส่งสัญญาณข้อมูลดิบและโครงร่างส่วนติดต่อผู้ใช้", "47"),
        ("บทที่ 5 สรุปผลการดำเนินงาน ปัญหา และแผนงานระยะถัดไป (Conclusion & Future Work)", "49"),
        ("    5.1 สรุปผลความก้าวหน้าของโครงงานที่ 1", "49"),
        ("    5.2 ปัญหา อุปสรรค และบทเรียนที่ได้รับ (Lessons Learned)", "50"),
        ("    5.3 แผนการดำเนินงานสำหรับโครงงานที่ 2 (Project 2 Implementation Plan)", "51"),
        ("บรรณานุกรม", "53")
    ]

    for title, pg in toc_items:
        r = t_toc.add_row()
        run0 = r.cells[0].paragraphs[0].add_run(title)
        run0.font.name = 'TH SarabunPSK'
        run0.font.size = Pt(14)
        run1 = r.cells[1].paragraphs[0].add_run(pg)
        run1.font.name = 'TH SarabunPSK'
        run1.font.size = Pt(14)
        r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_page_break()

    # =========================================================================
    # 8. LIST OF TABLES & LIST OF FIGURES
    # =========================================================================
    p_tab_title = doc.add_paragraph()
    p_tab_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tab_title.paragraph_format.space_after = Pt(12)
    r_tabt = p_tab_title.add_run("สารบัญตาราง")
    r_tabt.font.name = 'TH SarabunPSK'
    r_tabt.font.size = Pt(18)
    r_tabt.font.bold = True

    t_lot = doc.add_table(rows=1, cols=2)
    t_lot.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_lot.rows[0].cells[0].paragraphs[0].add_run("ตารางที่").font.bold = True
    t_lot.rows[0].cells[1].paragraphs[0].add_run("หน้า").font.bold = True
    t_lot.rows[0].cells[0].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_lot.rows[0].cells[1].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_lot.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    tables_list = [
        ("ตารางที่ 1.1 การแบ่งบทบาทหน้าที่และความรับผิดชอบของสมาชิกในกลุ่มโครงงาน", "4"),
        ("ตารางที่ 1.2 แผนการดำเนินงานวิจัยตลอดโครงงาน Project 1 และ Project 2", "6"),
        ("ตารางที่ 2.1 สรุปพารามิเตอร์การทำงานของ NTC 10 kΩ Semitec 103AT", "11"),
        ("ตารางที่ 2.2 คุณลักษณะทางไฟฟ้าของไอซีแรงดันอ้างอิง REF5045", "15"),
        ("ตารางที่ 2.3 ตารางเปรียบเทียบคุณลักษณะของเครื่องมือวัดตามหลักมาตรวิทยา", "21"),
        ("ตารางที่ 3.1 สรุปผลการทดสอบแรงดันอ้างอิงบน Breadboard กับการแก้ปัญหา", "29"),
        ("ตารางที่ 3.2 ตารางเปรียบเทียบคุณลักษณะระหว่างเซนเซอร์ BPX61 (TO-39) กับ TEMD5010X01 (SMD)", "40"),
        ("ตารางที่ 5.1 ตารางกำหนดเวลาและกิจกรรมการดำเนินงานใน Project 2 (Gantt Chart)", "52")
    ]
    for t_desc, pg in tables_list:
        r = t_lot.add_row()
        r.cells[0].paragraphs[0].add_run(t_desc).font.name = 'TH SarabunPSK'
        r.cells[1].paragraphs[0].add_run(pg).font.name = 'TH SarabunPSK'
        r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_sp_fig = doc.add_paragraph()
    p_sp_fig.paragraph_format.space_before = Pt(20)

    p_fig_title = doc.add_paragraph()
    p_fig_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_fig_title.paragraph_format.space_after = Pt(12)
    r_figt = p_fig_title.add_run("สารบัญรูป")
    r_figt.font.name = 'TH SarabunPSK'
    r_figt.font.size = Pt(18)
    r_figt.font.bold = True

    t_lof = doc.add_table(rows=1, cols=2)
    t_lof.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_lof.rows[0].cells[0].paragraphs[0].add_run("รูปที่").font.bold = True
    t_lof.rows[0].cells[1].paragraphs[0].add_run("หน้า").font.bold = True
    t_lof.rows[0].cells[0].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_lof.rows[0].cells[1].paragraphs[0].runs[0].font.name = 'TH SarabunPSK'
    t_lof.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    figures_list = [
        ("รูปที่ 1.1 ภาพรวมแนวคิดสถาปัตยกรรมเครื่องวัดไพราโนมิเตอร์แบบแยกส่วน", "2"),
        ("รูปที่ 2.1 วงจรสมมูลของโฟโตไดโอดและการตอบสนองต่อแสงอาทิตย์", "8"),
        ("รูปที่ 2.2 วงจรแบ่งแรงดันของ NTC Thermistor 10 kΩ และกราฟความสัมพันธ์ R-T", "10"),
        ("รูปที่ 2.3 โครงสร้างภายในของไอซีแปลงสัญญาณ ADC 24 บิต ADS1220", "13"),
        ("รูปที่ 2.4 ผังขาและวงจรภายในของไอซีแรงดันอ้างอิง REF5045", "15"),
        ("รูปที่ 2.5 📷 กราฟผลการจำลอง Directional Characteristics: Simulation with Metal Can", "17"),
        ("รูปที่ 2.6 📷 ภาพการเกิดเงาจำลอง High-Zoom Die Surface: Progressive Shadow by TO-39 Can", "18"),
        ("รูปที่ 2.7 📷 แผนภาพคุณลักษณะทิศทาง Directional Characteristics จาก Datasheet BPX61", "18"),
        ("รูปที่ 2.8 แผนภูมิห่วงโซ่การสอบกลับได้ของการวัด (Measurement Traceability Chain)", "22"),
        ("รูปที่ 3.1 บล็อกไดอะแกรมสถาปัตยกรรมระบบโดยรวม (System Block Diagram)", "24"),
        ("รูปที่ 3.2 แผนภาพวงจร (Schematic) บอร์ดขยายสัญญาณแอนะล็อก L1_Analog", "26"),
        ("รูปที่ 3.3 ภาพจำลองการต่อวงจรทดสอบขั้นแรกบนเบรดบอร์ด (Breadboard Setup)", "28"),
        ("รูปที่ 3.4 📷 ภาพถ่ายวงจรต้นแบบจริงบนแผงต่อวงจรเบรดบอร์ด (Breadboard Testing)", "28"),
        ("รูปที่ 3.5 📷 ภาพถ่ายวงจรต้นแบบบนแผงวงจรอเนกประสงค์ (Perfboard) มุมมองด้านบน (Top View)", "29"),
        ("รูปที่ 3.6 📷 ภาพถ่ายวงจรต้นแบบบนแผงวงจรอเนกประสงค์ (Perfboard) มุมมองด้านล่าง (Bottom View)", "30"),
        ("รูปที่ 3.7 แผนภาพวงจรบอร์ดประมวลผลไมโครคอนโทรลเลอร์ L2_MCU", "31"),
        ("รูปที่ 3.8 แผนภาพวงจรบอร์ดจัดการพลังงาน L3_Switching", "32"),
        ("รูปที่ 3.9 ผังงานการทำงานของโปรแกรมหลักบนไมโครคอนโทรลเลอร์ (Firmware Flowchart)", "34"),
        ("รูปที่ 3.10 แผนผังการทำงานของอัลกอริทึม Software Cosine & Temperature Correction", "36"),
        ("รูปที่ 3.11 📷 ภาพถ่ายหน้าจอส่วนติดต่อผู้ใช้เว็บแอปพลิเคชันและการปักหมุดพิกัด Google Maps", "38"),
        ("รูปที่ 3.12 โครงสร้างเปรียบเทียบเซนเซอร์ตัวถังกระป๋อง TO-39 กับ Flat SMD (TEMD5010X01)", "41"),
        ("รูปที่ 4.1 กราฟเปรียบเทียบเสถียรภาพแรงดันเอาต์พุตของการต่อ C10 ชิดขา REF5045", "43"),
        ("รูปที่ 4.2 กราฟความคลาดเคลื่อนโคไซน์จากการจำลองเงาเทียบกับฟังก์ชัน Ideal Cosine", "46"),
        ("รูปที่ 4.3 📷 ภาพหน้าจอการแสดงผลข้อมูลดิบผ่านพอร์ตสื่อสารอนุกรม (Serial Monitor)", "48")
    ]
    for f_desc, pg in figures_list:
        r = t_lof.add_row()
        r.cells[0].paragraphs[0].add_run(f_desc).font.name = 'TH SarabunPSK'
        r.cells[1].paragraphs[0].add_run(pg).font.name = 'TH SarabunPSK'
        r.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 1: บทนำ
    # =========================================================================
    add_heading_1(doc, "บทที่ 1\nบทนำ")

    add_heading_2(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา")
    add_body_p(
        doc,
        "ข้อมูลการแผ่รังสีแสงอาทิตย์รวมในแนวราบ (Global Horizontal Irradiance: GHI) มีความสำคัญอย่างยิ่งต่อการศึกษาด้านอุตุนิยมวิทยา "
        "การเกษตรกรรมแม่นยำ และการประเมินประสิทธิภาพของโรงไฟฟ้าพลังงานแสงอาทิตย์ อุปกรณ์มาตรฐานที่ใช้ในการตรวจวัดคือ ไพราโนมิเตอร์ (Pyranometer) "
        "ซึ่งต้องมีคุณสมบัติตอบสนองต่อแสงตามกฎของโคไซน์ (Lambert's Cosine Law) อย่างเคร่งครัด อย่างไรก็ตาม อุปกรณ์เกรดมาตรฐานวิจัยในท้องตลาดมีราคาสูง "
        "ในขณะที่การนำเซนเซอร์โฟโตไดโอดราคาประหยัดมาประยุกต์ใช้งานมักประสบปัญหาความคลาดเคลื่อนรุนแรง 3 ประการ ได้แก่:"
    )
    add_bullet_p(doc, "1. ความคลาดเคลื่อนเชิงมุม (Cosine Error):", "โครงสร้างทางกลของตัวถังเซนเซอร์ (เช่น ขอบกระป๋องโลหะ TO-39) บดบังแสงที่มุมตกกระทบสูง ทำให้ค่าแสงที่วัดได้ตกลงต่ำกว่าทฤษฎีอย่างมีนัยสำคัญ")
    add_bullet_p(doc, "2. ความคลาดเคลื่อนจากอุณหภูมิ (Temperature Drift):", "ค่าสัมประสิทธิ์การตอบสนองต่อแสง (Responsivity) ของสารกึ่งตัวนำซิลิกอนจะแปรผันตามอุณหภูมิผิวของตัวรับแสง (+0.1% ถึง +0.2%/°C)")
    add_bullet_p(doc, "3. สัญญาณรบกวนทางไฟฟ้า (Circuit Noise & Parasitic Leakage):", "สัญญาณแรงดันขนาดเล็กระดับไมโครโวลต์ถูกรบกวนได้ง่ายจากสัญญาณสวิตชิ่งและกระแสไฟฟ้ารั่วไหลบนแผงวงจร")

    add_photo_placeholder(
        doc,
        "1.1",
        "ภาพรวมแนวคิดสถาปัตยกรรมเครื่องวัดไพราโนมิเตอร์แบบแยกส่วน",
        "แผนภาพแสดงการรับแสงอาทิตย์เข้าสู่เซนเซอร์ BPX61, การปรับสภาพสัญญาณผ่าน L1_Analog, แปลงค่าด้วย ADS1220 24 บิต, ส่งข้อมูลสู่ L2_MCU และประมวลผลอัลกอริทึมบนเว็บแอปพลิเคชัน",
        "Export แผนภาพสถาปัตยกรรมระบบโดยรวม (Conceptual System Architecture) ในความละเอียดสูง"
    )

    add_heading_2(doc, "1.2 วัตถุประสงค์ของโครงงาน")
    add_bullet_p(doc, "1.", "เพื่อออกแบบและพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (High-Precision Solar Irradiance Measurement System) โดยใช้สถาปัตยกรรมฮาร์ดแวร์แบบแยกส่วน")
    add_bullet_p(doc, "2.", "เพื่อศึกษาและวิเคราะห์พฤติกรรมการบดบังแสงของตัวถังเซนเซอร์ BPX61 (TO-39 Metal Can) และข้อจำกัดในการจำลองการกระเจิงแสงของวัสดุ PTFE")
    add_bullet_p(doc, "3.", "เพื่อออกแบบและสร้างวงจรต้นแบบทดสอบขั้นกลางบนแผงวงจรอเนกประสงค์ (Perfboard) และชดเชยผลกระทบของอุณหภูมิด้วย NTC Thermistor 10 kΩ")
    add_bullet_p(doc, "4.", "เพื่อพัฒนาอัลกอริทึมชดเชยความคลาดเคลื่อนโคไซน์บนซอฟต์แวร์ (Software Cosine Correction) ร่วมกับการคำนวณตำแหน่งดวงอาทิตย์ (SPA) และเตรียมความพร้อมสำหรับฮาร์ดแวร์เซนเซอร์ Flat SMD ในโครงงานที่ 2")

    add_heading_2(doc, "1.3 ขอบเขตของโครงงาน (Project 1 และ Project 2)")
    add_body_p(
        doc,
        "เพื่อให้การดำเนินงานวิจัยสอดคล้องกับกรอบระยะเวลาและสามารถควบคุมความเสี่ยงเชิงวิศวกรรมได้อย่างมีประสิทธิภาพ คณะผู้จัดทำได้กำหนดขอบเขตการทำงานแบ่งเป็น 2 ระยะดังนี้:"
    )
    add_bullet_p(doc, "ขอบเขตโครงงานที่ 1 (Project 1 - ภาคการศึกษาที่ 1):", 
                 "มุ่งเน้นการพิสูจน์การทำงานของวงจรภาคแอนะล็อกความแม่นยำสูง โดยประกอบวงจรต้นแบบลงบนแผงวงจรอเนกประสงค์ (Perfboard), การแก้ปัญหาความไม่เสถียรของแรงดันอ้างอิง REF5045, "
                 "การอ่านค่าอุณหภูมิ NTC 10 kΩ, การจำลองทางแสง (Optical Ray-Tracing) เพื่อหาฟังก์ชันการบังแสงของตัวถัง TO-39, และการออกแบบโครงสร้างสถาปัตยกรรมซอฟต์แวร์ต้นแบบ (Software & Web Prototype)")
    add_bullet_p(doc, "ขอบเขตโครงงานที่ 2 (Project 2 - ภาคการศึกษาที่ 2):", 
                 "มุ่งเน้นการแก้ปัญหาที่ต้นตอกายภาพด้วยการย้ายแพ็กเกจเซนเซอร์ไปสู่ Flat SMD (Vishay TEMD5010X01), การผลิตแผ่นวงจรพิมพ์มาตรฐาน 4 เลเยอร์ (Custom 4-Layer PCB), "
                 "การทดสอบเปรียบเทียบ A/B Testing บนแท่นหมุนวัดมุม (Goniometer) ในห้องปฏิบัติการทัศนศาสตร์ อาคารจุฬาภรณ์ 1 และการสอบเทียบค่าความเข้มแสงกับเครื่องจำลองแสงอาทิตย์ Solar Light XPS-300 อาคาร 55, "
                 "ตลอดจนการพัฒนา Web Application ฉบับสมบูรณ์ที่เชื่อมต่อคลาวด์")

    add_heading_2(doc, "1.4 การแบ่งบทบาทและความรับผิดชอบในทีม (Division of Responsibilities)")
    add_body_p(
        doc,
        "โครงงานนี้เป็นการบูรณาการระหว่างหลักสูตรวิศวกรรมระบบไอโอทีและสารสนเทศ ร่วมกับฟิสิกส์อุตสาหกรรม จึงได้จัดแบ่งหน้าที่ความรับผิดชอบตามความเชี่ยวชาญ ดังแสดงในตารางที่ 1.1:"
    )

    # Table 1.1 Division of Responsibilities
    t_div = doc.add_table(rows=3, cols=3)
    t_div.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_div_headers = ["สมาชิกในทีม", "บทบาทหน้าที่หลัก", "ขอบเขตงานที่รับผิดชอบ"]
    for c_idx, h_text in enumerate(t_div_headers):
        c = t_div.cell(0, c_idx)
        set_cell_shading(c, "E2E8F0")
        run = c.paragraphs[0].add_run(h_text)
        run.font.name = 'TH SarabunPSK'
        run.font.bold = True

    div_rows = [
        (
            "นายทชินทร์ รักษาวงศ์\n(รหัสนักศึกษา 6601xxxx)",
            "Hardware & Applied Physics Lead",
            "• ออกแบบวงจร Analog Front-End (AFE), REF5045 และ ADC 24 บิต บน KiCad\n"
            "• ประกอบและทดสอบวงจรต้นแบบบน Breadboard และ Perfboard\n"
            "• ดำเนินการจำลองทางแสง Optical Ray-Tracing วิเคราะห์เงาตัวถัง TO-39 และวัสดุ PTFE\n"
            "• ออกแบบ Custom PCB 4 เลเยอร์สำหรับเซนเซอร์ Flat SMD ใน Project 2\n"
            "• รับผิดชอบการทดสอบในห้องปฏิบัติการฟิสิกส์ (Optics Lab & Solar Simulator)"
        ),
        (
            "(ชื่อนักศึกษาคนที่ 2)\n(รหัสนักศึกษา 6601xxxx)",
            "Software & IoT Cloud Lead",
            "• พัฒนาเฟิร์มแวร์ไมโครคอนโทรลเลอร์ (L2_MCU) สำหรับอ่านค่า ADC และ NTC\n"
            "• ออกแบบสถาปัตยกรรมระบบสื่อสารข้อมูล (IoT Telemetry / REST API / MQTT)\n"
            "• พัฒนา Web Application Dashboard และส่วนติดต่อผู้ใช้ (UI Prototype)\n"
            "• พัฒนาอัลกอริทึม Solar Position Algorithm (SPA) ดึงพิกัด Google Maps\n"
            "• ดำเนินการประมวลผล Software Cosine Correction บนระบบ Cloud ใน Project 2"
        )
    ]
    for r_idx, row_data in enumerate(div_rows, start=1):
        for c_idx, val in enumerate(row_data):
            c = t_div.cell(r_idx, c_idx)
            run = c.paragraphs[0].add_run(val)
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(14)

    add_heading_2(doc, "1.5 ประโยชน์ที่คาดว่าจะได้รับ")
    add_bullet_p(doc, "1.", "ได้ระบบตรวจวัดการแผ่รังสีแสงอาทิตย์ที่มีความแม่นยำสูง ต้นทุนต่ำ และสามารถนำไปประยุกต์ใช้งานในสถานีตรวจวัดสภาพอากาศและโรงไฟฟ้าพลังงานแสงอาทิตย์ได้จริง")
    add_bullet_p(doc, "2.", "ได้ระเบียบวิธีแก้ไขปัญหา Cosine Error ด้วยการผสมผสานระหว่างการชดเชยทางซอฟต์แวร์ (Computational Sensing) และการเลือกใช้อุปกรณ์ฮาร์ดแวร์ที่เหมาะสม")
    add_bullet_p(doc, "3.", "เข้าใจหลักการทางมาตรวิทยา การควบคุมสัญญาณรบกวนในวงจรแอนะล็อกความไวสูง และการบูรณาการองค์ความรู้ระหว่างฮาร์ดแวร์ ทัศนศาสตร์ และไอโอที")

    add_heading_2(doc, "1.6 แผนการดำเนินงานตลอดโครงงาน")
    add_body_p(doc, "ตารางที่ 1.2 แสดงแผนการดำเนินงานวิจัยตลอดโครงงาน Project 1 และ Project 2:")

    # Table 1.2 Plan
    t_plan = doc.add_table(rows=7, cols=4)
    t_plan.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_plan_headers = ["ระยะเวลา", "กิจกรรมการดำเนินงาน", "ผลผลิตที่คาดหวัง", "สถานะ"]
    for c_idx, h_text in enumerate(t_plan_headers):
        c = t_plan.cell(0, c_idx)
        set_cell_shading(c, "E2E8F0")
        run = c.paragraphs[0].add_run(h_text)
        run.font.name = 'TH SarabunPSK'
        run.font.bold = True

    plan_rows = [
        ("ส.ค. - ก.ย. 2569", "ศึกษาทฤษฎีเซนเซอร์, วงจร AFE และสร้างโมเดลจำลองเงา TO-39 & PTFE", "แบบจำลอง Ray-tracing Simulation กราฟเงา", "เสร็จสิ้น"),
        ("ก.ย. - ต.ค. 2569", "ออกแบบ Schematic บอร์ด L1, L2, L3 บน KiCad 10", "ไฟล์วงจร KiCad Schematic ครบ 3 บอร์ด", "เสร็จสิ้น"),
        ("ต.ค. 2569", "ทดสอบแก้บั๊กวงจรบน Breadboard และประกอบลง Perfboard", "วงจรต้นแบบ Perfboard พร้อมทดสอบสัญญาณ", "เสร็จสิ้น"),
        ("ต.ค. 2569", "พัฒนาโครงร่าง Web UI และส่งรายงานความก้าวหน้า Project 1", "เล่มรายงาน Project 1 และสไลด์นำเสนอ", "กำลังดำเนินการ"),
        ("พ.ย. - ธ.ค. 2569", "ออกแบบ Custom PCB สำหรับ Flat SMD (TEMD5010X01) และทำ Web App", "แผ่นวงจรพิมพ์ PCB 4 เลเยอร์ และเว็บแดชบอร์ด", "แผนงาน Project 2"),
        ("ม.ค. - ก.พ. 2570", "ทดสอบกวาดมุมในแล็บทัศนศาสตร์ และสอบเทียบกับเครื่องจำลองแสง XPS-300", "ผลการทดสอบ A/B Testing และเล่มปริญญานิพนธ์สมบูรณ์", "แผนงาน Project 2")
    ]
    for r_idx, row_data in enumerate(plan_rows, start=1):
        for c_idx, val in enumerate(row_data):
            c = t_plan.cell(r_idx, c_idx)
            run = c.paragraphs[0].add_run(val)
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(14)

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 2: ทฤษฎีและงานวิจัยที่เกี่ยวข้อง
    # =========================================================================
    add_heading_1(doc, "บทที่ 2\nทฤษฎีและงานวิจัยที่เกี่ยวข้อง")

    add_heading_2(doc, "2.1 หลักการตรวจวัดการแผ่รังสีแสงอาทิตย์และกฎของโคไซน์ (Cosine Law)")
    add_body_p(
        doc,
        "การแผ่รังสีแสงอาทิตย์รวมในแนวราบ (Global Horizontal Irradiance: GHI) ประกอบด้วยสองส่วนหลักคือ รังสีตรงจากดวงอาทิตย์ (Direct Beam Irradiance) "
        "และรังสีกระจายจากท้องฟ้า (Diffuse Sky Irradiance) ตามกฎของแลมเบิร์ต (Lambert's Cosine Law) ความหนาแน่นฟลักซ์การแผ่รังสีที่ตกกระทบลงบนพื้นผิวระนาบแนวนอน "
        "จะแปรผันตรงกับค่าโคไซน์ของมุมตกกระทบ (Incident Angle หรือ Solar Zenith Angle: θz) ดังสมการ:"
    )
    add_body_p(doc, "E(θz) = E_0 · cos(θz)")
    add_body_p(
        doc,
        "เครื่องวัดไพราโนมิเตอร์ในอุดมคติต้องมีสภาพไวเชิงมุม (Directional Sensitivity) สอดคล้องกับฟังก์ชัน cos(θz) อย่างสมบูรณ์ "
        "อย่างไรก็ตาม ในทางปฏิบัติ เซนเซอร์ตรวจวัดแสงทั่วไปจะมีความคลาดเคลื่อนเชิงมุม (Cosine Error) สูงมากเมื่อมุมตกกระทบเกิน 40 องศาขึ้นไป "
        "อันส่งผลให้ค่าพลังงานแสงอาทิตย์ที่ตรวจวัดได้ในช่วงเช้าและช่วงเย็นคลาดเคลื่อนต่ำกว่าความเป็นจริงอย่างรุนแรง"
    )

    add_photo_placeholder(
        doc,
        "2.1",
        "วงจรสมมูลของโฟโตไดโอดและการตอบสนองต่อแสงอาทิตย์",
        "แผนภาพวงจรสมมูลของซิลิกอนโฟโตไดโอด BPX61 แสดงแหล่งกำเนิดกระแส I_ph, ตัวต้านทานขนาน R_sh, ตัวเก็บประจุรอยต่อ C_j และตัวต้านทานอนุกรม R_s",
        "นำภาพไดอะแกรมวงจรสมมูลของโฟโตไดโอด"
    )

    add_heading_2(doc, "2.2 คุณสมบัติของ Silicon Photodiode และผลกระทบของอุณหภูมิ")
    add_body_p(
        doc,
        "ซิลิกอนโฟโตไดโอด (Silicon Photodiode) รุ่น BPX61 เป็นเซนเซอร์ตรวจจับแสงเกรดความแม่นยำสูงที่มีพื้นที่รับแสงขนาดใหญ่ (7.02 mm²) "
        "และมีสภาพไวสเปกตรัม (Spectral Responsivity) ครอบคลุมความยาวคลื่นตั้งแต่ 400 nm ถึง 1100 nm อย่างไรก็ตาม สารกึ่งตัวนำซิลิกอนมีสัมประสิทธิ์อุณหภูมิที่แปรผันอย่างมีนัยสำคัญ "
        "โดยเฉพาะในช่วงความยาวคลื่นใกล้อินฟราเรด ซึ่งแถบพลังงานต้องห้าม (Bandgap Energy: Eg) จะแคบลงเมื่ออุณหภูมิเพิ่มขึ้นตามสมการของ Varshni ส่งผลให้ค่า Responsivity "
        "เพิ่มขึ้นตามอุณหภูมิ (α ≈ +0.1% ถึง +0.2%/°C) หากหัววัดตากแดดจนมีอุณหภูมิสะสมสูงถึง 60°C ค่าความเข้มแสงที่วัดได้จะสูงเกินจริงถึง 5–7%"
    )
    add_body_p(
        doc,
        "ในการตรวจวัดอุณหภูมิเพื่อนำมาชดเชยค่า ได้เลือกใช้ NTC Thermistor 10 kΩ (Semitec 103AT) ซึ่งมีความแม่นยำสูง (B25/85 = 3435 K ±1%) "
        "โดยจัดวงจรแบ่งแรงดันร่วมกับตัวต้านทานความแม่นยำสูง ค่าความต้านทานที่เปลี่ยนไปตามอุณหภูมิสามารถแปลงเป็นอุณหภูมิสัมบูรณ์ได้ด้วยสมการ Steinhart-Hart:"
    )
    add_body_p(doc, "1 / T = A + B · ln(R) + C · (ln(R))^3")

    add_photo_placeholder(
        doc,
        "2.2",
        "วงจรแบ่งแรงดันของ NTC Thermistor 10 kΩ และกราฟความสัมพันธ์ R-T",
        "ภาพวงจรการต่อ NTC ร่วมกับตัวต้านทานอ้างอิงความแม่นยำสูง และกราฟแสดงความสัมพันธ์ระหว่างค่าความต้านทานกับอุณหภูมิตามพารามิเตอร์ Beta 3435 K",
        "นำภาพวงจรและกราฟ R-T characteristic จากเอกสาร Semitec 103AT"
    )

    add_heading_2(doc, "2.3 วงจรแปลงสัญญาณแอนะล็อกเป็นดิจิทัล 24 บิต (ADS1220)")
    add_body_p(
        doc,
        "ไอซี Texas Instruments ADS1220 เป็น ADC แบบ Delta-Sigma ขนาด 24 บิต ที่มีวงจรขยาย Programmable Gain Amplifier (PGA) สัญญาณรบกวนต่ำในตัว "
        "ทำให้สามารถตรวจจับการเปลี่ยนแปลงของแรงดันไฟฟ้าได้ละเอียดถึงระดับไมโครโวลต์ (µV) การเชื่อมต่ออินพุตเป็นแบบดิฟเฟอเรนเชียล (Differential Input) "
        "ช่วยขจัดสัญญาณรบกวนในโหมดร่วมได้อย่างดีเยี่ยม และมีอัตราการสุ่มตัวอย่างปรับได้ตั้งแต่ 20 SPS ถึง 2000 SPS"
    )

    add_photo_placeholder(
        doc,
        "2.3",
        "โครงสร้างภายในของไอซีแปลงสัญญาณ ADC 24 บิต ADS1220",
        "บล็อกไดอะแกรมภายในของ ADS1220 แสดง Multiplexer, Low-Noise PGA, 24-bit Modulator, Digital Filter และ SPI Interface",
        "นำภาพจาก Datasheet ของ Texas Instruments ADS1220"
    )

    add_heading_2(doc, "2.4 วงจรแรงดันอ้างอิงความเสถียรสูง (REF5045) และปัญหา High-Z Leakage")
    add_body_p(
        doc,
        "ไอซี REF5045 จ่ายแรงดันอ้างอิงคงที่ 4.500 V พร้อมค่าการเปลี่ยนแปลงตามอุณหภูมิต่ำเป็นพิเศษ (Temperature Drift 3 ppm/°C) "
        "อย่างไรก็ตาม จากการทดสอบวงจรจริงบนเบรดบอร์ดพบปัญหาสำคัญที่ขา Pin 5 (Trim/NR) ซึ่งเป็นจุดเชื่อมต่อโครงข่ายความต้านทานภายใน R_int ≈ 100 kΩ (High-Z Node) "
        "การเสียบชิปลงบนเบรดบอร์ดทำให้เกิดกระแสไฟฟ้ารั่วไหลข้ามรางสปริง ส่งผลให้แรงดันอ้างอิงตกลงถึง 118 mV "
        "นอกจากนี้ หากวงจรขาดตัวเก็บประจุ C10 (10 µF ชนิด Low-ESR) ที่ขาเอาต์พุต วงจรป้อนกลับภายในของไอซีจะเกิดการสั่นพวง (Loop Oscillation) ที่ความถี่ประมาณ 250 kHz ทันที"
    )

    add_photo_placeholder(
        doc,
        "2.4",
        "ผังขาและวงจรภายในของไอซีแรงดันอ้างอิง REF5045",
        "ภาพผังขา SOIC-8 แสดงตำแหน่งขา 5 (Trim/NR จุด High-Z) และขา 6 (VOUT พร้อมตัวเก็บประจุ C10)",
        "ใช้ภาพไดอะแกรมผังขา REF5045 ที่มีสัญลักษณ์เตือน High-Z Node"
    )

    add_heading_2(doc, "2.5 ข้อจำกัดของวัสดุ PTFE และการเกิดเงาของตัวถัง TO-39 (Cosine Simulation)")
    add_body_p(
        doc,
        "ในการออกแบบเครื่องวัดไพราโนมิเตอร์ช่วงเริ่มต้น คณะผู้จัดทำได้ตั้งสมมติฐานว่าสามารถแก้ไขความคลาดเคลื่อนโคไซน์ของเซนเซอร์ BPX61 "
        "ได้โดยการสร้างฝาครอบโดมกระจายแสงโพลีเตตระฟลูออโรเอทิลีน (PTFE Diffuser Dome) ครอบบนตัวรับแสง เช่นเดียวกับโครงสร้างที่ใช้ในไพราโนมิเตอร์เชิงพาณิชย์ "
        "อย่างไรก็ตาม เมื่อได้เริ่มทำการจำลองทางแสงเชิงลึก (Optical Ray-Tracing Simulation) และศึกษาคุณสมบัติการกระเจิงแสงในเนื้อวัสดุ (Bulk Scattering Physics) "
        "ได้ค้นพบข้อจำกัดสำคัญ 3 ประการดังนี้:"
    )
    add_bullet_p(doc, "1. ความซับซ้อนของตัวแปรการกระเจิงแสง (Scattering Parameter Complexity):", 
                 "การกระเจิงแสงในเนื้อ PTFE ไม่สามารถจำลองด้วยดัชนีหักเหแสงทั่วไปได้ แต่ต้องกำหนดค่าสัมประสิทธิ์การกระเจิงแบบกลุ่ม (Bulk Scattering Coefficient: μs), "
                 "สัมประสิทธิ์การดูดกลืนแสง (Absorption Coefficient: μa) และฟังก์ชันการกระจายมุมแบบ Henyey-Greenstein (Anisotropy Factor: g) "
                 "ซึ่งค่าเหล่านี้ขึ้นอยู่กับขนาดอนุภาค ความหนาแน่นของการอัด และอุณหภูมิการเผาผนึก (Sintering) ของเนื้อ PTFE อย่างยิ่ง")
    add_bullet_p(doc, "2. ปัญหาข้อแลกเปลี่ยนด้านความหนา (Thickness vs Attenuation Trade-off):", 
                 "หากความหนาของโดม PTFE มากเกินไป แสงจะถูกลดทอนจนสัญญาณกระแสแสงดรอปต่ำกว่าระดับสัญญาณรบกวน (Noise Floor) ของวงจรขยาย "
                 "แต่หากโดมบางเกินไป โดมจะไม่สามารถกระเจิงแสงได้อย่างสมบูรณ์ (Direct Ballistic Transmission) ทำให้ยังคงเกิด Cosine Error รุนแรง")
    add_bullet_p(doc, "3. ข้อจำกัดทางกลของตัวถัง TO-39 (Aperture Cavity Shadowing):", 
                 "ตัวถังกระป๋องโลหะ TO-39 ของ BPX61 มีขอบยื่นสูงรอบผลึกรับแสง (Depth ~1.5–2 mm) ส่งผลให้เกิดหลุมลึก แสงที่กระเจิงเฉียงๆ จากฐานของโดม PTFE "
                 "ยังคงถูกขอบกระป๋องโลหะบดบังเนื้อซิลิคอนอยู่ดี (Internal Vignetting) ส่งผลให้การตอบสนองที่มุม θ > 40° ตกลงอย่างรวดเร็ว และเกิดจุดตัดแสงสิ้นเชิง (Cutoff) ที่ θ ≥ 70°")

    add_photo_placeholder(
        doc,
        "2.5",
        "📷 กราฟผลการจำลอง Directional Characteristics: Simulation with Metal Can",
        "กราฟเปรียบเทียบระหว่างเส้นประสีดำ Ideal Cosine Law (cos θ) กับเส้นสีน้ำเงิน BPX61 with Can (Simulation) แสดงการตกลงอย่างรวดเร็วที่มุมเกิน 40° และจุดตัดแสงที่ 70°",
        "📷 นำภาพกราฟ Directional Characteristics: Simulation with Metal Can ของพี่มาวางที่นี่"
    )

    add_photo_placeholder(
        doc,
        "2.6",
        "📷 ภาพการเกิดเงาจำลอง High-Zoom Die Surface: Progressive Shadow by TO-39 Can",
        "ภาพ Heatmap แสดงความเข้มแสงตกกระทบผิว Die 10 ช่อง (θ = 0° ถึง 90°) แสดงการเกิดเงาโค้งของขอบกระป๋อง TO-39 บดบังผิว Die อย่างชัดเจนที่ θ = 50° และ 60°",
        "📷 นำภาพ Heatmap 10 ช่อง High-Zoom Die Surface ของพี่มาวางที่นี่"
    )

    add_photo_placeholder(
        doc,
        "2.7",
        "📷 แผนภาพคุณลักษณะทิศทาง Directional Characteristics จาก Datasheet BPX61",
        "กราฟ Polar Plot จากคู่มือทางการของ BPX61 แสดงมุมตัดแสงที่ 60°–70° สอดคล้องกับผลการจำลอง",
        "📷 นำภาพ Polar Plot จาก Datasheet ของพี่มาวางที่นี่"
    )

    add_body_p(
        doc,
        "จากการประเมินความเสี่ยงเชิงวิศวกรรม (Engineering Risk Assessment) พบว่าการพยายามขึ้นรูปโดม PTFE ด้วยเครื่องอัดไฮดรอลิกในห้องปฏิบัติการสารกึ่งตัวนำ "
        "โดยปราศจากเครื่องมือวัดการกระจายแสงเฉพาะทาง (Goniophotometer หรือ Integrating Sphere) เพื่อควบคุมคุณภาพ จะไม่สามารถรับประกันความสม่ำเสมอและคุณสมบัติที่ทำซ้ำได้ (Repeatability) "
        "และเสี่ยงต่อการทำให้โครงงานไม่เสร็จตามกำหนดเวลา คณะผู้จัดทำจึงตัดสินใจปรับแนวทางเชิงกลยุทธ์ (Pragmatic Pivot) "
        "โดยใช้ Software Cosine Inversion บนเว็บแอปพลิเคชันเพื่อชดเชยเงาของตัวถัง BPX61 ในโครงงานที่ 1 และวางแผนย้ายไปใช้เซนเซอร์แบบแบนราบ Flat SMD ในโครงงานที่ 2"
    )

    add_heading_2(doc, "2.6 หลักการมาตรวิทยา การสอบเทียบ และห่วงโซ่การสอบกลับได้ (Traceability)")
    add_body_p(
        doc,
        "ตามหลักการมาตรวิทยา ผลการวัดที่น่าเชื่อถือต้องสามารถเชื่อมโยงกลับไปยังหน่วยวัดมาตรฐานสากล (SI Units) ผ่านห่วงโซ่การสอบเทียบ (Traceability Chain) "
        "โดยไม่ขาดตอน โดยเครื่องวัดไพราโนมิเตอร์จะต้องได้รับการสอบเทียบเปรียบเทียบกับเครื่องมือวัดอ้างอิงมาตรฐาน (Working Standard) "
        "เช่น ไพราโนมิเตอร์ Class A หรือการทดสอบภายใต้แหล่งกำเนิดแสงจำลองมาตรฐาน (Solar Simulator Class AAA)"
    )

    add_photo_placeholder(
        doc,
        "2.8",
        "แผนภูมิห่วงโซ่การสอบกลับได้ของการวัด (Measurement Traceability Chain)",
        "แผนภาพแสดงการสอบกลับได้จากมาตรฐาน SI ผ่านสถาบันมาตรวิทยาแห่งชาติ (NIMT) ลงมายังห้องปฏิบัติการสอบเทียบ และเครื่องมือวัดใช้งาน",
        "นำภาพไดอะแกรม Traceability Chain จากสไลด์วิชามาตรวิทยาบทที่ 3"
    )

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 3: การออกแบบระบบและการดำเนินงาน (System Design & Implementation)
    # =========================================================================
    add_heading_1(doc, "บทที่ 3\nการออกแบบระบบและการดำเนินงาน\n(System Design & Implementation)")

    add_heading_2(doc, "3.1 ภาพรวมสถาปัตยกรรมระบบแบบ Computational Sensing")
    add_body_p(
        doc,
        "ระบบวัดไพราโนมิเตอร์ถูกออกแบบโดยใช้แนวคิดการวัดเชิงคำนวณ (Computational Sensing) โดยแบ่งระบบออกเป็น 2 ชั้นการทำงานหลัก:"
    )
    add_bullet_p(doc, "1. ฮาร์ดแวร์เก็บข้อมูลสัญญาณต่ำ (Low-Noise Data Acquisition Hardware):", "ทำหน้าที่แปลงกระแสแสงขนาดไมโครแอมแปร์เป็นแรงดันไฟฟ้า ขยายสัญญาณ แปลงเป็นดิจิทัล 24 บิต และตรวจวัดอุณหภูมิผิวของเซนเซอร์")
    add_bullet_p(doc, "2. เลเยอร์ซอฟต์แวร์ชดเชยค่าเชิงอัลกอริทึม (Algorithmic Software Layer):", "รับค่าพิกัดภูมิศาสตร์ (Latitude/Longitude) จาก Google Maps ร่วมกับเวลาจริง เพื่อคำนวณ Solar Zenith Angle และทำการคูณชดเชยฟังก์ชันเงาของกระป๋อง TO-39 และอุณหภูมิ")

    add_photo_placeholder(
        doc,
        "3.1",
        "บล็อกไดอะแกรมสถาปัตยกรรมระบบโดยรวม (System Block Diagram)",
        "แผนภาพบล็อกไดอะแกรมแสดงการเชื่อมต่อ: BPX61 + NTC -> L1_Analog (AFE & 24-bit ADC) -> L2_MCU (ESP32) -> Web Interface (Solar Angle Calculation & Cosine Inversion Model)",
        "วาดเป็น Block Diagram แสดงการไหลของสัญญาณและเลเยอร์ซอฟต์แวร์อย่างชัดเจน"
    )

    add_heading_2(doc, "3.2 การออกแบบฮาร์ดแวร์ต้นแบบ (Hardware Prototype Implementation)")
    add_heading_3(doc, "3.2.1 วงจรแอนะล็อกส่วนหน้า (L1_Analog Front-End)")
    add_body_p(
        doc,
        "วงจร L1_Analog ออกแบบบนโปรแกรม KiCad 10 โดยใช้ชิป ADS1220 ร่วมกับ REF5045 "
        "โดยจัดโครงสร้างสัญญาณแบบ Differential เพื่อลดสัญญาณรบกวนในโหมดร่วม"
    )

    add_photo_placeholder(
        doc,
        "3.2",
        "แผนภาพวงจร (Schematic) บอร์ดขยายสัญญาณแอนะล็อก L1_Analog",
        "ภาพแคปเจอร์วงจร Schematic ของแผ่นวงจร L1_Analog จากโปรแกรม KiCad แสดงการต่อขาของ REF5045, ADS1220, ตัวเก็บประจุ C10, C11 และขั้วต่อเซนเซอร์",
        "Export ภาพ Schematic จากไฟล์ L1_Analog/L1_Analog.kicad_sch ในความละเอียดสูง"
    )

    add_heading_3(doc, "3.2.2 การสร้างและแก้ปัญหาบน Breadboard สู่การประกอบลง Perfboard")
    add_body_p(
        doc,
        "เพื่อพิสูจน์การทำงานจริงของวงจรก่อนสั่งผลิต PCB ได้แบ่งการทดสอบออกเป็น 2 ขั้นตอน:"
    )
    add_body_p(
        doc,
        "1. การทดสอบบน Breadboard: พบปัญหาแรงดันอ้างอิงตก 118 mV อันเนื่องมาจากกระแสรั่วไหลบนหน้าสัมผัสเบรดบอร์ด "
        "จึงได้ข้อสรุปเชิงวิศวกรรมให้เว้นว่าง (DNP) ตัวเก็บประจุ C11 บนขา 5 ในการทดสอบเบื้องต้น และต้องติดตั้ง C10 (10 µF) ชิดขาไอซีเสมอเพื่อป้องกันการสั่นพวง"
    )

    add_photo_placeholder(
        doc,
        "3.3",
        "ภาพจำลองการต่อวงจรทดสอบขั้นแรกบนเบรดบอร์ด (Breadboard Setup)",
        "ไดอะแกรมจำลองการเสียบชิปบน Breadboard แสดงเส้นทางกระแสรั่วไหลที่ทำให้เกิดปัญหาแรงดันตก 118 mV",
        "นำภาพไดอะแกรมจาก assets/ref5045/fig2_breadboard.svg"
    )

    add_photo_placeholder(
        doc,
        "3.4",
        "📷 ภาพถ่ายวงจรต้นแบบจริงบนแผงต่อวงจรเบรดบอร์ด (Breadboard Testing)",
        "ภาพถ่ายจริงของแผงเบรดบอร์ดที่ใช้ทดลองในแล็บ มีบอร์ดแปลง SOIC-8, สายจัมเปอร์, และมิเตอร์วัดค่า",
        "📷 ถ่ายรูปแผงเบรดบอร์ดขณะทำการทดลองจริงในห้องทดลอง"
    )

    add_body_p(
        doc,
        "2. การประกอบลงแผงวงจรอเนกประสงค์ (Perfboard): ย้ายวงจรมาบัดกรีลงแผงบอร์ดไข่ปลาเพื่อขจัดปัญหาหน้าสัมผัสหลวม "
        "ลดค่าความจุแฝง และป้องกันกระแสไฟฟ้ารั่วไหล โดยวางตัวเก็บประจุ C10 ชิดขาชิปมากที่สุด และใช้การเดินสายแบบ Star Ground"
    )

    add_photo_placeholder(
        doc,
        "3.5",
        "📷 ภาพถ่ายวงจรต้นแบบบนแผงวงจรอเนกประสงค์ (Perfboard) มุมมองด้านบน (Top View)",
        "ภาพถ่ายบอร์ดไข่ปลาที่ประกอบเสร็จด้านบน เห็นการจัดวางไอซี, ซ็อกเก็ต, เทอร์มินัลบล็อก และตัวเก็บประจุ C10 ชิดขาชิป",
        "📷 นำบอร์ด Perfboard ที่บัดกรีเสร็จแล้วมาถ่ายภาพมุมตั้งฉาก 90 องศา บนพื้นผิวที่มีแสงสว่างชัดเจน"
    )

    add_photo_placeholder(
        doc,
        "3.6",
        "📷 ภาพถ่ายวงจรต้นแบบบนแผงวงจรอเนกประสงค์ (Perfboard) มุมมองด้านล่าง (Bottom View)",
        "ภาพถ่ายด้านหลังของบอร์ดไข่ปลา แสดงรอยบัดกรีตะกั่วที่แน่นหนา สะอาด และการเดินสายเชื่อมต่อแบบสั้น",
        "📷 ถ่ายรูปลายบัดกรีของบอร์ด Perfboard ซูมให้เห็นความเรียบร้อยของจุดบัดกรีและสายไฟ"
    )

    add_heading_3(doc, "3.2.3 บอร์ดประมวลผลและการส่งข้อมูล (L2_MCU Board)")
    add_body_p(
        doc,
        "บอร์ดประมวลผล L2_MCU ใช้ไมโครคอนโทรลเลอร์ ESP32 ทำหน้าที่เชื่อมต่อสื่อสารกับไอซี ADS1220 ผ่านบัส SPI "
        "เพื่ออ่านค่าดิจิทัลดิบ (Raw ADC Counts) ทั้งจากช่องสัญญาณกระแสแสงและช่องสัญญาณวัดอุณหภูมิ NTC "
        "จากนั้นส่งข้อมูลออกไปยังคอมพิวเตอร์ผ่านพอร์ต USB-UART เพื่อเตรียมส่งต่อเข้าสู่ระบบเว็บแอปพลิเคชัน"
    )

    add_photo_placeholder(
        doc,
        "3.7",
        "แผนภาพวงจรบอร์ดประมวลผลไมโครคอนโทรลเลอร์ L2_MCU",
        "ภาพ Schematic ของบอร์ด L2_MCU จาก KiCad แสดงไมโครคอนโทรลเลอร์และการต่อขาพอร์ต",
        "Export ภาพวงจรจากไฟล์ L2_MCU/L2_MCU.kicad_sch"
    )

    add_heading_3(doc, "3.2.4 บอร์ดจัดการพลังงานลดสัญญาณรบกวน (L3_Switching Board)")
    add_body_p(
        doc,
        "บอร์ดจ่ายพลังงาน L3_Switching รับแรงดันไฟฟ้าอินพุตจากภายนอก (เช่น 12V หรือ USB 5V) "
        "และทำการแปลงแรงดันผ่านวงจร Switching Regulator พร้อมวงจรกรอง LC Filter เพื่อสร้างแรงดันคงที่ 5.0 V และ 3.3 V "
        "จ่ายให้แก่ระบบ โดยมีการแยกสายดินระหว่างภาคดิจิทัลและแอนะล็อกอย่างเข้มงวด"
    )

    add_photo_placeholder(
        doc,
        "3.8",
        "แผนภาพวงจรบอร์ดจัดการพลังงาน L3_Switching",
        "ภาพ Schematic ของบอร์ดจ่ายไฟ L3_Switching แสดงวงจรเรกูเลเตอร์และวงจรกรอง LC",
        "Export ภาพวงจรจากไฟล์ L3_Switching.kicad_sch"
    )

    add_heading_2(doc, "3.3 สถาปัตยกรรมซอฟต์แวร์และอัลกอริทึมการชดเชยค่า (Software Calibration)")
    add_heading_3(doc, "3.3.1 ผังการทำงานของเฟิร์มแวร์ไมโครคอนโทรลเลอร์")
    add_body_p(
        doc,
        "เฟิร์มแวร์บนชิปประมวลผลถูกออกแบบให้ทำงานแบบ State Machine เริ่มต้นจากการกำหนดค่ารีจิสเตอร์ของ ADS1220, "
        "ทำการอ่านค่าแรงดันแสงและอุณหภูมิ NTC แบบต่อเนื่อง จากนั้นทำการกรองสัญญาณรบกวนความถี่สูงด้วย Digital Moving Average Filter "
        "และจัดรูปแบบข้อมูลเป็นโครงสร้าง JSON เพื่อส่งผ่านพอร์ตสื่อสาร"
    )

    add_photo_placeholder(
        doc,
        "3.9",
        "ผังงานการทำงานของโปรแกรมหลักบนไมโครคอนโทรลเลอร์ (Firmware Flowchart)",
        "Flowchart แสดงขั้นตอนการอ่านค่า ADC -> Digital Moving Average Filter -> ส่งออกข้อมูลผ่าน Serial",
        "วาดเป็นแผนผัง Flowchart มาตรฐานที่ชัดเจน"
    )

    add_heading_3(doc, "3.3.2 อัลกอริทึมคำนวณตำแหน่งดวงอาทิตย์และการชดเชยเงา (SPA & Shadow Inversion)")
    add_body_p(
        doc,
        "อัลกอริทึมชดเชยค่าทำงานบนพื้นฐานของแบบจำลองคณิตศาสตร์ โดยนำพิกัดภูมิศาสตร์และเวลาสากลมาคำนวณหา Solar Zenith Angle (θz) "
        "ตามมาตรฐาน NREL Solar Position Algorithm (SPA) จากนั้นนำมุม θz ไปเปิดตาราง Look-up Table เพื่อหาค่าสัมประสิทธิ์การบังแสงของกระป๋อง S_model(θz) "
        "และนำไปคูณหักล้างร่วมกับอุณหภูมิผิวเซนเซอร์ตามสมการรวม:"
    )
    add_body_p(doc, "E_GHI = I_measured / (R_0 · S_model(θz) · [1 + α_NTC · (T - T_0)])")

    add_photo_placeholder(
        doc,
        "3.10",
        "แผนผังการทำงานของอัลกอริทึม Software Cosine & Temperature Correction",
        "แผนผังแสดงอินพุต (Google Maps Lat/Long + วันเวลา) -> คำนวณ Solar Zenith Angle -> เปิดตาราง Inversion Model -> คูณแก้ค่ากระแสแสงและอุณหภูมิ NTC -> เอาต์พุตค่า GHI (W/m²)",
        "วาดเป็น Block Diagram แสดงขั้นตอนการประมวลผลทางคณิตศาสตร์"
    )

    add_heading_3(doc, "3.3.3 สถาปัตยกรรมเว็บแอปพลิเคชันและโครงร่างหน้าจอแสดงผล (Web UI Prototype)")
    add_body_p(
        doc,
        "ในโครงงานที่ 1 ซอฟต์แวร์ส่วนหน้าถูกจัดเตรียมเป็นโครงร่างระบบ (UI Mockup & Prototype Interface) "
        "โดยมีช่องสำหรับระบุพิกัดสถานที่ติดตั้งผ่านแผนที่ Google Maps, กราฟแสดงผลค่าความเข้มแสงและอุณหภูมิแบบ Real-time "
        "และระบบส่งออกข้อมูล (Data Export) ซึ่งจะได้รับการพัฒนาเป็นระบบ Cloud Telemetry สมบูรณ์โดยผู้รับผิดชอบฝั่งซอฟต์แวร์ในโครงงานที่ 2"
    )

    add_photo_placeholder(
        doc,
        "3.11",
        "📷 ภาพถ่ายหน้าจอส่วนติดต่อผู้ใช้เว็บแอปพลิเคชันและการปักหมุดพิกัด Google Maps",
        "ภาพแคปเจอร์หน้าจอ Web Application แสดงกราฟสัญญาณ, ช่องกรอกพิกัดละติจูด-ลองจิจูดจาก Google Maps, และมุมดวงอาทิตย์ที่คำนวณได้",
        "📷 แคปภาพหน้าจอเว็บแอปพลิเคชันขณะกำลังรันและแสดงผล"
    )

    add_heading_2(doc, "3.4 แผนงานและแนวทางแก้ปัญหาฮาร์ดแวร์ใน Project 2 (Hardware Evolution Roadmap)")
    add_body_p(
        doc,
        "เพื่อแก้ปัญหา Cosine Error ที่ต้นตอทางกายภาพอย่างถาวร ในโครงงานที่ 2 คณะผู้จัดทำได้วางแผนปรับเปลี่ยนแพ็กเกจเซนเซอร์ "
        "จากเดิมที่ใช้ BPX61 (TO-39 Metal Can) ไปสู่เซนเซอร์แบบแบนราบ Flat Surface-Mount Device (Vishay TEMD5010X01) "
        "ซึ่งมีคุณลักษณะเปรียบเทียบดังแสดงในตารางที่ 3.2:"
    )

    # Table 3.2 Comparison
    t_comp = doc.add_table(rows=6, cols=3)
    t_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_comp_headers = ["พารามิเตอร์เชิงวิศวกรรม", "เซนเซอร์ BPX61 (TO-39)", "เซนเซอร์ TEMD5010X01 (Flat SMD)"]
    for c_idx, h_text in enumerate(t_comp_headers):
        c = t_comp.cell(0, c_idx)
        set_cell_shading(c, "E2E8F0")
        run = c.paragraphs[0].add_run(h_text)
        run.font.name = 'TH SarabunPSK'
        run.font.bold = True

    comp_rows = [
        ("ลักษณะแพ็กเกจ (Package Style)", "กระป๋องโลหะ TO-39 พร้อมกระจกครอบ", "SMD แบนราบพิเศษ (Ultra-thin Flat Package)"),
        ("ความสูงเหนือแผ่นวงจร (Profile Height)", "ประมาณ 4.5 มิลลิเมตร (ขอบลึก 1.5–2 mm)", "1.12 มิลลิเมตร (ไม่มีขอบบัง)"),
        ("พื้นที่รับแสงไวต่อแสง (Radiant Sensitive Area)", "7.02 mm²", "7.5 mm²"),
        ("การบดบังแสงมุมเอียง (Aperture Shadowing)", "เกิดเงาที่ θ > 40° และตัดแสงที่ 70°", "ไม่มีขอบกระป๋อง แสงตกกระทบผิวได้ถึงมุม > 80°"),
        ("แนวทางการสอบเทียบและปรับปรุง", "ต้องใช้ Software Cosine Inversion ชดเชย", "ตอบสนองใกล้เคียง Cosine Law ตามธรรมชาติ")
    ]
    for r_idx, row_data in enumerate(comp_rows, start=1):
        for c_idx, val in enumerate(row_data):
            c = t_comp.cell(r_idx, c_idx)
            run = c.paragraphs[0].add_run(val)
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(14)

    add_photo_placeholder(
        doc,
        "3.12",
        "โครงสร้างเปรียบเทียบเซนเซอร์ตัวถังกระป๋อง TO-39 กับ Flat SMD (TEMD5010X01)",
        "ภาพเปรียบเทียบโครงสร้างเชิงกล: ด้านซ้ายแสดงขอบกระป๋อง TO-39 ที่ทำให้เกิดเงาหลุมบัง Die / ด้านขวาแสดงเซนเซอร์ Flat SMD หนาเพียง 1.12 มม. ที่รับแสงได้รอบทิศทาง 180°",
        "วาดเป็นไดอะแกรม Cross-section View เปรียบเทียบมุมตกกระทบของแสง"
    )

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 4: ผลการดำเนินงานและการวิเคราะห์ผล (Results & Discussion)
    # =========================================================================
    add_heading_1(doc, "บทที่ 4\nผลการดำเนินงานและการวิเคราะห์ผล\n(Results & Discussion)")

    add_heading_2(doc, "4.1 ผลการทดสอบวงจร Analog AFE บน Breadboard และ Perfboard")
    add_body_p(
        doc,
        "จากการทดสอบวงจรภาคแอนะล็อก L1_Analog สามารถสรุปผลการทดลองและประเด็นสำคัญได้ดังนี้:"
    )
    add_bullet_p(doc, "1. การแก้ปัญหา High-Z Leakage:", 
                 "การทดสอบบน Breadboard ส่งผลให้แรงดันอ้างอิงของ REF5045 ตกลงเหลือ 4.382 V (ตกไป 118 mV) จากกระแสรั่วไหลบนรางเบรดบอร์ด "
                 "เมื่อย้ายวงจรมาประกอบลงบนแผงวงจรอเนกประสงค์ (Perfboard) และเว้นว่าง (DNP) ขา C11 แรงดันอ้างอิงกลับมานิ่งที่ 4.500 V อย่างสมบูรณ์")
    add_bullet_p(doc, "2. ความเสถียรของลูปป้อนกลับ:", 
                 "การติดตั้งตัวเก็บประจุ C10 ขนาด 10 µF ชนิด Low-ESR ชิดขาเอาต์พุตของ REF5045 สามารถขจัดปัญหาการสั่นพวง (Oscillation) ที่ความถี่ 250 kHz ได้อย่างเด็ดขาด "
                 "ทำให้ได้ระดับสัญญาณรบกวน (Noise Floor) ต่ำกว่า 5 µVp-p")
    add_bullet_p(doc, "3. ความเป็นเชิงเส้นของสัญญาณกระแสแสง:", 
                 "วงจร Transimpedance ขยายสัญญาณกระแสแสงจาก BPX61 เข้าสู่ ADS1220 ได้อย่างเป็นเชิงเส้น โดยมีค่าความละเอียดในการแปลงสัญญาณสูงถึง 24 บิต")

    add_photo_placeholder(
        doc,
        "4.1",
        "กราฟเปรียบเทียบเสถียรภาพแรงดันเอาต์พุตของการต่อ C10 ชิดขา REF5045",
        "กราฟสัญญาณจากออสซิลโลสโคปเปรียบเทียบระหว่างกรณีไม่มี C10 (เกิดสัญญาณแกว่ง 250 kHz) กับกรณีมี C10 10 µF (แรงดันคงที่ 4.500 V สัญญาณเรียบ)",
        "นำภาพแคปเจอร์รูปคลื่นจากออสซิลโลสโคปหรือกราฟผลการทดลอง"
    )

    add_heading_2(doc, "4.2 ผลการตรวจวัดและชดเชยอุณหภูมิด้วย NTC Thermistor")
    add_body_p(
        doc,
        "การทดสอบอ่านค่าความต้านทานของ NTC Thermistor 10 kΩ ร่วมกับสมการ Steinhart-Hart พบว่าสามารถตรวจวัดอุณหภูมิผิวของเซนเซอร์ "
        "ได้ในช่วง 25°C ถึง 65°C โดยมีความละเอียดในการวัดดีกว่า ±0.1°C ซึ่งเพียงพอต่อการนำไปคำนวณชดเชยค่าความเข้มแสงตามสัมประสิทธิ์อุณหภูมิ (α = +0.18%/°C) "
        "ช่วยลดความคลาดเคลื่อนจากความร้อนสะสมของแสงอาทิตย์ได้ทั้งหมด"
    )

    add_heading_2(doc, "4.3 ผลการจำลองพฤติกรรมเชิงแสง (Optical Ray-Tracing Analysis)")
    add_body_p(
        doc,
        "ผลจากการทำ Ray-tracing Optical Simulation บนแบบจำลอง 3 มิติของตัวถัง TO-39 แสดงให้เห็นอย่างชัดเจนว่า:"
    )
    add_bullet_p(doc, "ช่วงมุมตกกระทบ 0° ถึง 40°:", "ค่าความเข้มแสงที่ตกกระทบลงบนผิว Die สอดคล้องกับฟังก์ชันโคไซน์อุดมคติ (Ideal Cosine Law) โดยมีความคลาดเคลื่อนต่ำกว่า 2.5%")
    add_bullet_p(doc, "ช่วงมุมตกกระทบ 40° ถึง 65°:", "ขอบโลหะของตัวถังกระป๋องเริ่มทอดเงาบดบังผิวรับแสงรูปสี่เหลี่ยม (Progressive Shadowing) ทำให้ค่าแสงดรอปลงอย่างรวดเร็ว ส่งผลให้ความคลาดเคลื่อนโคไซน์พุ่งสูงขึ้นถึง 25–40%")
    add_bullet_p(doc, "ช่วงมุมตกกระทบ ≥ 70°:", "เงาของขอบกระป๋องทอดบังผิวรับแสงเกือบ 100% ทำให้เกิดจุดตัดแสงสิ้นเชิง (Total Optical Cutoff)")

    add_photo_placeholder(
        doc,
        "4.2",
        "กราฟความคลาดเคลื่อนโคไซน์จากการจำลองเงาเทียบกับฟังก์ชัน Ideal Cosine",
        "กราฟแสดงเปอร์เซ็นต์ความคลาดเคลื่อน Cosine Error (%) เทียบกับมุมตกกระทบ θz แสดงจุดวิกฤตที่เริ่มคลาดเคลื่อนเกินเกณฑ์มาตรฐาน ISO 9060",
        "นำภาพกราฟ Error Curve จากผลการจำลอง Simulation"
    )

    add_heading_2(doc, "4.4 ผลการทดสอบส่งสัญญาณข้อมูลดิบและโครงร่างส่วนติดต่อผู้ใช้")
    add_body_p(
        doc,
        "ไมโครคอนโทรลเลอร์ ESP32 สามารถอ่านค่าจาก ADS1220 และส่งผ่านข้อมูล Raw Counts ทั้งช่องสัญญาณแสงและอุณหภูมิออกมาทางพอร์ต USB Serial "
        "ได้อย่างต่อเนื่องที่อัตรา 20 SPS โดยไม่มีปัญหาข้อมูลสูญหาย และสามารถเชื่อมต่อไปยังโครงร่างหน้าจอ Web UI Prototype "
        "เพื่อเตรียมความพร้อมสำหรับการพัฒนาระบบคลาวด์แดชบอร์ดฉบับสมบูรณ์ในโครงงานที่ 2"
    )

    add_photo_placeholder(
        doc,
        "4.3",
        "📷 ภาพหน้าจอการแสดงผลข้อมูลดิบผ่านพอร์ตสื่อสารอนุกรม (Serial Monitor)",
        "ภาพแคปเจอร์หน้าต่าง Serial Monitor หรือ Terminal แสดงข้อมูล Raw Counts ของแสงและอุณหภูมิที่ส่งออกมาอย่างสม่ำเสมอ",
        "📷 แคปภาพหน้าจอ Serial Monitor ขณะกำลังอ่านค่าจากบอร์ด Perfboard จริง"
    )

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 5: สรุปผลการดำเนินงาน ปัญหา และแผนงานระยะถัดไป
    # =========================================================================
    add_heading_1(doc, "บทที่ 5\nสรุปผลการดำเนินงาน ปัญหา และแผนงานระยะถัดไป\n(Conclusion & Future Work)")

    add_heading_2(doc, "5.1 สรุปผลความก้าวหน้าของโครงงานที่ 1")
    add_body_p(
        doc,
        "ในการดำเนินงานโครงงาน Project 1 คณะผู้จัดทำได้บรรลุเป้าหมายสำคัญตามวัตถุประสงค์ที่ตั้งไว้ ดังนี้:"
    )
    add_bullet_p(doc, "1. การพิสูจน์การทำงานของวงจร AFE:", "สามารถออกแบบและประกอบวงจรภาคปรับสภาพสัญญาณแอนะล็อกความไวสูงและไอซีแปลงสัญญาณ 24 บิต (ADS1220) ลงบนแผง Perfboard ได้สำเร็จ โดยขจัดปัญหา High-Z Leakage และ Loop Oscillation ได้อย่างสมบูรณ์")
    add_bullet_p(doc, "2. การค้นพบทางทัศนศาสตร์และการจำลอง:", "สามารถระบุสาเหตุที่แท้จริงของ Cosine Error ว่าเกิดจากขอบตัวถัง TO-39 พร้อมทั้งวิเคราะห์ข้อจำกัดของวัสดุ PTFE จนนำไปสู่การวางแผนแก้ปัญหาอย่างเป็นระบบ")
    add_bullet_p(doc, "3. สถาปัตยกรรมระบบแบบบูรณาการ:", "วางรากฐานการชดเชยค่าด้วยซอฟต์แวร์ (Software Cosine Correction) ร่วมกับตำแหน่งดวงอาทิตย์ (SPA) และโครงร่าง Web UI เพื่อส่งต่อการพัฒนาสู่ Project 2")

    add_heading_2(doc, "5.2 ปัญหา อุปสรรค และบทเรียนที่ได้รับ (Lessons Learned)")
    add_body_p(
        doc,
        "การดำเนินงานในโครงงานที่ 1 ได้ให้บทเรียนเชิงวิศวกรรมและการวิจัยที่สำคัญยิ่ง:"
    )
    add_bullet_p(doc, "1. การคัดเลือกแพ็กเกจเซนเซอร์มีความสำคัญเท่ากับคุณสมบัติทางไฟฟ้า:", "แม้เซนเซอร์ BPX61 จะมี Linearity และ Dark Current ที่ยอดเยี่ยม แต่โครงสร้างทางกลของตัวถังกระป๋องโลหะกลับเป็นอุปสรรคสำคัญต่อการวัดมุมกว้าง การเลือกอุปกรณ์จึงต้องพิจารณาทั้งมิติทางไฟฟ้าและมิติทางกลควบคู่กัน")
    add_bullet_p(doc, "2. การประเมินความเป็นไปได้เชิงวิศวกรรม (Engineering Feasibility):", "การตระหนักรู้ว่าการขึ้นรูปโดม PTFE เองในแล็บมีความเสี่ยงสูงเกินไป และตัดสินใจปรับกลยุทธ์มาใช้การชดเชยด้วยซอฟต์แวร์ควบคู่กับการเปลี่ยนแพ็กเกจเซนเซอร์ ช่วยป้องกันความล้มเหลวของโครงงานและทำให้การดำเนินงานเดินหน้าได้อย่างมั่นคง")
    add_bullet_p(doc, "3. สัญญาณรบกวนในวงจรแอนะล็อกระดับไมโครโวลต์:", "การทดสอบบน Breadboard ชี้ให้เห็นถึงความเปราะบางของจุด High-Z และความสำคัญอย่างยิ่งยวดของตัวเก็บประจุ Low-ESR Decoupling ในการควบคุมเสถียรภาพของวงจร")

    add_heading_2(doc, "5.3 แผนการดำเนินงานสำหรับโครงงานที่ 2 (Project 2 Implementation Plan)")
    add_body_p(
        doc,
        "ในภาคการศึกษาที่ 2 (Project 2) คณะผู้จัดทำได้วางแผนการดำเนินงานวิจัยและทดสอบเพื่อสร้างระบบที่สมบูรณ์ตามตารางที่ 5.1:"
    )

    # Table 5.1 Project 2 Roadmap
    t_p2 = doc.add_table(rows=6, cols=4)
    t_p2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_p2_headers = ["ระยะเวลา", "กิจกรรมการดำเนินงาน", "ผู้รับผิดชอบหลัก", "ผลผลิตที่คาดหวัง"]
    for c_idx, h_text in enumerate(t_p2_headers):
        c = t_p2.cell(0, c_idx)
        set_cell_shading(c, "E2E8F0")
        run = c.paragraphs[0].add_run(h_text)
        run.font.name = 'TH SarabunPSK'
        run.font.bold = True

    p2_rows = [
        (
            "พ.ย. 2569",
            "ออกแบบและสั่งผลิต Custom 4-Layer PCB สำหรับเซนเซอร์ Flat SMD (TEMD5010X01)",
            "ทชินทร์ รักษาวงศ์ (Hardware Lead)",
            "แผ่นวงจรพิมพ์ PCB 4 เลเยอร์มาตรฐานโรงงานอุตสาหกรรม"
        ),
        (
            "พ.ย. - ธ.ค. 2569",
            "พัฒนาระบบ Cloud Telemetry, ฐานข้อมูล และ Web Application Dashboard สมบูรณ์",
            "(ชื่อเพื่อนคนที่ 2) (Software Lead)",
            "Web App Dashboard พร้อมระบบเชื่อมต่อ Google Maps และ SPA อัตโนมัติ"
        ),
        (
            "ม.ค. 2570",
            "ทดสอบกวาดมุมองศา (Angular Test) ด้วยแท่น Goniometer ในห้องปฏิบัติการทัศนศาสตร์ อาคารจุฬาภรณ์ 1 ชั้น 3",
            "ทชินทร์ รักษาวงศ์ ร่วมกับ (ชื่อเพื่อนคนที่ 2)",
            "ข้อมูลเปรียบเทียบ A/B Testing ระหว่าง BPX61 (TO-39) กับ TEMD5010X01 (SMD)"
        ),
        (
            "ม.ค. - ก.พ. 2570",
            "ทดสอบสอบเทียบค่าความเข้มแสงกับเครื่องจำลองแสงอาทิตย์ Solar Light XPS-300 อาคาร 55 ชั้น 3",
            "คณะผู้จัดทำ",
            "กราฟ Calibration Curve และรายงานค่าความไม่แน่นอนของการวัด (Uncertainty)"
        ),
        (
            "ก.พ. 2570",
            "จัดทำรายงานปริญญานิพนธ์ฉบับสมบูรณ์ และเตรียมสอบป้องกันปริญญานิพนธ์",
            "คณะผู้จัดทำ",
            "เล่มปริญญานิพนธ์ฉบับสมบูรณ์ และผลงานวิจัยพร้อมเผยแพร่"
        )
    ]
    for r_idx, row_data in enumerate(p2_rows, start=1):
        for c_idx, val in enumerate(row_data):
            c = t_p2.cell(r_idx, c_idx)
            run = c.paragraphs[0].add_run(val)
            run.font.name = 'TH SarabunPSK'
            run.font.size = Pt(14)

    doc.add_page_break()

    # =========================================================================
    # REFERENCES (บรรณานุกรม)
    # =========================================================================
    add_heading_1(doc, "บรรณานุกรม")

    references = [
        "Texas Instruments, \"ADS1220: 4-Channel, 2-kSPS, 24-Bit, ADC with Integrated PGA and Voltage Reference,\" Datasheet SBAS501C, Revised 2021.",
        "Texas Instruments, \"REF50xx: Low-Noise, Very Low Drift, Precision Voltage Reference,\" Datasheet SBOS410J, Revised 2022.",
        "Semitec Corporation, \"AT-Thermistor: High-Precision NTC Thermistor (103AT),\" Technical Datasheet, 2022.",
        "ams-OSRAM, \"BPX61: Silicon Photodiode for High Precision Photometry in TO-39 Metal Can,\" Datasheet Version 1.6, 2023.",
        "Vishay Semiconductors, \"TEMD5010X01: High Speed Silicon PIN Photodiode (Flat Miniature SMD Package),\" Document Number 81507, Rev. 1.8, 2022.",
        "National Renewable Energy Laboratory (NREL), \"Solar Position Algorithm for Solar Radiation Applications (NREL/TP-560-34302),\" Golden, Colorado, 2008.",
        "International Organization for Standardization, \"ISO 9060:2018: Solar Energy - Specification and Classification of Instruments for Measuring Hemispherical Solar and Direct Solar Radiation,\" 2018.",
        "World Meteorological Organization (WMO), \"Guide to Instruments and Methods of Observation (WMO-No. 8),\" Chapter 7: Measurement of Radiation, 2021.",
        "Joint Committee for Guides in Metrology (JCGM), \"Evaluation of Measurement Data - Guide to the Expression of Uncertainty in Measurement (GUM),\" JCGM 100:2008, 2008.",
        "สถาบันมาตรวิทยาแห่งชาติ (มว.), \"หลักการวัดและการสอบกลับได้ของผลการวัดในงานวิศวกรรม,\" เอกสารเผยแพร่ทางวิชาการ, 2567."
    ]

    for ref in references:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.left_indent = Inches(0.5)
        p_ref.paragraph_format.first_line_indent = Inches(-0.5)
        p_ref.paragraph_format.space_after = Pt(6)
        r_ref = p_ref.add_run(ref)
        r_ref.font.name = 'TH SarabunPSK'
        r_ref.font.size = Pt(14)

    target_docx = r"c:\Users\usEr\Documents\KiCad\10.0\projects\Pyranometer_Over_engineering\Report_doc\69_IoT_Project1_Report.docx"
    try:
        doc.save(target_docx)
        print(f"Updated Report Document successfully generated at: {target_docx}")
    except PermissionError:
        fallback_docx = r"c:\Users\usEr\Documents\KiCad\10.0\projects\Pyranometer_Over_engineering\Report_doc\69_IoT_Project1_Report_Updated.docx"
        doc.save(fallback_docx)
        print(f"[Word is open] Saved updated document to: {fallback_docx}")


if __name__ == "__main__":
    create_report()
