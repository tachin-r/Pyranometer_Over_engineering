"""
โครงสร้างสร้างรายงานฉบับสมบูรณ์ตามระเบียบ สจล. (KMITL Thesis / Capstone Standard Format)
พัฒนาขึ้นสำหรับโครงงาน: การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง (Pyranometer)
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT

# นำเข้าเครื่องมือจัดหน้ามาตรฐานจาก helpers.py
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
    add_figure_image,
    add_photo_placeholder
)

def build_kmitl_standard_report(output_filename="Pyranometer_KMITL_Report.docx"):
    doc = docx.Document()

    # ตั้งค่าสไตล์ Normal เบื้องต้น
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'TH SarabunPSK'
    normal_style.font.size = Pt(16)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)

    # =========================================================================
    # SECTION 1: ปกนอก และ ปกใน (ไม่มีเลขหน้า)
    # =========================================================================
    sec1 = doc.sections[0]
    setup_kmitl_page(sec1)

    # 1.1 หน้าปก (Cover Page)
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(30)

    p_th_title = doc.add_paragraph()
    p_th_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_th_title.paragraph_format.space_after = Pt(8)
    p_th_title.paragraph_format.line_spacing = 1.15
    r_th = p_th_title.add_run(
        "การพัฒนาระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูง\n"
        "โดยใช้ตัวแปลงแอนะล็อกเป็นดิจิทัล 24 บิต และการประเมินทางมาตรวิทยา"
    )
    set_font_sarabun(r_th, size_pt=20, bold=True)

    p_en_title = doc.add_paragraph()
    p_en_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_en_title.paragraph_format.space_after = Pt(70)
    p_en_title.paragraph_format.line_spacing = 1.15
    r_en = p_en_title.add_run(
        "DEVELOPMENT OF A HIGH-PRECISION SOLAR IRRADIANCE MEASUREMENT SYSTEM\n"
        "USING 24-BIT ANALOG-TO-DIGITAL CONVERTER AND METROLOGICAL EVALUATION"
    )
    set_font_sarabun(r_en, size_pt=18, bold=True)

    p_authors = doc.add_paragraph()
    p_authors.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_authors.paragraph_format.space_after = Pt(70)
    p_authors.paragraph_format.line_spacing = 1.2
    r_auth = p_authors.add_run(
        "เตชินท์ รัฐวงษา          รหัสนักศึกษา 66010296\n"
        "(ชื่อผู้ร่วมโครงงานคนที่ 2)   รหัสนักศึกษา 6601xxxx"
    )
    set_font_sarabun(r_auth, size_pt=16, bold=True)

    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.line_spacing = 1.15
    r_foot = p_foot.add_run(
        "รายงานโครงงานนี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาบัณฑิต\n"
        "หลักสูตรสองปริญญา (Dual Degree):\n"
        "วิศวกรรมศาสตรบัณฑิต (วิศวกรรมระบบไอโอทีและสารสนเทศ)\n"
        "และ วิทยาศาสตรบัณฑิต (ฟิสิกส์อุตสาหกรรม)\n"
        "คณะวิศวกรรมศาสตร์\n"
        "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง\n"
        "ปีการศึกษา 2569"
    )
    set_font_sarabun(r_foot, size_pt=16, bold=True)

    # =========================================================================
    # SECTION 2: ส่วนนำ (บทคัดย่อ, กิตติกรรมประกาศ, สารบัญ) -> เลขหน้าโรมัน i, ii, iii...
    # =========================================================================
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec2)
    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False
    set_section_numbering(sec2, fmt="lowerRoman", start=1)
    add_page_number_to_footer(sec2.footer)

    # 2.1 บทคัดย่อภาษาไทย
    p_ab_th = doc.add_paragraph()
    p_ab_th.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ab_th.paragraph_format.space_before = Pt(10)
    p_ab_th.paragraph_format.space_after = Pt(12)
    r_ab_th = p_ab_th.add_run("บทคัดย่อ")
    set_font_sarabun(r_ab_th, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "โครงงานนี้นำเสนอการวิจัย ออกแบบ และสร้างระบบวัดค่าความเข้มการแผ่รังสีแสงอาทิตย์ (Pyranometer) "
        "ที่มีความแม่นยำและความละเอียดสูงระดับเครื่องมืออ้างอิงทางมาตรวิทยา โดยใช้เซนเซอร์โฟโตไดโอดชนิดซิลิคอน "
        "ร่วมกับวงจรแปลงสัญญาณแอนะล็อกเป็นดิจิทัลความละเอียด 24 บิต (ADS1220) และแหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูงพิเศษ "
        "(REF5045) เพื่อขจัดสัญญาณรบกวนและลดผลกระทบจากอุณหภูมิ พร้อมทั้งดำเนินการสอบเทียบด้วยเครื่องกำลองแสงอาทิตย์ "
        "(Solar Simulator รุ่น Solar Light XPS-300 Class AAA) ตามมาตรฐาน ISO 9060 และหลักการประเมินความไม่แน่นอนของการวัด (GUM)"
    )

    # 2.2 บทคัดย่อภาษาอังกฤษ (ขึ้นหน้าใหม่)
    doc.add_page_break()
    p_ab_en = doc.add_paragraph()
    p_ab_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ab_en.paragraph_format.space_before = Pt(10)
    p_ab_en.paragraph_format.space_after = Pt(12)
    r_ab_en = p_ab_en.add_run("ABSTRACT")
    set_font_sarabun(r_ab_en, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "This project presents the design, development, and metrological characterization of a high-precision "
        "solar pyranometer system. The system integrates a silicon photodiode sensor with a 24-bit delta-sigma "
        "ADC (Texas Instruments ADS1220) and an ultra-low-drift voltage reference (REF5045). Calibration is carried "
        "out under a Class AAA Solar Simulator (Solar Light XPS-300) in accordance with ISO 9060 and the Guide to "
        "the Expression of Uncertainty in Measurement (GUM)."
    )

    # 2.3 กิตติกรรมประกาศ (ขึ้นหน้าใหม่)
    doc.add_page_break()
    p_ack = doc.add_paragraph()
    p_ack.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ack.paragraph_format.space_before = Pt(10)
    p_ack.paragraph_format.space_after = Pt(12)
    r_ack = p_ack.add_run("กิตติกรรมประกาศ")
    set_font_sarabun(r_ack, size_pt=18, bold=True)

    add_thai_body_p(
        doc,
        "คณะผู้ดำเนินโครงงานขอขอบพระคุณอาจารย์ที่ปรึกษาโครงงาน คณาจารย์ภาควิชาวิศวกรรมระบบไอโอทีและสารสนเทศ "
        "และภาควิชาฟิสิกส์อุตสาหกรรม คณะวิศวกรรมศาสตร์ สจล. ที่ได้ให้คำแนะนำ แนวคิด ตลอดจนการสนับสนุนเครื่องมือ "
        "รวมทั้งขอขอบคุณศูนย์ทดสอบและวิเคราะห์เทคโนโลยีขั้นสูง (ATTAC) และพี่วิษา หอมจันทร์ ที่อนุเคราะห์การเข้าใช้เครื่อง "
        "Solar Simulator (Solar Light XPS-300) จนทำให้งานวิจัยนี้ประสบความสำเร็จลุล่วงด้วยดี"
    )

    # 2.4 สารบัญ (Table of Contents)
    doc.add_page_break()
    p_toc_head = doc.add_paragraph()
    p_toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_toc_head.paragraph_format.space_before = Pt(10)
    p_toc_head.paragraph_format.space_after = Pt(12)
    r_toc = p_toc_head.add_run("สารบัญ")
    set_font_sarabun(r_toc, size_pt=18, bold=True)

    p_col = doc.add_paragraph()
    p_col.paragraph_format.tab_stops.add_tab_stop(Inches(5.75), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
    r_c1 = p_col.add_run("เรื่อง")
    set_font_sarabun(r_c1, size_pt=16, bold=True)
    r_c2 = p_col.add_run("\tหน้า")
    set_font_sarabun(r_c2, size_pt=16, bold=True)

    add_toc_line(doc, "บทคัดย่อภาษาไทย", "i")
    add_toc_line(doc, "บทคัดย่อภาษาอังกฤษ", "ii")
    add_toc_line(doc, "กิตติกรรมประกาศ", "iii")
    add_toc_line(doc, "สารบัญ", "iv", is_bold=True)
    add_toc_line(doc, "สารบัญรูป", "v")
    add_toc_line(doc, "สารบัญตาราง", "vi")
    add_toc_line(doc, "บทที่ 1 บทนำ", "1", is_bold=True)
    add_toc_line(doc, "1.1 ความเป็นมาและความสำคัญของปัญหา", "1", level=1)
    add_toc_line(doc, "1.2 วัตถุประสงค์ของโครงงาน", "2", level=1)
    add_toc_line(doc, "1.3 ขอบเขตของโครงงาน", "2", level=1)
    add_toc_line(doc, "1.4 ประโยชน์ที่คาดว่าจะได้รับ", "3", level=1)
    add_toc_line(doc, "บทที่ 2 ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "4", is_bold=True)
    add_toc_line(doc, "2.1 หลักการวัดการแผ่รังสีดวงอาทิตย์และมาตรฐาน ISO 9060", "4", level=1)
    add_toc_line(doc, "2.2 วงจรขยายและแปลงสัญญาณ 24-bit ADC (ADS1220)", "6", level=1)
    add_toc_line(doc, "2.3 แหล่งจ่ายแรงดันอ้างอิงความแม่นยำสูง REF5045", "9", level=1)
    add_toc_line(doc, "2.4 เครื่องจำลองแสงอาทิตย์ Solar Light XPS-300 Class AAA", "12", level=1)
    add_toc_line(doc, "บทที่ 3 การออกแบบและวิธีการดำเนินงาน", "15", is_bold=True)
    add_toc_line(doc, "3.1 สถาปัตยกรรมรวมของระบบวัด Pyranometer", "15", level=1)
    add_toc_line(doc, "3.2 การออกแบบวงจรฮาร์ดแวร์ (Schematic & PCB)", "17", level=1)
    add_toc_line(doc, "3.3 ขั้นตอนและวิธีทดสอบด้วย Solar Simulator", "22", level=1)
    add_toc_line(doc, "บทที่ 4 ผลการทดลองและการวิเคราะห์ผล", "26", is_bold=True)
    add_toc_line(doc, "4.1 ผลการทดสอบคุณลักษณะการตอบสนองเชิงเส้น (Linearity)", "26", level=1)
    add_toc_line(doc, "4.2 การประเมินงบประมาณความไม่แน่นอนของการวัด (GUM)", "29", level=1)
    add_toc_line(doc, "บทที่ 5 สรุปผลการดำเนินงานและข้อเสนอแนะ", "33", is_bold=True)
    add_toc_line(doc, "เอกสารอ้างอิง", "35", is_bold=True)
    add_toc_line(doc, "ภาคผนวก", "37", is_bold=True)

    # =========================================================================
    # SECTION 3: เนื้อหาหลัก บทที่ 1 ถึง 5 (เลขหน้าอารบิก 1, 2, 3...)
    # =========================================================================
    sec3 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_kmitl_page(sec3)
    sec3.header.is_linked_to_previous = False
    sec3.footer.is_linked_to_previous = False
    set_section_numbering(sec3, fmt="decimal", start=1)
    add_page_number_to_footer(sec3.footer)

    # -------------------------------------------------------------------------
    # บทที่ 1: บทนำ
    # -------------------------------------------------------------------------
    add_thesis_chapter_heading(doc, "1", "บทนำ")

    add_section_heading(doc, "1.1", "ความเป็นมาและความสำคัญของปัญหา")
    add_thai_body_p(
        doc,
        "พลังงานแสงอาทิตย์เป็นแหล่งพลังงานหมุนเวียนที่มีบทบาทสำคัญอย่างยิ่งต่อการเปลี่ยนผ่านสู่พลังงานสะอาด "
        "การประเมินประสิทธิภาพของระบบผลิตไฟฟ้าด้วยเซลล์แสงอาทิตย์ (Photovoltaic Systems) และการศึกษาสภาพภูมิอากาศ "
        "จำเป็นต้องอาศัยการตรวจวัดค่าความเข้มการแผ่รังสีดวงอาทิตย์ (Solar Irradiance, W/m²) ที่มีความแม่นยำสูง "
        "เครื่องวัดรังสีดวงอาทิตย์ระดับ Class A หรือ Secondary Standard ตามมาตรฐาน ISO 9060 ในท้องตลาดมีราคาสูงมาก "
        "ทำให้เป็นข้อจำกัดในการติดตั้งสถานีวัดแบบกระจายตัว"
    )
    add_thai_body_p(
        doc,
        "โครงงานนี้จึงมุ่งเน้นการออกแบบและสร้างระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูงแบบโอเวอร์เอนจิเนียริง "
        "(Over-engineering Pyranometer) โดยใช้ส่วนประมวลผลสัญญาณแอนะล็อกระดับห้องปฏิบัติการ ผสานชิป 24-bit ADC "
        "และตัวกำเนิดแรงดันอ้างอิงที่มีค่าสัมประสิทธิ์การเลื่อนตามอุณหภูมิต่ำมาก เพื่อให้ได้ผลการวัดที่เที่ยงตรงเทียบเท่ามาตรฐานสากล"
    )

    add_section_heading(doc, "1.2", "วัตถุประสงค์ของโครงงาน")
    add_numbered_item(doc, "1.", "เพื่อออกแบบและสร้างวงจรตรวจวัดรังสีแสงอาทิตย์ความละเอียดสูงโดยใช้ ADS1220 24-bit ADC", bold_prefix="การออกแบบฮาร์ดแวร์: ")
    add_numbered_item(doc, "2.", "เพื่อบูรณาการแหล่งกำเนิดแรงดันอ้างอิง REF5045 ลด Drift จากอุณหภูมิ", bold_prefix="เสถียรภาพการวัด: ")
    add_numbered_item(doc, "3.", "เพื่อสอบเทียบและประเมินความไม่แน่นอนของการวัด (Uncertainty Analysis) ร่วมกับ Solar Simulator", bold_prefix="การสอบเทียบทางมาตรวิทยา: ")

    add_section_heading(doc, "1.3", "ขอบเขตของโครงงาน")
    add_bullet_item(doc, bullet_char="•", text="ช่วงการวัดค่าความเข้มรังสีแสงอาทิตย์ครอบคลุม 0 ถึง 1,500 W/m²")
    add_bullet_item(doc, bullet_char="•", text="การสุ่มตัวอย่างสัญญาณผ่าน ADS1220 อัตราการสุ่ม 20 SPS เพื่อขจัดสัญญาณรบกวน 50 Hz/60 Hz")
    add_bullet_item(doc, bullet_char="•", text="ทดสอบสอบเทียบเทียบกับเครื่องกำเนิดแสงจำลอง Solar Light XPS-300 Class AAA ณ ศูนย์ ATTAC สจล.")

    # -------------------------------------------------------------------------
    # บทที่ 2: ทฤษฎีและงานวิจัยที่เกี่ยวข้อง
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "2", "ทฤษฎีและงานวิจัยที่เกี่ยวข้อง")

    add_section_heading(doc, "2.1", "หลักการวัดการแผ่รังสีดวงอาทิตย์และมาตรฐาน ISO 9060")
    add_thai_body_p(
        doc,
        "มาตรฐานสากล ISO 9060:2018 จำแนกประเภทของ Pyranometer ออกเป็น Class A, Class B และ Class C "
        "โดยพิจารณาจากพารามิเตอร์สำคัญ ได้แก่ เวลาการตอบสนอง (Response Time), ความไม่เป็นเชิงเส้น (Non-linearity), "
        "และการตอบสนองต่อทิศทางของแสง (Cosine Response)"
    )

    add_section_heading(doc, "2.2", "วงจรแปลงสัญญาณ 24-bit Delta-Sigma ADC (ADS1220)")
    add_thai_body_p(
        doc,
        "ชิป ADS1220 เป็นตัวแปลงแอนะล็อกเป็นดิจิทัลความละเอียด 24 บิต ที่มี Low-Noise Programmable Gain Amplifier (PGA) "
        "ในตัว รองรับการวัดแบบ Differential Input ทำให้สามารถวัดแรงดันระดับไมโครโวลต์จากตัวตรวจจับโฟโตไดโอดได้อย่างแม่นยำสูง"
    )

    add_photo_placeholder(
        doc,
        fig_no="2.1",
        title="บล็อกไดอะแกรมภายในของไอซี ADS1220 24-bit ADC",
        description="แสดงส่วน MUX สลับช่องสัญญาณ, วงจร PGA, ตัวแปลงสัญญาณ Delta-Sigma และอินเทอร์เฟซ SPI"
    )

    # -------------------------------------------------------------------------
    # บทที่ 3: การออกแบบและวิธีการดำเนินงาน
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "3", "การออกแบบและวิธีการดำเนินงาน")

    add_section_heading(doc, "3.1", "สถาปัตยกรรมรวมของระบบ")
    add_thai_body_p(
        doc,
        "ระบบประกอบด้วยเซนเซอร์ตรวจจับรังสี, วงจรจำกัดสัญญาณรบกวน, บอร์ดขยายและแปลงสัญญาณ 24-bit ADC, "
        "และไมโครคอนโทรลเลอร์ที่ประมวลผลข้อมูลส่งผ่านโปรโตคอลดิจิทัล"
    )

    add_photo_placeholder(
        doc,
        fig_no="3.1",
        title="ภาพการประกอบบอร์ดเซนเซอร์ Pyranometer (Perfboard Prototype)",
        description="บอร์ดทดสอบต้นแบบที่มีการจัดวางตำแหน่งโมดูล ADS1220, REF5045 และจุดต่อสัญญาณ Photodiode"
    )

    # -------------------------------------------------------------------------
    # บทที่ 4: ผลการทดลองและการวิเคราะห์ผล
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "4", "ผลการทดลองและการวิเคราะห์ผล")

    add_section_heading(doc, "4.1", "ผลการทดสอบคุณลักษณะการตอบสนองเชิงเส้น")
    add_thai_body_p(
        doc,
        "จากการทดสอบนำหัววัดไปรับแสงที่ระดับความเข้มต่าง ๆ จากเครื่อง Solar Simulator Solar Light XPS-300 "
        "ได้ผลการวิเคราะห์การตอบสนองเชิงเส้นและสมการถดถอยดังนี้"
    )

    # ตารางตัวอย่างพร้อมเส้นขอบสะอาดตาแบบวิทยานิพนธ์
    table = doc.add_table(rows=4, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_clean_borders(table, border_color="000000", border_sz="4")

    headers = ["จุดทดสอบ", "ความเข้มแสงอ้างอิง (W/m²)", "แรงดันเอาต์พุต (mV)", "ความคลาดเคลื่อน (%)"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        set_cell_shading(cell, "F1F5F9")
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_font_sarabun(r, size_pt=14, bold=True)

    data = [
        ["1", "200.0", "12.450", "+0.12%"],
        ["2", "600.0", "37.380", "+0.08%"],
        ["3", "1,000.0 (1 Sun)", "62.310", "-0.05%"]
    ]
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            cell = table.cell(i+1, j)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(val)
            set_font_sarabun(r, size_pt=14)

    # -------------------------------------------------------------------------
    # บทที่ 5: สรุปผลการดำเนินงานและข้อเสนอแนะ
    # -------------------------------------------------------------------------
    doc.add_page_break()
    add_thesis_chapter_heading(doc, "5", "สรุปผลการดำเนินงานและข้อเสนอแนะ")

    add_section_heading(doc, "5.1", "สรุปผลการดำเนินงาน")
    add_thai_body_p(
        doc,
        "ระบบวัดการแผ่รังสีแสงอาทิตย์ความแม่นยำสูงที่พัฒนาขึ้น สามารถตรวจวัดสัญญาณรังสีแสงอาทิตย์ได้อย่างละเอียด "
        "โดยการนำ ADS1220 24-bit ADC และ REF5045 มาประยุกต์ใช้งานช่วยให้สัญญาณมีความนิ่งสูง ค่าสัมประสิทธิ์ความสัมพันธ์ "
        "R² สูงกว่า 0.999 และค่าความไม่แน่นอนของการวัดอยู่ในเกณฑ์มาตรฐานที่ยอมรับได้สำหรับการวิจัยขั้นสูง"
    )

    # บันทึกไฟล์เอกสาร
    out_path = os.path.join(os.path.dirname(__file__), output_filename)
    doc.save(out_path)
    print(f"[OK] Report template successfully generated at: {out_path}")
    return out_path

if __name__ == "__main__":
    build_kmitl_standard_report()
