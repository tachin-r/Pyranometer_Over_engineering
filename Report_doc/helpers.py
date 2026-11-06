import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==============================================================================
# 1. ฟังก์ชันตั้งค่าฟอนต์ TH SarabunPSK ระดับ OXML (ป้องกันฟอนต์เพี้ยนใน Word)
# ==============================================================================

def set_font_sarabun(run, size_pt=16, bold=False, italic=False, color_rgb=None):
    """
    กำหนดฟอนต์ TH SarabunPSK โดยฉีดแท็ก XML w:rFonts ให้ครบทุกระบบ
    (ascii, hAnsi, และ cs สำหรับ Complex Script ภาษาไทย)
    รับประกันฟอนต์ไม่สลับเป็น Angsana หรือ Cordia เมื่อเปิดใน MS Word
    """
    run.font.name = 'TH SarabunPSK'
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'TH SarabunPSK')
    rFonts.set(qn('w:hAnsi'), 'TH SarabunPSK')
    rFonts.set(qn('w:cs'), 'TH SarabunPSK')
    rPr.append(rFonts)

# ==============================================================================
# 2. ฟังก์ชันตั้งค่าหน้ากระดาษตามมาตรฐานวิทยานิพนธ์/รายงานทางการ สจล. (KMITL)
# ==============================================================================

def setup_kmitl_page(section):
    """
    ตั้งค่าขนาดหน้ากระดาษ A4 และระยะขอบมาตรฐานรายงาน สจล.:
    - ขนาดกระดาษ: A4 (กว้าง 8.27 นิ้ว x ยาว 11.69 นิ้ว)
    - ขอบบน (Top): 1.5 นิ้ว (3.81 ซม.)
    - ขอบซ้าย (Left): 1.5 นิ้ว (3.81 ซม.) [เผื่อเย็บเล่มปกแข็ง/กระดูกงู]
    - ขอบล่าง (Bottom): 1.0 นิ้ว (2.54 ซม.)
    - ขอบขวา (Right): 1.0 นิ้ว (2.54 ซม.)
    """
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.5)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.5)
    section.right_margin = Inches(1.0)
    return section

def add_page_number_to_footer(footer, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """
    ใส่เลขหน้าแบบ Dynamic Field (w:fldSimple w:instr='PAGE') ใน Footer
    """
    p = footer.paragraphs[0]
    p.alignment = alignment
    r = p.add_run()
    set_font_sarabun(r, size_pt=14)
    fldSimple = OxmlElement('w:fldSimple')
    fldSimple.set(qn('w:instr'), 'PAGE')
    p._p.append(fldSimple)
    return p

def set_section_numbering(section, fmt="decimal", start=1):
    """
    กำหนดรูปแบบเลขหน้าของ Section:
    - fmt='decimal' (1, 2, 3...)
    - fmt='lowerRoman' (i, ii, iii...)
    - fmt='thaiCounting' หรือ format อื่น ๆ
    """
    sectPr = section._sectPr
    for child in list(sectPr):
        if child.tag.endswith('pgNumType'):
            sectPr.remove(child)
    pgNumType = OxmlElement('w:pgNumType')
    pgNumType.set(qn('w:fmt'), fmt)
    if start is not None:
        pgNumType.set(qn('w:start'), str(start))
    sectPr.append(pgNumType)

# ==============================================================================
# 3. ฟังก์ชันจัดหัวข้อและข้อความ (Headings & Typography Hierarchy)
# ==============================================================================

def add_thesis_chapter_heading(doc, chap_num, chap_title):
    """
    หัวข้อบทตามมาตรฐานวิทยานิพนธ์:
    บรรทัด 1: 'บทที่ X' (18 pt ตัวหนา กึ่งกลาง)
    บรรทัด 2: 'ชื่อบท' (18 pt ตัวหนา กึ่งกลาง)
    """
    p_num = doc.add_paragraph()
    p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_num.paragraph_format.space_before = Pt(20)
    p_num.paragraph_format.space_after = Pt(4)
    p_num.paragraph_format.keep_with_next = True
    r_num = p_num.add_run(f"บทที่ {chap_num}")
    set_font_sarabun(r_num, size_pt=18, bold=True)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(14)
    p_title.paragraph_format.keep_with_next = True
    r_title = p_title.add_run(chap_title)
    set_font_sarabun(r_title, size_pt=18, bold=True)
    return p_title

def add_section_heading(doc, sec_num, sec_title):
    """
    หัวข้อหลักระดับที่ 1 (เช่น 1.1 ความเป็นมาและขอบเขต):
    16 pt ตัวหนา ชิดซ้าย
    """
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    prefix = f"{sec_num} " if sec_num else ""
    r = p.add_run(f"{prefix}{sec_title}")
    set_font_sarabun(r, size_pt=16, bold=True)
    return p

def add_subsection_heading(doc, subsec_num, subsec_title):
    """
    หัวข้อย่อยระดับที่ 2 (เช่น 1.1.1 ส่วนตรวจวัดสัญญาณ):
    16 pt ตัวหนา ย่อหน้าเข้า 0.25 นิ้ว
    """
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    prefix = f"{subsec_num} " if subsec_num else ""
    r = p.add_run(f"{prefix}{subsec_title}")
    set_font_sarabun(r, size_pt=16, bold=True)
    return p

def add_thai_body_p(doc, text, space_after=6, indent=0.5, line_spacing=1.15):
    """
    ย่อหน้าเนื้อหาทั่วไปภาษาไทย:
    - ฟอนต์ TH SarabunPSK ขนาด 16 pt ปกติ
    - ย่อหน้าบรรทัดแรก 0.5 นิ้ว (1.27 ซม.)
    - ระยะห่างบรรทัด 1.15
    """
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Inches(indent)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = line_spacing
    r = p.add_run(text)
    set_font_sarabun(r, size_pt=16)
    return p

def add_numbered_item(doc, num_str, text, bold_prefix=None, space_after=3, left_indent=0.55, hanging_indent=0.35):
    """
    รายการแบบมีลำดับเลข โดยใช้ Hanging Indent + Tab Stop ทำให้ตัวหนังสือตรงแนวกันเป๊ะ
    """
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(left_indent)
    p.paragraph_format.first_line_indent = Inches(-hanging_indent)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.tab_stops.add_tab_stop(Inches(left_indent), WD_TAB_ALIGNMENT.LEFT)

    r_num = p.add_run(f"{num_str}\t")
    set_font_sarabun(r_num, size_pt=16, bold=True)

    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        set_font_sarabun(r_pre, size_pt=16, bold=True)

    r_text = p.add_run(text)
    set_font_sarabun(r_text, size_pt=16)
    return p

def add_bullet_item(doc, bullet_char="•", text="", bold_prefix=None, space_after=3, left_indent=0.55, hanging_indent=0.25):
    """
    รายการแบบ Bullet โดยใช้ Hanging Indent + Tab Stop
    """
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(left_indent)
    p.paragraph_format.first_line_indent = Inches(-hanging_indent)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.tab_stops.add_tab_stop(Inches(left_indent), WD_TAB_ALIGNMENT.LEFT)

    r_bullet = p.add_run(f"{bullet_char}\t")
    set_font_sarabun(r_bullet, size_pt=16, bold=True)

    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        set_font_sarabun(r_pre, size_pt=16, bold=True)

    r_text = p.add_run(text)
    set_font_sarabun(r_text, size_pt=16)
    return p

def add_toc_line(doc, title, page_str, level=0, is_bold=False, tab_stop_inch=5.75):
    """
    สร้างบรรทัดสารบัญพร้อมเส้นจุดไข่ปลา (Tab Leader Dots) วิ่งไปชิดขวาที่ระยะ tab_stop_inch
    """
    p = doc.add_paragraph()
    p.paragraph_format.tab_stops.add_tab_stop(Inches(tab_stop_inch), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.15
    if level == 1:
        p.paragraph_format.left_indent = Inches(0.35)
    elif level == 2:
        p.paragraph_format.left_indent = Inches(0.7)

    r_title = p.add_run(title)
    set_font_sarabun(r_title, size_pt=16, bold=is_bold)

    r_tab = p.add_run(f"\t{page_str}")
    set_font_sarabun(r_tab, size_pt=16, bold=is_bold)
    return p

# ==============================================================================
# 4. ฟังก์ชันจัดการตาราง (Table Formatting)
# ==============================================================================

def set_cell_shading(cell, color_hex):
    """กำหนดสีพื้นหลังของเซลล์ตารางด้วย HEX (เช่น 'F8FAFC', 'E2E8F0')"""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    """กำหนด Padding ด้านในของเซลล์ตาราง (หน่วยเป็น dxa: 20 dxa = 1 pt)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_clean_borders(table, border_color="000000", border_sz="4"):
    """
    กำหนดเส้นตารางแบบเส้นเดี่ยว สะอาด คมชัด สไตล์วิทยานิพนธ์ (ไม่หนาเทอะทะ)
    """
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:left w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:right w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:insideH w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'  <w:insideV w:val="single" w:sz="{border_sz}" w:space="0" w:color="{border_color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_table_caption(doc, table_no, title):
    """
    หัวตารางมาตรฐานวิทยานิพนธ์ (ตารางที่ X.X ชื่อตาราง) อยู่ด้านบนตาราง
    """
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r1 = p.add_run(f"ตารางที่ {table_no} ")
    set_font_sarabun(r1, size_pt=14, bold=True)
    r2 = p.add_run(title)
    set_font_sarabun(r2, size_pt=14)
    return p

# ==============================================================================
# 5. ฟังก์ชันใส่รูปภาพและ Photo Placeholder
# ==============================================================================

def add_figure_image(doc, img_path, fig_no, title, width_inches=5.2):
    """
    แทรกรูปภาพจริง จัดกึ่งกลาง พร้อมคำบรรยายใต้ภาพตามมาตรฐาน (รูปที่ X.X ชื่อรูป)
    """
    if not os.path.exists(img_path):
        return None
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(8)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_inches))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(8)
    r_cap = p_cap.add_run(f"รูปที่ {fig_no} {title}")
    set_font_sarabun(r_cap, size_pt=14, bold=True)
    return p_img

def add_two_figure_images_side_by_side(doc, img1_path, cap1, img2_path, cap2, fig_no, main_title, col_width_inch=2.7):
    """
    แทรกรูปภาพ 2 รูปเคียงข้างกันใน 1 แถว พร้อมคำบรรยายย่อย (ก) และ (ข) และคำบรรยายภาพหลัก
    """
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:bottom w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="none"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

    for col_idx, (ipath, cap) in enumerate([(img1_path, cap1), (img2_path, cap2)]):
        cell = table.cell(0, col_idx)
        set_cell_margins(cell, top=20, bottom=20, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        if os.path.exists(ipath):
            r_img = p.add_run()
            r_img.add_picture(ipath, width=Inches(col_width_inch))
        if cap:
            p_c = cell.add_paragraph()
            p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_c.paragraph_format.space_before = Pt(2)
            p_c.paragraph_format.space_after = Pt(4)
            r_c = p_c.add_run(cap)
            set_font_sarabun(r_c, size_pt=13)

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after = Pt(8)
    r_cap = p_cap.add_run(f"รูปที่ {fig_no} {main_title}")
    set_font_sarabun(r_cap, size_pt=14, bold=True)
    return table

def add_two_figure_images_stacked(doc, img1_path, cap1, img2_path, cap2, fig_no, main_title, width_inches=5.2):
    """
    แทรกรูปภาพ 2 รูปซ้อนกันในแนวตั้ง พร้อมคำบรรยายย่อย (ก) และ (ข) และคำบรรยายภาพหลัก
    """
    if os.path.exists(img1_path):
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(6)
        p1.paragraph_format.space_after = Pt(2)
        p1.paragraph_format.keep_with_next = True
        p1.add_run().add_picture(img1_path, width=Inches(width_inches))
        if cap1:
            pc1 = doc.add_paragraph()
            pc1.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pc1.paragraph_format.space_before = Pt(2)
            pc1.paragraph_format.space_after = Pt(6)
            pc1.paragraph_format.keep_with_next = True
            r1 = pc1.add_run(cap1)
            set_font_sarabun(r1, size_pt=13)

    if os.path.exists(img2_path):
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(6)
        p2.paragraph_format.space_after = Pt(2)
        p2.paragraph_format.keep_with_next = True
        p2.add_run().add_picture(img2_path, width=Inches(width_inches))
        if cap2:
            pc2 = doc.add_paragraph()
            pc2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pc2.paragraph_format.space_before = Pt(2)
            pc2.paragraph_format.space_after = Pt(6)
            pc2.paragraph_format.keep_with_next = True
            r2 = pc2.add_run(cap2)
            set_font_sarabun(r2, size_pt=13)

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(4)
    p_cap.paragraph_format.space_after = Pt(8)
    r_cap = p_cap.add_run(f"รูปที่ {fig_no} {main_title}")
    set_font_sarabun(r_cap, size_pt=14, bold=True)

def add_photo_placeholder(doc, fig_no, title, description="", instructions=""):
    """
    สร้างกรอบ Placeholder สวยงามสำหรับเตรียมวางรูปถ่าย/ภาพแคปหน้าจอ พร้อมคำอธิบาย
    """
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(8)
    p_cap.paragraph_format.space_after = Pt(4)
    p_cap.paragraph_format.keep_with_next = True
    r_cap = p_cap.add_run(f"รูปที่ {fig_no} {title}")
    set_font_sarabun(r_cap, size_pt=14, bold=True, color_rgb=RGBColor(0, 0, 0))

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_shading(cell, "F8FAFC")
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)

    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="0284C7"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run(f"📸 [ กรุณาแทรกรูปภาพที่ {fig_no}: {title} ]\n")
    set_font_sarabun(r1, size_pt=14, bold=True, color_rgb=RGBColor(2, 132, 199))

    if description:
        r2 = p.add_run(f"คำอธิบายภาพ: {description}\n")
        set_font_sarabun(r2, size_pt=13, color_rgb=RGBColor(51, 65, 85))

    if instructions:
        r3 = p.add_run(f"คำแนะนำ: {instructions}\n")
        set_font_sarabun(r3, size_pt=12, italic=True, color_rgb=RGBColor(100, 116, 139))

    r4 = p.add_run("*(สามารถลบตารางนี้แล้ววางรูปภาพจริงแทนที่ได้ทันที)*")
    set_font_sarabun(r4, size_pt=11, italic=True, color_rgb=RGBColor(148, 163, 184))

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_after = Pt(6)
    return table

# ==============================================================================
# 6. ฟังก์ชัน Backward Compatibility (สำหรับสคริปต์เดิมที่เคยเรียกใช้)
# ==============================================================================

def add_heading_1(doc, text):
    """เข้ากันได้กับโค้ดเดิม: หัวข้อหลักขนาด 18 pt ตัวหนา"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font_sarabun(r, size_pt=18, bold=True)
    return p

def add_heading_2(doc, text):
    """เข้ากันได้กับโค้ดเดิม: หัวข้อระดับสองขนาด 16 pt ตัวหนา"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font_sarabun(r, size_pt=16, bold=True)
    return p

def add_heading_3(doc, text):
    """เข้ากันได้กับโค้ดเดิม: หัวข้อระดับสามขนาด 15 pt ตัวหนา"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font_sarabun(r, size_pt=15, bold=True)
    return p

def add_body_p(doc, text):
    """เข้ากันได้กับโค้ดเดิม: เรียกใช้ add_thai_body_p"""
    return add_thai_body_p(doc, text, space_after=6, indent=0.5)

def add_bullet_p(doc, bold_prefix, text):
    """เข้ากันได้กับโค้ดเดิม: เรียกใช้ add_bullet_item"""
    return add_bullet_item(doc, bullet_char="•", text=text, bold_prefix=f"{bold_prefix} ")

print("[OK] Helpers loaded successfully with KMITL standard formatting & TH SarabunPSK XML font support.")
