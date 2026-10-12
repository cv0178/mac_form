from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUTPUT = Path(r"D:\MyPython\fill_inForm\test_templates\复杂表格个人登记.docx")


def set_cell_shading(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color)
    tc_pr.append(shading)


def set_cell_text(cell, text: str, *, bold: bool = False, color: str | None = None) -> None:
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(10.5)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def label(cell, text: str) -> None:
    set_cell_shading(cell, "D9EAF7")
    set_cell_text(cell, text, bold=True, color="163A5F")


def input_cell(cell, hint: str = "") -> None:
    set_cell_shading(cell, "FFF6D9")
    set_cell_text(cell, hint, color="8A6D3B")


def add_section_title(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(10)
    paragraph.paragraph_format.space_after = Pt(5)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(31, 78, 121)


def make_table(document: Document, rows: int, cols: int):
    table = document.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    table.autofit = False
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return table


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    normal = document.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10.5)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("复杂表格个人登记表")
    title_run.bold = True
    title_run.font.name = "Microsoft YaHei"
    title_run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    title_run.font.size = Pt(18)

    note = document.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    note.paragraph_format.space_after = Pt(6)
    note_run = note.add_run("测试模板：浅黄色单元格为待填内容；请保留蓝色标签和说明文字。")
    note_run.font.name = "Microsoft YaHei"
    note_run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    note_run.font.size = Pt(9)
    note_run.font.color.rgb = RGBColor(102, 102, 102)

    add_section_title(document, "一、基础信息")
    basic = make_table(document, 5, 4)
    basic.rows[0].cells[0].width = Cm(2.4)
    basic.rows[0].cells[1].width = Cm(5.3)
    basic.rows[0].cells[2].width = Cm(2.4)
    basic.rows[0].cells[3].width = Cm(5.3)
    basic_data = [
        ("姓名", "", "性别", ""),
        ("出生日期", "", "联系电话", ""),
        ("工作单位", "", "职位", ""),
        ("常住城市", "", "最高学历", ""),
        ("电子邮箱", "", "工作年限", ""),
    ]
    for row, values in zip(basic.rows, basic_data):
        label(row.cells[0], values[0])
        input_cell(row.cells[1], values[1])
        label(row.cells[2], values[2])
        input_cell(row.cells[3], values[3])

    add_section_title(document, "二、专业能力与项目经历")
    skills = make_table(document, 1, 2)
    label(skills.cell(0, 0), "专业技能（100 字以内）")
    input_cell(skills.cell(0, 1))
    skills.rows[0].height = Cm(2.2)

    projects = make_table(document, 3, 4)
    headers = ["项目名称", "项目时间", "担任角色", "核心贡献或成果"]
    for index, header in enumerate(headers):
        label(projects.cell(0, index), header)
        projects.cell(0, index).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row_index in range(1, 3):
        for col_index in range(4):
            hint = ""
            input_cell(projects.cell(row_index, col_index), hint)
        projects.rows[row_index].height = Cm(1.5)

    add_section_title(document, "三、个人陈述")
    profile = make_table(document, 1, 2)
    label(profile.cell(0, 0), "个人工作简介（300 字以内）")
    input_cell(profile.cell(0, 1))
    profile.rows[0].height = Cm(4.2)

    reflection = make_table(document, 1, 2)
    label(reflection.cell(0, 0), "工作心得（500 字以内）")
    input_cell(reflection.cell(0, 1))
    reflection.rows[0].height = Cm(5.5)

    add_section_title(document, "四、确认项")
    confirm = make_table(document, 2, 2)
    label(confirm.cell(0, 0), "是否接受出差")
    set_cell_text(confirm.cell(0, 1), "□ 是    □ 否    □ 视情况而定")
    label(confirm.cell(1, 0), "填表人签名")
    set_cell_text(confirm.cell(1, 1), "（此处保留人工填写）", color="8A6D3B")

    document.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    main()
