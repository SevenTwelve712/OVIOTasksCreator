from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.section import Section


class ExtendedSection:
    def __init__(self, section: Section):
        self.section = section

    def get_text_area_width(self):
        """Возвращает ширину поля, где пишется текст в EMU"""
        return self.section.page_width - self.section.left_margin - self.section.right_margin

    def set_cols(self, num: int):
        sectPr = self.section._sectPr
        cols = sectPr.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            sectPr.append(cols)
        cols.set(qn('w:num'), str(num))
