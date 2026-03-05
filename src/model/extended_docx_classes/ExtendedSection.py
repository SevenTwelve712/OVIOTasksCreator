from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.section import Section


class ExtendedSection:
    def __init__(self, section: Section):
        self.section = section
        self.sectPr = section._sectPr

    def get_text_area_width(self):
        """Возвращает ширину поля, где пишется текст в EMU"""
        return self.section.page_width - self.section.left_margin - self.section.right_margin

    def set_cols(self, num: int):
        cols = self.sectPr.find(qn('w:cols'))
        if cols is None:
            cols = OxmlElement('w:cols')
            self.sectPr.append(cols)
        cols.set(qn('w:num'), str(num))

    def set_size(self, width: int, height: int):
        pgSz = self.sectPr.find(qn('w:pgSz'))
        if pgSz is None:
            pgSz = OxmlElement('w:pgSz')
            self.sectPr.append(pgSz)
        pgSz.set(qn('w:w'), str(width))
        pgSz.set(qn('w:h'), str(height))

    def set_size_a4(self):
        self.set_size(11906, 16838)

    def set_margins(self, top: int, bottom: int, left: int, right: int, header: int, footer: int, gutter: int):
        pgMar = self.sectPr.find(qn('w:pgMar'))
        if pgMar is None:
            pgMar = OxmlElement('w:pgMar')
            self.sectPr.append(pgMar)
        pgMar.set(qn('w:top'), str(top))
        pgMar.set(qn('w:bottom'), str(bottom))
        pgMar.set(qn('w:left'), str(left))
        pgMar.set(qn('w:right'), str(right))
        pgMar.set(qn('w:header'), str(header))
        pgMar.set(qn('w:footer'), str(footer))
        pgMar.set(qn('w:gutter'), str(gutter))