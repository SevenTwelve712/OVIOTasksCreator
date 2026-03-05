from enum import Enum

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


class JcTypes(Enum):
    BOTH ="both"
    CENTER = "center"
    END = "end"


class ExtendedParagraph:
    def __init__(self, paragraph: Paragraph):
        self.par = paragraph
        self._parPr = paragraph._p.get_or_add_pPr()

    def set_jc(self, jc_type: JcTypes):
        jc = self._parPr.jc
        if jc is None:
            jc = OxmlElement("w:jc")
            self._parPr.append(jc)
        jc.set(qn("w:val"), jc_type.value)