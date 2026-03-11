from enum import Enum

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


class JcTypes(Enum):
    """Класс типов выравниваний параграфа. Варианты:\n
    1) BOTH -> выравнивание относительно обоих краев
    2) CENTER -> выравнивание по центру
    3) END -> выравнивание по левому краю"""
    BOTH ="both"
    CENTER = "center"
    END = "end"

# TODO: проверить, всегда ли надо лезть в xml, нельзя ли где то обойтись вызовами python-docx api
class ExtendedParagraph:
    def __init__(self, paragraph: Paragraph):
        self.par = paragraph
        self._parPr = paragraph._p.get_or_add_pPr()
        self.fmt = paragraph.paragraph_format

    def _get_or_add_pPr_node(self, node_name: str):
        node = self._parPr.find(qn(node_name))
        if node is None:
            node = OxmlElement(node_name)
            self._parPr.append(node)
        return node

    def set_jc(self, jc_type: JcTypes):
        """Задает выравнивание в параграфе"""
        self._get_or_add_pPr_node("w:jc").set(qn("w:val"), jc_type.value)

    def set_indent(self, right: int=None, left: int=None, first_line: int=None):
        """Задает отступ строки от края родителя, значения принимаются в pt"""
        ind = self._get_or_add_pPr_node("w:ind")
        if right is not None:
            ind.set(qn("w:right"), str(right))
        if left is not None:
            ind.set(qn("w:left"), str(left))
        if first_line is not None:
            ind.set(qn("w:firstLine"), str(first_line))

    def rm_spacings(self):
        self.fmt.space_after = 0
        self.fmt.space_before = 0


    def set_spacing(self, before: int=None, after: int=None):
        if before is not None:
            self.fmt.space_before = before
        if after is not None:
            self.fmt.space_after = after