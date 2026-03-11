from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import _Cell

from model.extended_docx_classes.TableHelpClasses import WidthTypes, TblBorder


class ExtendedCell:
    def __init__(self, cell: _Cell):
        self._cell = cell
        self.tcPr = cell._tc.get_or_add_tcPr()

    def _get_or_add_tcPr_node(self, node_name: str):
        node = self.tcPr.find(qn(node_name))
        if node is None:
            node = OxmlElement(node_name)
            self.tcPr.append(node)
        return node

    def set_width(self, wtype: WidthTypes, width: int=0):
        tcW = self.tcPr.get_or_add_tcW()
        tcW.set(qn("w:type"), wtype.value)
        tcW.set(qn("w:w"), str(width))

    def set_borders(self, borders: list[TblBorder] | TblBorder):
        """Задает границы для таблицы.\n
        Если передан список границ, то границы задаются в порядке: верхняя, левая, нижняя, правая, внутренняя горизонтальная,
        внутренняя вертикальная. Если в списке меньше 6 значений, то зададутся первые n границ.\n
        Если передана только одна граница, то все границы будут заданы по этому шаблону."""

        tags = ("top", "start", "bottom", "end", "insideH", "insideV")
        if isinstance(borders, TblBorder):
            borders = [borders] * len(tags)

        tcBorders = self._get_or_add_tcPr_node("w:tcBorders")
        for tag, border in zip(tags, borders):

            # Ищем узел границы таблицы
            border_elem = tcBorders.find(qn(f"w:{tag}"))
            if border_elem is None:
                border_elem = OxmlElement(f"w:{tag}")
                tcBorders.append(border_elem)

            border_elem.set(qn("w:val"), border.val)
            border_elem.set(qn("w:space"), str(border.space))
            border_elem.set(qn("w:sz"), str(border.sz))
            border_elem.set(qn("w:color"), border.color)