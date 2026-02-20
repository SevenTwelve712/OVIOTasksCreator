from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import _Cell

from src.model.docx_extended_utils.ExtendedTable import WidthTypes


class ExtendedCell:
    def __init__(self, cell: _Cell):
        self._cell = cell

    def set_width(self, wtype: WidthTypes, width: int=0):
        tcPr = self._cell._tc.get_or_add_tcPr()
        tcW = tcPr.get_or_add_tcW()
        tcW.set(qn("w:type"), wtype.value)
        tcW.set(qn("w:w"), str(width))

    def set_wrapping(self, can_wrap: bool=False):
        tcPr = self._cell._tc.get_or_add_tcPr()
        noWrap = tcPr.find(qn("w:noWrap"))
        if can_wrap and noWrap is not None:
            tcPr.remove(qn("w:noWrap"))
        elif not can_wrap and noWrap is None:
            noWrap = OxmlElement(qn("w:noWrap"))
            tcPr.append(noWrap)