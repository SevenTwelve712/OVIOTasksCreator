from docx.oxml import OxmlElement, CT_Tc
from docx.shared import Pt, Emu
from docx.table import Table, _Cell
from enum import Enum
from docx.oxml.ns import qn

class WidthTypes(Enum):
    DXA = "dxa"
    AUTO = "auto"

class LayoutTypes(Enum):
    FIXED = "fixed"
    AUTOFIT = "autofit"


class ExtendedTable:
    def __init__(self, table: Table):
        self._table = table
        self.pr = table._tbl.tblPr
        self.grid = table._tbl.tblGrid

    def _get_or_add_tblW(self):
        tblW = self.pr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement(qn("tblW"))
            self.pr.append(tblW)
        return tblW

    def set_width(self, wtype: WidthTypes, width: int=0):
        tblW = self._get_or_add_tblW()
        tblW.set(qn("w:type"), wtype.value)
        if wtype is WidthTypes.DXA:
            tblW.set(qn("w:w"), str(width))

    def set_layout(self, wtype: LayoutTypes):
        tblL = self.pr.get_or_add_tblLayout()
        tblL.set(qn("w:type"), wtype.value)

    def set_grids(self, grids: list[float] | None):
        if grids is None:
            self.grid.clear()
            return

        for i, grid in enumerate(self.grid.gridCol_lst):
            if i == len(grids):
                return
            grid.set(qn("w:w"), str(grids[i]))

    def set_font_size_for_all_table(self, font_size: Pt):
        for row in self._table.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    for run in par:
                        run.font.size = font_size

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