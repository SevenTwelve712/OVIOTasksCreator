from docx.oxml import OxmlElement, CT_Tc
from docx.shared import Pt, Emu
from docx.table import Table, _Cell
from enum import Enum
from docx.oxml.ns import qn
from dataclasses import dataclass

from model.extended_docx_classes.ExtendedCell import ExtendedCell
from model.extended_docx_classes.ExtendedParagraph import JcTypes, ExtendedParagraph
from model.extended_docx_classes.TableHelpClasses import WidthTypes, TblBorder, LayoutTypes


class ExtendedTable:
    def __init__(self, table: Table):
        self._table = table
        self.pr = table._tbl.tblPr
        self.grid = table._tbl.tblGrid

    def _get_or_add_tblPr_node(self, node_name: str):
        node = self.pr.find(qn(node_name))
        if node is None:
            node = OxmlElement(node_name)
            self.pr.append(node)
        return node

    def set_width(self, wtype: WidthTypes, width: int=0):
        tblW = self._get_or_add_tblPr_node("w:tblW")
        tblW.set(qn("w:type"), wtype.value)
        if wtype is WidthTypes.DXA:
            tblW.set(qn("w:w"), str(width))
        print(self.pr.xml)

    def set_borders(self, borders: list[TblBorder] | TblBorder):
        """Задает границы для таблицы.\n
        Если передан список границ, то границы задаются в порядке: верхняя, левая, нижняя, правая, внутренняя горизонтальная,
        внутренняя вертикальная. Если в списке меньше 6 значений, то зададутся первые n границ.\n
        Если передана только одна граница, то все границы будут заданы по этому шаблону."""

        tags = ("top", "start", "bottom", "end", "insideH", "insideV")
        if isinstance(borders, TblBorder):
            borders = [borders] * len(tags)

        tblBorders = self._get_or_add_tblPr_node("w:tblBorders")
        for tag, border in zip(tags, borders):

            # Ищем узел границы таблицы
            border_elem = tblBorders.find(qn(f"w:{tag}"))
            if border_elem is None:
                border_elem = OxmlElement(f"w:{tag}")
                tblBorders.append(border_elem)

            border_elem.set(qn("w:val"), border.val)
            border_elem.set(qn("w:space"), str(border.space))
            border_elem.set(qn("w:sz"), str(border.sz))
            border_elem.set(qn("w:color"), border.color)

    def set_layout(self, wtype: LayoutTypes):
        tblL = self.pr.get_or_add_tblLayout()
        tblL.set(qn("w:type"), wtype.value)

    def set_jc(self, jc_type: JcTypes):
        if jc_type == JcTypes.BOTH:
            raise ValueError("Невозможно задать выравнивание таблицы по обоим краям")

        jc = self._get_or_add_tblPr_node("w:jc")
        jc.set(qn("w:type"), jc_type.value)

    def set_indent(self, indent: int):
        tblInd = self._get_or_add_tblPr_node("w:tblInd")
        tblInd.set(qn("w:w"), str(indent))
        tblInd.set(qn("w:type"), "dxa")

    def set_all_cells_borders(self, borders: list[TblBorder] | TblBorder):
        """Задает границы для всех ячеек"""
        for cell in self._table._cells:
            ExtendedCell(cell).set_borders(borders)

    def set_grids(self, grids: list[float] | None):
        if grids is None:
            self.grid.clear()
            return

        for i, grid in enumerate(self.grid.gridCol_lst):
            if i == len(grids):
                return
            grid.set(qn("w:w"), str(grids[i]))

    def set_cell_spacing(self, spacing: int):
        """Задает отступ между ячейками в dxa"""
        tblCellSpacing = self._get_or_add_tblPr_node("w:tblCellSpacing")
        tblCellSpacing.set(qn("w:w"), str(spacing))
        tblCellSpacing.set(qn("w:type"), "dxa")

    def set_cell_margins(self, margins: list[int] | int):
        """Задает отступы текста внутри ячеек в порядке верхний, левый, нижний, правый"""
        tags = ("top", "start", "bottom", "right")
        tblCellMar = self._get_or_add_tblPr_node("w:tblCellMar")

        if isinstance(margins, int):
            margins = [margins] * len(tags)

        for tag, margin in zip(tags, margins):
            # Ищем xml отступа
            margin_elem = tblCellMar.find(qn(f"w:{tag}"))
            if margin_elem is None:
                margin_elem = OxmlElement(f"w:{tag}")
                tblCellMar.append(margin_elem)

            margin_elem.set(qn("w:w"), str(margin))
            margin_elem.set(qn("w:type"), "dxa")

    def set_font_size_for_all_table(self, font_size: Pt):
        for row in self._table.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    for run in par:
                        run.font.size = font_size

    def rm_spacings_in_cells(self):
        for row in self._table.rows:
            for cell in row.cells:
                for par in cell.paragraphs:
                    ExtendedParagraph(par).rm_spacings()