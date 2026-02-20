from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run


class ExtendedRun:
    def __init__(self, run: Run):
        self.run = run
        self._rPr = run._r.get_or_add_rPr()

    def set_spacing(self, val: int):
        spacing = self._rPr.find(qn("w:spacing"))
        if spacing is None:
            spacing = OxmlElement("w:spacing")
            self._rPr.append(spacing)

        spacing.set(qn("w:val"), str(val))