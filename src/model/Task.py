from pathlib import Path

from docx import Document

from src.model.TourTemplate import TourTemplate


class Task:
    name = "Неопределенное задание"
    cond = ""

    def __init__(self, tour_template: TourTemplate):
        self.tour_templ = tour_template


    def save_docx(self, path: Path):
        pass

    def save_xml(self, path_to_save: Path):
        pass

    def make_docx(self, doc: Document):
        self.tour_templ.make_docx(doc, self.name, self.cond)
        return self.tour_templ.doc

