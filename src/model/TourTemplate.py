from docx.document import Document
from datetime import date
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.model.TourData import Forms, Tours


class TourTemplate:
    months = {
        1: "января",
        2: "февраля",
        3: "марта",
        4: "апреля",
        5: "мая",
        6: "июня",
        7: "июля",
        8: "августа",
        9: "сентября",
        10: "октября",
        11: "ноября",
        12: "декабря"
    }

    def __init__(self, form: Forms, tour: Tours, tour_date: date, place: str):
        self.tour = tour
        self.form = form
        self.tour_date = tour_date
        self.place = place
        self.doc = None

    def make_xml(self):
        pass

    def make_docx(self, doc: Document, task_name: str):
        par = doc.paragraphs[0]
        par.style = "OVIOHeader"
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER

        par.add_run(f'{self.tour.value} ОВИО "Наше наследие" среди {self.form.value} классов, ').bold = True

        datetime_data = par.add_run(f"{self.tour_date.day} {self.months[self.tour_date.month]} {self.tour_date.year}, {self.place}")
        datetime_data.italic = True
        datetime_data.add_break()

        task_data = par.add_run(f'ЗАДАНИЕ "{task_name}"')
        task_data.italic = True
        task_data.add_break()
        self.doc = doc
