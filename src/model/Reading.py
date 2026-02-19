from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Pt, Emu

from src.model.TourTemplate import TourTemplate
from src.model.docx_extended_utils.ExtendedTable import ExtendedTable, LayoutTypes, WidthTypes, ExtendedCell
from src.utils.docx_documents_utils import set_cols

from random import shuffle


class ReadingTable1Task:
    def __init__(self, doc: Document, words: list[str], preferred_font_size: int):
        """
        Реализует класс для удобной работы с объектом таблицы из первого задания для чтения. Она состоит из 1 строки и
        нескольких столбцов (по количеству слов), в каждой ячейке находится одно слово
        :param doc: Документ, в который должна быть добавлена таблица.
        :param words: Слова, которые должны находиться в таблице
        :param preferred_font_size: Предпочитаемый размер шрифта (может быть уменьшен)
        """
        if preferred_font_size not in [8, 10, 12]:
            raise ValueError("Размер шрифта должен быть 8, 10 или 12")

        self.ext_table = None
        self._table = None
        self.doc = doc
        self.words = words
        self.font_size = preferred_font_size
        self.default_table_width = self._calc_document_text_area_width()

    def create_table(self):
        """Добавляет саму таблицу в конец документа"""
        table = self.doc.add_table(1, len(self.words))
        self.ext_table = ExtendedTable(table)
        self._table = table

    def _calc_document_text_area_width(self):
        """Считает ширину текстового поля документа (для того, чтобы заполнить таблицу на всю ширину), возвращает ответ в dxa"""
        section = self.doc.sections[-1]
        return Emu(section.page_width - section.left_margin - section.right_margin).pt * 20

    def fill_words(self, bold: bool=False):
        """Заполняет таблицу заданными словами"""
        if self._table is None:
            raise RuntimeError("Таблица еще не создана")
        for i, cell in enumerate(self._table._cells):
            run = cell.paragraphs[0].add_run()
            run.text = f"{i}. {self.words[i]}"
            run.bold = bold
            run.font.size = Pt(self.font_size)
            run.font.name = "Times New Roman"

    def set_preferred_grid_cols_widths(self):
        """Задает предпочитаемую ширину для колонок"""
        grid_widths = [Pt(self.calc_min_cell_width(word)).twips for word in self.words]
        self.ext_table.set_grids(grid_widths)

    def calc_min_cell_width(self, text: str) -> float:
        """Считает примерную максимальную достижимую ширину ячейки в dxa (twips), учитывая шрифт (Times New Roman) и его размер, цифра получается очень приблизительная,
        она получена эмпирическим опытом"""
        if self.font_size == 8:
            return (len(text) * 59 + 75) / 7 * 20
        elif self.font_size == 10:
            return (75 + 73 * len(text)) / 7 * 20
        else:
            return (75 + 87 * len(text)) / 7 * 20

    def normalize_widths(self) -> bool:
        """Пытается нормализовать ширину таблицы: выставляет фиксированную ширину таблицы,
        задает предпочитаемые ширину колонок, устанавливает автоматическое распределение ширины"""
        self.ext_table.set_layout(LayoutTypes.AUTOFIT)
        self.ext_table.set_width(WidthTypes.DXA, width=self.default_table_width)
        self.set_preferred_grid_cols_widths()
        # for i, cell in enumerate(self._table._cells):
        #     min_cell_width_twips = int(self.calc_min_cell_width(self.words[i]) * 20)
        #     ExtendedCell(cell).set_width(WidthTypes.DXA, min_cell_width_twips)
        return True


class Reading:
    name = "Чтение"

    def __init__(self, tour_templ: TourTemplate, text: str, matches: dict[str, str], questions: list[str], word: tuple[str, str], mistake_words: tuple[str, str]):
        """
        Класс задания чтение
        :param tour_templ: Шаблон заголовка задания.
        :param text: Текст, по которому выполняется задание.
        :param matches: Задание соответствий (задание 1), в формате {слово: ассоциация}.
        :param questions: Вопросы (задание 2).
        :param word: Слово по определению (задание 3), в формате (слово, его определение).
        :param mistake_words: Ошибочное и верные слова в формате (слово, слово)
        """

        self.tour_templ = tour_templ
        self.text = text
        self.matches = matches
        self.questions = questions
        self.word = word
        self.mistake_words = mistake_words
        self.doc = None

    def make_xml(self):
        pass

    def make_docx(self, doc: Document):
        self.tour_templ.make_docx(doc, self.name)
        doc = self.tour_templ.doc

        # Добавляем секцию текста
        text_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        set_cols(text_sec, 2)
        doc.add_paragraph(self.text, style="ReadingTask")

        tasks_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        set_cols(tasks_sec, 1)

        # 1 задание
        f_task_par = doc.add_paragraph(style="ReadingTask")
        f_task_par.add_run("1. Заполните таблицу. Под каждым словом запишите НОМЕР соответствующего ему слова из списка (по 1 баллу за соответствие):").bold = True

        f_task_cond = doc.add_table(rows=1, cols=len(self.matches))
        for i, key in enumerate(self.matches.keys()):
            f_task_cond.cell(0, i).TEXT = f"{i + 1}. {key.capitalize()}"

        f_task_solution = doc.add_table(rows=2, cols = len(self.matches))
        match_items = list(self.matches.values())
        shuffle(match_items)

        for i, item in enumerate(match_items):
            f_task_solution.cell(0, i).TEXT = item.upper()

        # 2 задание
        s_task_par = doc.add_paragraph(style="ReadingTask")
        s_task_par.add_run("2. Заполните таблицу (по 2 балла за правильное заполнение. Слова должны быть написаны без ошибок):").bold = True
        s_task = doc.add_table(rows=len(self.questions), cols=2)

        for i, question in enumerate(self.questions):
            s_task.cell(0, i).TEXT = f"2.{i}. {question}"

        # 3 задание
        t_task_par = doc.add_paragraph(style="ReadingTask")
        t_task_par.add_run("3. Определите слово по описанию (2 балла). Это слово обязательно должно быть в тексте.").bold = True
        t_task_par.add_run(f"{'_' * int(len(self.word[0]) / 0.7)} — {self.word[1]} ({len(self.word[0])} букв)")

        # 4 задание
        fo_task_par = doc.add_paragraph(style="ReadingTask")
        fo_task_par.add_run("4. Найдите в тексте ошибочное слово и замените его на верное (найденное – 1 балл, правильная замена – 1 балл):").bold = True
        fo_task = doc.add_table(rows=2, cols=2)
        fo_task.cell(0, 0).paragraphs[0].add_run("Ошибочное").bold = True
        fo_task.cell(0, 1).paragraphs[0].add_run(f"Правильное ({len(self.mistake_words[1])} букв)").bold = True
        self.doc = doc