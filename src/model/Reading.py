from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, ns

from src.model.TourTemplate import TourTemplate
from src.utils.docx_documents_utils import set_cols

from random import shuffle


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
            f_task_cond.cell(0, i).text = f"{i + 1}. {key.capitalize()}"

        f_task_solution = doc.add_table(rows=2, cols = len(self.matches))
        match_items = list(self.matches.values())
        shuffle(match_items)

        for i, item in enumerate(match_items):
            f_task_solution.cell(0, i).text = item.upper()

        # 2 задание
        s_task_par = doc.add_paragraph(style="ReadingTask")
        s_task_par.add_run("2. Заполните таблицу (по 2 балла за правильное заполнение. Слова должны быть написаны без ошибок):").bold = True
        s_task = doc.add_table(rows=len(self.questions), cols=2)

        for i, question in enumerate(self.questions):
            s_task.cell(0, i).text = f"2.{i}. {question}"

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