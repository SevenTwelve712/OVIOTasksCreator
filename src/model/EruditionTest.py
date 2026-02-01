from docx.enum.section import WD_SECTION

from docx import Document
from pathlib import Path

from docx.oxml import OxmlElement, ns

from src.model.TourTemplate import TourTemplate



class EruditionTest:
    name = "Тест"
    def __init__(self, questions: dict[str: tuple[list[str], str]], template: TourTemplate):
        """
        Класс теста на эрудицию.
        :param questions: Список вопросов в виде {"вопрос": (["4 варианта ответа"], 'буква правильного ответа')
        :param template: Шаблон заглавной плашки данного тура олимпиады
        """
        self.questions = questions
        self.template = template

    def make_xml(self):
        pass

    def save_xml(self, path_to_save: Path):
        pass

    def make_docx(self, doc: Document):
        letters = ['А', 'Б', 'В', 'Г']
        doc = self.template.make_docx(doc, self.name)

        # Добавляем главную секцию с колонками
        task_sect = doc.add_section(WD_SECTION.CONTINUOUS)
        cols = OxmlElement('w:cols')
        cols.set(ns.qn('w:num'), '3')
        task_sect._sectPr.append(cols)

        num = 0
        for question, answers_ in self.questions.items():
            num += 1
            par = doc.add_paragraph(style='ETestTask')
            answers, letter = answers_

            # Прогон вопроса
            question_run = par.add_run(f"{str(num)}. {question}")
            question_run.bold = True

            # Прогон вариантов ответов
            answer_run = par.add_run()

            for a_num, answer in enumerate(answers):
                answer_run.add_break()
                answer_run.add_text(f"{letters[a_num]}. {answer}")
        return doc


    def save_docx(self, path_to_save: Path):
        pass