from docx import Document
from pathlib import Path

from model.TourTemplate import TourTemplate


class EruditionTest:
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

    def make_docx(self):
        pass

    def save_docx(self, path_to_save: Path):
        pass