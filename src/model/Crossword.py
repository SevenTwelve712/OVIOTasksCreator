from enum import EnumType
from xml.dom.minidom import Document

from src.model.Task import Task

class HeightTypes(EnumType):
    PX = "pixels"
    CELLS = "cells"


class Crossword(Task):
    name = "Кроссворд"
    cond = ""

    CELL_SIZE = 20 # размер одной ячейки кроссворда в px для генерации jpg
    BORDER_SIZE = 5 # толщина границы вокруг каждой ячейки в px

    def __init__(self, words: list[tuple[str, str]], max_height: int, height_type: HeightTypes):
        """
        Класс задания кроссворд
        :param words: слова в формате списка из множеств (слово, описание)
        :param max_height: максимальная высота (в пикселях или в ячейках)
        :param height_type: тип максимальной высоты
        """

        self.words = words

        if height_type is HeightTypes.PX:
            self.max_height = self._count_height_cells(max_height)
        else:
            self.max_height = max_height

    def _count_height_cells(self, px_height: int):
        return int(px_height / self.CELL_SIZE - self.BORDER_SIZE * 2)

    def make_docx(self, doc: Document):
        doc = super().make_docx(doc)

