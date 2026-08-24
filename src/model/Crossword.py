from enum import EnumType
from io import BytesIO
from docx.document import Document

from docx.enum.section import WD_SECTION
from docx.shared import Emu, Inches, Mm

from model.extended_docx_classes.data_and_enums import Direction, JcTypes
from model.extended_docx_classes.ExtendedParagraph import ExtendedParagraph
from model.extended_docx_classes.ExtendedSection import ExtendedSection
from model.Task import Task
from model.TourTemplate import TourTemplate
from model.vendor.complexstring import ComplexString
from model.vendor.genxword import Crossword


class HeightTypes(EnumType):
    PX = "pixels"
    CELLS = "cells"


# TODO: test
class OVIOCrossword(Task):
    name = "Кроссворд"
    cond = ""

    CELL_SIZE = Mm(8).emu  # размер одной ячейки кроссворда в emu для генерации jpg
    BORDER_SIZE = Mm(0.3).emu  # толщина границы вокруг каждой ячейки в emu
    DPI = 300

    def __init__(
        self,
        words: list[tuple[str, str]],
        max_height: int,
        height_type: HeightTypes,
        tour_template: TourTemplate,
    ):
        """
        Класс задания кроссворд
        :param words: слова в формате списка из множеств (слово, описание)
        :param max_height: максимальная высота (в пикселях или в ячейках)
        :param height_type: тип максимальной высоты
        """
        super().__init__(tour_template)

        self.words = words

        if height_type is HeightTypes.PX:
            self.max_height = self._px_to_cells(max_height)
        else:
            self.max_height = max_height

        self._words_clues = [
            [ComplexString(word.upper()), clue] for word, clue in words
        ]

    def _px_to_cells(self, px_size: int):
        cell_size_px = self.emu_to_pixels(self.CELL_SIZE, self.DPI)
        border_size_px = self.emu_to_pixels(self.BORDER_SIZE, self.DPI)
        res = int(px_size / cell_size_px - border_size_px * 2)
        if res <= 0:
            raise ValueError("CELL_SIZE and BORDER_SIZE are not compatible")
        else:
            return res

    def emu_to_pixels(self, emu: float, dpi: int) -> int:
        return int(Emu(int(emu)).inches * dpi)

    def make_docx(self, doc: Document):
        SECT_MAR = {
            "top": 250,
            "bottom": 250,
            "left": 720,
            "right": 720,
            "header": 708,
            "footer": 339,
            "gutter": 0,
        }
        CROSS_TIME_GENERATING = 0.02
        STYLE = "ReadingTask"
        doc = super().make_docx(doc)

        # =============================
        # Do img section
        img_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        ExtendedSection(img_sec).set_size_a4()
        ExtendedSection(img_sec).set_margins(**SECT_MAR)

        # generating crossword
        emu_width = ExtendedSection(img_sec).get_text_area_width()
        COLS = self._px_to_cells(self.emu_to_pixels(emu_width, self.DPI))
        ROWS = self.max_height
        print(f"grid sizes: {COLS} x {ROWS}")
        cross = Crossword(ROWS, COLS, available_words=self._words_clues)
        if not cross.validate():
            return
        res = cross.compute_crossword(CROSS_TIME_GENERATING)
        iter = 0  # TODO: delete iter
        while not res and self.CELL_SIZE >= Mm(5):
            print(iter)
            iter += 1
            self.CELL_SIZE -= Mm(1)
            self.BORDER_SIZE -= Mm(0.05)
            COLS = self._px_to_cells(self.emu_to_pixels(emu_width, self.DPI))
            ROWS = self.max_height

            cross = Crossword(ROWS, COLS, available_words=self._words_clues)
            if not cross.validate():
                return
            res = cross.compute_crossword(CROSS_TIME_GENERATING)

        print("Crossword computed" if res else "Too much words, failure")
        print(*cross.best_grid, sep="\n")
        cross.remove_blank_lines()

        print("Start generating image")
        img = cross.gen_img(
            self.emu_to_pixels(self.CELL_SIZE, self.DPI),
            self.emu_to_pixels(self.BORDER_SIZE, self.DPI),
        )
        print("Image generated")
        img_bytes = BytesIO()
        img.save(img_bytes, format="JPeG")

        doc.add_picture(img_bytes, width=Emu(emu_width))

        # =============================
        # Do clues section
        clues_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        ExtendedSection(clues_sec).set_size_a4()
        ExtendedSection(clues_sec).set_margins(**SECT_MAR)
        ExtendedSection(clues_sec).set_cols(2)
        print(f"len(best_wordlist) is 16: {len(cross.best_wordlist) == 16}")

        for word in cross.best_wordlist:
            try:
                print(word[5], word[4])
            except IndexError:
                print(*cross.best_wordlist, sep="\n")
                print(*cross.best_grid, sep="\n")
                print(word)
                print("Error!!")
                raise IndexError
        words = sorted(cross.best_wordlist.copy(), key=lambda x: (x[5], x[4]))

        hor = doc.add_paragraph(style=STYLE)
        hor.add_run("По горизонтали:").bold = True
        ExtendedParagraph(hor).set_borders(24, 1, "#000000", Direction.BOTTOM)

        hor_clues = doc.add_paragraph(style=STYLE)
        ExtendedParagraph(hor_clues).set_jc(JcTypes.BOTH)
        for word in words:
            word, clue, x, y, align, num = tuple(word)
            if align == 1:  # vertical words
                continue
            r = hor_clues.add_run(f"{num}. {clue}")
            r.add_break()

        ver = doc.add_paragraph(style=STYLE)
        ver.add_run("По вертикали:").bold = True
        ExtendedParagraph(hor).set_borders(24, 1, "#000000", Direction.BOTTOM)

        ver_clues = doc.add_paragraph(style=STYLE)
        ExtendedParagraph(hor_clues).set_jc(JcTypes.BOTH)
        for word in words:
            word, clue, x, y, align, num = tuple(word)
            if align == 0:  # horizontal words
                continue
            r = ver_clues.add_run(f"{num}. {clue}")
            r.add_break()

        self.doc = doc
