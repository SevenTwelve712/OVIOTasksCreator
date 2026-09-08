from enum import EnumType
from io import BytesIO
from PIL.Image import Image
from docx.document import Document

from docx.enum.section import WD_SECTION
from docx.shared import Emu, Inches, Mm, Pt, Twips

from model.extended_docx_classes.data_and_enums import Direction, JcTypes
from model.extended_docx_classes.ExtendedParagraph import ExtendedParagraph
from model.extended_docx_classes.ExtendedSection import ExtendedSection
from model.Task import Task
from model.TourTemplate import TourTemplate
from model.vendor.complexstring import ComplexString
from model.vendor.genxword import Crossword
from loguru import logger
# TODO: clean logs, do feature when user can generate crossword with his own params not only auto


class HeightTypes(EnumType):
    PX = "pixels"
    CELLS = "cells"


class NotValidCrosswordException(Exception):
    """вызывается когда кросворд невалиден"""


class ComputeCrosswordError(Exception):
    """Вызывается когда произошла ошибка при генерации кроссворда"""


# TODO: test
class OVIOCrossword(Task):
    name = "Кроссворд"
    cond = ""

    CELL_SIZE = Mm(13).emu  # размер одной ячейки кроссворда в emu для генерации jpg
    BORDER_SIZE = Mm(0.5).emu  # толщина границы вокруг каждой ячейки в emu
    DPI = 600
    CROSS_TIME_GENERATING = 0.1

    def __init__(
        self,
        words: list[tuple[str, str]],
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
        self.words.sort(key=lambda x: len(x[0]), reverse=True)
        self.font_size = 12

        self.img_height = 0
        self.img_width = 0

        self._words_clues = [
            [ComplexString(word.upper()), clue] for word, clue in words
        ]

        logger.debug("got a crossword with data:")
        logger.debug("\n".join(f"{word.upper()}: {clue}" for word, clue in self.words))

    def _clc_img_height(self, margins: dict[str, int], font_size: int):
        page_width = Mm(210)
        page_height = Mm(297)
        columns_gap = Twips(720)

        header_height = Pt(65)

        letter_width_coeff = 0.5

        total_syms = sum(len(clue) + 4 for word, clue in self.words)
        logger.debug(f"clues total syms: {total_syms}")
        clue_column_width = Mm(
            (
                page_width.mm
                - Twips(margins["left"]).mm
                - Twips(margins["right"]).mm
                - columns_gap.mm
            )
            / 2
        )
        one_line_capacity = int(
            clue_column_width.pt / (letter_width_coeff * font_size)
        )  # сколько символов вмещается в одной строке одного столбца
        logger.debug(f"one line capcity: {one_line_capacity}")
        clue_rows_sum = sum(
            self._clc_clue_rows(one_line_capacity, clue) for word, clue in self.words
        )
        clue_rows_amount = (
            (clue_rows_sum / 2) + 4
        )  # чистое количество строк, занимаемых определениями, 4 запас на отступы и надписи по горизонтали и вертикали

        # for num, (word, clue) in enumerate(self.words):
        # logger.debug(
        #     f"{num} clue is {clue}, clues length is {len(clue)} and it takes up {self._clc_clue_rows(one_line_capacity, clue)} rows"
        # )
        logger.debug(f"all clues take up {clue_rows_amount} rows")

        clues_height = Pt(
            clue_rows_amount * 1.15 * font_size
        )  # 1.15 коэффициент высоты линии (стандартный отступ меджду строками)
        logger.debug(
            f"clues take up {clues_height.mm} mm, its {round(clues_height / page_height * 100, 2)} % of page "
        )

        img_height = (
            page_height
            - header_height
            - Twips(margins["top"])
            - Twips(margins["bottom"])
            - clues_height
        )
        return img_height

    def _clc_clue_rows(self, line_capacity: int, clue: str):
        "Calulates how many lines will take the clue"
        ind = 0
        rows = 0
        while len(clue) > line_capacity:
            ind = line_capacity - 1
            if clue[ind] != " ":
                # find nearest whitespace
                while clue[ind] != " ":
                    ind -= 1

            clue = clue[ind + 1 :]
            rows += 1

        rows += 1
        return rows

    def _add_image_fit_bounds(
        self, img: Image, doc: Document, max_width_pt: float, max_height_pt: float
    ):
        # 1. Считываем оригинальные размеры картинки в пикселях
        orig_w, orig_h = img.size

        # 2. Считаем соотношение сторон (Aspect Ratio)
        aspect_ratio = orig_w / orig_h

        # 3. Проверяем, в какое ограничение упрется картинка первым:
        # Если ширина при max_height превышает max_width — упираемся в ширину,
        # иначе — упираемся в высоту.
        if max_height_pt * aspect_ratio > max_width_pt:
            # Ограничивающим фактором стала ширина
            final_w = Pt(max_width_pt)
            final_h = Pt(max_width_pt / aspect_ratio)

            # В python-docx достаточно задать только width для сохранения пропорций
            img_bytes = BytesIO()
            img.save(img_bytes, format="JPeG")
            doc.add_picture(img_bytes, width=final_w)
        else:
            # Ограничивающим фактором стала высота
            final_w = Pt(max_height_pt * aspect_ratio)
            final_h = Pt(max_height_pt)

            # В python-docx достаточно задать только height для сохранения пропорций
            img_bytes = BytesIO()
            img.save(img_bytes, format="JPeG")

            doc.add_picture(img_bytes, height=final_h)

        return final_w, final_h

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

    def _clc_recommended_crossword_sizes(self):
        """Считает рекомандованные размеры для кроссворда"""
        total_syms = sum(len(w) for w, c in self.words)
        cells_amount = total_syms * 1.3
        max_word_len = len(max([w for w, c in self.words], key=len))

        # делаем кроссворд чуть вытянутым
        height = int((cells_amount / 1.2) ** 0.5)
        width = int(1.2 * height) + 1

        if max_word_len > width:
            width = max_word_len
            height = int(cells_amount // width + 1)

        return (width, height)

    def _check_size_already_been(self, size: tuple, sizes: list[tuple]):
        sizes.sort(key=lambda x: x[0] * x[1])
        for cols, rows in sizes:
            if cols >= size[0] and rows >= size[1]:
                return True
        return False

    def _generate_crossword(self, img_sec, page_margins):
        sizes = []
        self.img_width = ExtendedSection(img_sec).get_text_area_width()

        for font_reduction in range(4):  # идем до 9 кегля включительно
            self.font_size = 12 - font_reduction
            self.img_height = self._clc_img_height(page_margins, self.font_size)
            self.CELL_SIZE = Mm(13)
            self.BORDER_SIZE = Mm(0.5)

            # если при текущем кегле получили что картинка не может сгенерироваться сразу переходим к следующему
            if self.img_height <= 0:
                logger.error(
                    f"With font size {self.font_size} img height is {round(self.img_height, 2)}"
                )
                continue

            for cell_size_reduction in range(0, 50, 3):
                self.CELL_SIZE = Mm(13 - cell_size_reduction / 10)
                self.BORDER_SIZE -= Mm(0.01)

                max_rows = int(self.img_height / self.CELL_SIZE)
                max_cols = int(self.img_width / self.CELL_SIZE)
                logger.debug(
                    f"Генерим кроссворд с размером клетки {round(Emu(int(self.CELL_SIZE)).mm, 4)} и шрифтом {self.font_size}"
                )
                logger.debug(f"Максимальные размеры сетки: {max_cols} x {max_rows}")

                # если такие размеры (или меньшие) уже были, то даже не пытаемся генерить
                if self._check_size_already_been((max_cols, max_rows), sizes):
                    logger.debug("Подобные размеры уже были")
                    continue

                cross_max = Crossword(
                    max_rows, max_cols, available_words=self._words_clues
                )
                if cross_max.validate() is False:
                    logger.debug("Кроссворд с такой сеткой не валиден")
                    continue

                # пытаемся сгенерить кроссворд с оптимальными размерами сетки
                rows, cols = self._clc_recommended_crossword_sizes()
                logger.debug(f"Оптимальные размеры сетки: {cols} x {rows}")

                if (
                    rows <= max_rows
                    and cols <= max_cols
                    and self._check_size_already_been((cols, rows), sizes) is False
                ):
                    cross_optimal = Crossword(
                        rows, cols, available_words=self._words_clues
                    )
                    if (
                        cross_optimal.compute_crossword(self.CROSS_TIME_GENERATING)
                        is True
                    ):
                        logger.info(
                            "Получилось сгенерировать кроссворд с оптимальной сеткой"
                        )
                        return cross_optimal
                    logger.debug(
                        "Не получилось сгенерировать кроссворд с оптимальной сеткой"
                    )
                    sizes.append((cols, rows))

                # если не получилось используем маскимальные размеры сетки
                if cross_max.compute_crossword(self.CROSS_TIME_GENERATING) is True:
                    logger.debug("Получилось сгенерить кроссворд с максимальной сеткой")
                    return cross_max
                logger.debug("Не получилось сгенерить кроссворд с максимальной сеткой")
                sizes.append((max_cols, max_rows))
                logger.debug("===== Переходим к следующей итерации ====")
        logger.error("Не удалось посчитать кроссворд")
        raise ComputeCrosswordError

    def _setup_crossword(self, img_width, img_height):
        logger.info("setting up crossword")
        COLS = img_width // self.CELL_SIZE
        ROWS = img_height // self.CELL_SIZE
        logger.info(f"crossword sizes: {COLS} x {ROWS}")
        logger.info(
            f"crossword img sizes: {round(Emu(img_width).mm, 2)} x {round(Emu(img_height).mm, 2)}"
        )
        cross = Crossword(ROWS, COLS, available_words=self._words_clues)

        if not cross.validate():
            logger.error("crossword is not valid")
            return False

        logger.info("=====================")

        return cross

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
        STYLE = "ReadingTask"
        doc = super().make_docx(doc)

        # =============================
        # Do img section
        img_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        ExtendedSection(img_sec).set_size_a4()
        ExtendedSection(img_sec).set_margins(**SECT_MAR)

        cross = self._generate_crossword(img_sec, SECT_MAR)

        img = cross.gen_img(
            self.emu_to_pixels(self.CELL_SIZE, self.DPI),
            self.emu_to_pixels(self.BORDER_SIZE, self.DPI),
        )
        logger.info("img generated")

        self._add_image_fit_bounds(
            img, doc, Emu(self.img_width).pt, Emu(self.img_height).pt
        )
        # =============================
        # Do clues section
        clues_sec = doc.add_section(WD_SECTION.CONTINUOUS)
        ExtendedSection(clues_sec).set_size_a4()
        ExtendedSection(clues_sec).set_margins(**SECT_MAR)
        ExtendedSection(clues_sec).set_cols(2)

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
            r.font.size = Pt(self.font_size)
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
            r.font.size = Pt(self.font_size)
            r.add_break()

        self.doc = doc
