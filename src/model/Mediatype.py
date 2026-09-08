from pathlib import Path

from docx.document import Document

from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.section import Section
from docx.shared import Cm, Pt, Mm, Emu, Twips
from PIL import Image

from loguru import logger
from model.Task import Task
from model.TourTemplate import ExtendedSection, TourTemplate
from model.extended_docx_classes.ExtendedParagraph import ExtendedParagraph
from model.extended_docx_classes.ExtendedTable import ExtendedTable
from model.extended_docx_classes.data_and_enums import JcTypes, TblBorder
# TODO: пофиксить размер чеек для записи слов


class Mediatype(Task):
    name = "Медианаборщик"
    write_cell_size = Cm(0.5)

    CELL_PADDING = Mm(2)

    def __init__(
        self,
        word: str,
        word_imgs: list[tuple[str, Path]],
        tour_template: TourTemplate,
    ):
        super().__init__(tour_template)
        self.word = word
        self.word_imgs = word_imgs
        self.font_size = 12

    def _clc_img_height(self, sec: Section):
        """Считает высоту одного изображения в Emu"""
        page_height = Mm(297)
        header_height = Cm(2.5)
        descr_height = Cm(1.5)
        table_height = (
            page_height
            - header_height
            - descr_height
            - sec.bottom_margin.emu
            - sec.top_margin.emu
        )
        logger.debug(f"table height: {round(Emu(table_height).cm, 3)}")

        all_imgs_height = (
            table_height
            - self.write_cell_size * 4  # учет ячеек для записи
            - self.CELL_PADDING * 2 * 8  # учет отступов ячейки
            - Pt(self.font_size) * 1.2 * 4  # учет текста
        )
        img_height = int(all_imgs_height / 4 - Mm(2))  # еще 1 мм запаса
        logger.debug(round(Emu(img_height).cm, 2))
        return img_height

    def _clc_image_width(self, sec: Section):
        workarea_width = ExtendedSection(sec).get_text_area_width()
        img_width = (workarea_width - self.CELL_PADDING * 2 * 4) // 4
        return img_width

    def _get_img_sizes_fit_bounds(
        self, img: Path, max_width_pt: float, max_height_pt: float
    ):
        # 1. Считываем оригинальные размеры картинки в пикселях
        orig_w, orig_h = Image.open(img).size

        # 2. Считаем соотношение сторон (Aspect Ratio)
        aspect_ratio = orig_w / orig_h

        # 3. Проверяем, в какое ограничение упрется картинка первым:
        # Если ширина при max_height превышает max_width — упираемся в ширину,
        # иначе — упираемся в высоту.
        if max_height_pt * aspect_ratio > max_width_pt:
            # Ограничивающим фактором стала ширина
            final_w = Pt(max_width_pt)
            final_h = Pt(max_width_pt / aspect_ratio)
        else:
            # Ограничивающим фактором стала высота
            final_w = Pt(max_height_pt * aspect_ratio)
            final_h = Pt(max_height_pt)
        return final_w, final_h

    def make_docx(self, doc: Document):
        doc = super().make_docx(doc)

        STYLE = "ReadingTask"
        SECT_MAR = {
            "top": 284,
            "bottom": 180,
            "left": 567,
            "right": 851,
            "header": 0,
            "footer": 0,
            "gutter": 0,
        }

        sec = doc.add_section(WD_SECTION.CONTINUOUS)
        ExtendedSection(sec).set_size_a4()
        ExtendedSection(sec).set_margins(**SECT_MAR)

        # make description
        par = doc.add_paragraph(style=STYLE)
        ExtendedParagraph(par).set_jc(JcTypes.CENTER)
        par.add_run(f"Из букв слова ").bold = True
        r = par.add_run(f"({self.word})")
        r.font.size = Pt(20)
        r.bold = True
        par.add_run(
            " составьте слова, соответствующие изображениям и указанному числу букв, и подпишите их под картинками. Слова идут в алфавитном порядке."
        ).bold = True

        # make table
        max_img_width_pt = Emu(self._clc_image_width(sec)).pt
        max_img_height_pt = Emu(self._clc_img_height(sec)).pt
        tbl = doc.add_table(8, 4)
        ExtendedTable(tbl).set_cell_margins([self.CELL_PADDING.twips] * 4)
        ExtendedTable(tbl).set_all_cells_borders(TblBorder(sz=4, val="single"))

        for num, (word, img) in enumerate(self.word_imgs):
            cell = tbl.cell(2 * (num // 4), num % 4)
            w, h = self._get_img_sizes_fit_bounds(
                img, max_img_width_pt, max_img_height_pt
            )
            cell.paragraphs[0].add_run().add_picture(str(img.absolute()), w, h)
            cell.add_paragraph().add_run(
                f"{num + 1}. {len(word)} {'буквы' if len(word) < 5 else 'букв'}"
            ).bold = True
            ExtendedParagraph(cell.paragraphs[1]).set_jc(JcTypes.CENTER)
            for par in cell.paragraphs:
                ExtendedParagraph(par).rm_spacings()

        for num, row in enumerate(tbl.rows):
            if num % 2:
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                row.height = self.write_cell_size
        tail_p = doc.add_paragraph()
        ExtendedParagraph(tail_p).rm_spacings()
        tail_p.paragraph_format.line_spacing = Pt(1)
        r_tail = tail_p.add_run()
        r_tail.font.size = Pt(1)

        self.doc = doc
