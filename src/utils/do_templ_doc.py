from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, ns
from docx.shared import Cm, Pt

from Configs import PathConfig
from src.utils.docx_documents_utils import do_rulang_xml


def do_template_docx():
    doc = Document()
    section = doc.sections[0]

    doc.core_properties.language = "ru-RU"

    # Поля
    section.left_margin = Cm(1.27)
    section.top_margin = Cm(0.5)
    section.right_margin = Cm(1.27)
    section.bottom_margin = Cm(1.0)

    styles = []

    # Стили
    e_test_style = doc.styles.add_style('ETestTask', WD_STYLE_TYPE.PARAGRAPH)
    e_test_style.font.name = 'Times New Roman'
    e_test_style.font.size = Pt(10)
    styles.append(e_test_style)

    ovio_header_style = doc.styles.add_style('OVIOHeader', WD_STYLE_TYPE.PARAGRAPH)
    ovio_header_style.font.name = 'Times New Roman'
    ovio_header_style.font.size = Pt(12)
    styles.append(ovio_header_style)

    reading_style = doc.styles.add_style('ReadingTask', WD_STYLE_TYPE.PARAGRAPH)
    reading_style.font.name = 'Times New Roman'
    reading_style.font.size = Pt(10)
    styles.append(reading_style)



    # Заменяем язык на русский во всех стилях
    for style in styles:
        style._element.get_or_add_rPr().append(do_rulang_xml())

    doc.save(str(PathConfig.TEMPL_PATH))


if __name__ == "__main__":
    do_template_docx()