from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, ns
from docx.shared import Cm, Pt


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



    # Заменяем язык на русский во всех стилях
    for style in styles:
        lang = OxmlElement('w:lang')
        lang.set(ns.qn('w:val'), 'ru-RU')
        style._element.get_or_add_rPr().append(lang)

    doc.save("../ETestTempl.docx")


if __name__ == "__main__":
    do_template_docx()