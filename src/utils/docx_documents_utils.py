from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn


def set_cols(section, num):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num))

def do_rulang_xml():
    lang = OxmlElement('w:lang')
    lang.set(ns.qn('w:val'), 'ru-RU')
    return lang