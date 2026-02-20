from docx.oxml import OxmlElement, ns

def do_rulang_xml():
    lang = OxmlElement('w:lang')
    lang.set(ns.qn('w:val'), 'ru-RU')
    return lang