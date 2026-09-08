from pathlib import Path

from docx import Document

from Configs import PathConfig
from model.Mediatype import Mediatype
from tests import tour_templ_ex

word = "МЕДНОГОРСК"
words = [
    "gore",
    "gorn",
    "gorod",
    "grom",
    "donor",
    "donos",
    "komod",
    "kondor",
    "korm",
    "koroed",
    "krem",
    "more",
    "mors",
    "nomer",
    "rodeo",
    "srok",
]
imgs = [
    Path(img).absolute()
    for img in Path(
        "/home/seventwelve712/PycharmProjects/OVIOTasksCreator/tests/mtimgs"
    ).iterdir()
]
words_img = list(zip(words, imgs))


def crossword_test(word, words_img):
    SAVE_PATH = Path(PathConfig.SAVE_DIR, "mediatype.docx")
    mt = Mediatype(word, words_img, tour_templ_ex)
    doc = Document(PathConfig.TEMPL_PATH)
    mt.make_docx(doc)
    doc.save(str(SAVE_PATH))


if __name__ == "__main__":
    crossword_test(word, words_img)
