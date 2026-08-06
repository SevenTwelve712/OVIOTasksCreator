from pathlib import Path


class PathConfig:
    BASE_DIR = Path(__file__).parent.parent
    TEMPL_PATH = Path(BASE_DIR, "resources", "BaseTempl.docx")
    SAVE_DIR = Path(BASE_DIR, "tests", "result_files")
    RESOURCES_DIR = Path(BASE_DIR, "resources")

