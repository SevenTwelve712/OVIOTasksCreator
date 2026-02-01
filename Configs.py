from pathlib import Path

class PathConfig:
    BASE_DIR = Path(__file__).parent
    TEMPL_PATH = Path(BASE_DIR, "resources", "ETestTempl.docx")