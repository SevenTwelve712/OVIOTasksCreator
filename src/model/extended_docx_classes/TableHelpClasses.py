from dataclasses import dataclass
from enum import Enum


class WidthTypes(Enum):
    DXA = "dxa"
    AUTO = "auto"

class LayoutTypes(Enum):
    FIXED = "fixed"
    AUTOFIT = "autofit"

@dataclass
class TblBorder:
    """Класс для хранения данных о конкретной границе таблицы (left, top etc).
    Для подробной информации смотри http://officeopenxml.com/WPtableBorders.php"""
    color: str="auto"
    space: int=0
    sz: int=0
    val: str="none"