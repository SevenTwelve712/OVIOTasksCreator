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

class JcTypes(Enum):
    """Класс типов выравниваний параграфа. Варианты:\n
    1) BOTH -> выравнивание относительно обоих краев
    2) CENTER -> выравнивание по центру
    3) END -> выравнивание по левому краю"""
    BOTH ="both"
    CENTER = "center"
    END = "end"

class Direction(Enum):
    BOTTOM = "bottom"
    LEFT = "left"
    RIGHT = "right"
    TOP = "top"
    ALL = "all"