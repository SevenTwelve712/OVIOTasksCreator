from enum import Enum


class Forms(Enum):
    PRESCHOOL = "??"
    FIRST = "1"
    SECOND = "2"
    PRIMARY = "3-4"
    MID = "5-6"
    OLD = "7-11"


class Tours(Enum):
    SCHOOL = "Школьный тур"
    MUNICIPAL = "Муниципальный тур"
    REGIONAL = "Региональный тур"
    FINAL = "Финал"