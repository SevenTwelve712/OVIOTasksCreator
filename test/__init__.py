from datetime import date

from src.model.TourData import Tours, Forms
from src.model.TourTemplate import TourTemplate

tour_templ_ex = TourTemplate(Forms.OLD, Tours.FINAL, date(2023, 11, 4), "Москва")