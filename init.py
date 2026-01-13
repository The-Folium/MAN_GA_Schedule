import excel_parser as ep
from data_base import DataBase
import openpyxl
from config import *

db = DataBase()
wb_hours = openpyxl.load_workbook(hours_filename, data_only=True)
ws_hours = wb_hours.active

ep.process_table(ws_hours, db)
wb_preferences = ep.create_preferences_workbook(db, preferences_filename)

