import pandas as pd
from rpa_bot.log import lte, log
import os


class ExcelProcess:
    """
    The ExcelProcess class is designed to handle operations related to Excel files.

    Attributes:
    -----------
    path : str
        The path where the Excel file is located.

    Methods:
    --------
    __init__(self, path: str):
        Initializes the ExcelProcess with the specified path.

    lastBusinessDay(self) -> datetime.date:
        Returns the last business day of the current month.

    get_calendar(self) -> pd.DataFrame:
        Reads the 'Calendar.xlsx' file from the specified path and filters the rows where 'Entity' is 'Eisen'.
        Returns the filtered DataFrame.
    """
    def __init__(self, path):
        self.path = path

    def last_business_day(self):
        import datetime
        from pandas.tseries.offsets import BMonthEnd

        today = datetime.date.today()
        offset = BMonthEnd()
        last_bussines_day = offset.rollback(today)

        return last_bussines_day

    def get_calendar(self):
        try:
            df_calendar = pd.read_excel(os.path.join(self.path, 'Calendar.xlsx'))
            df_3B5 = df_calendar[df_calendar['Entity'] == '3B5']
            df_V436 = df_calendar[df_calendar['Entity'] == 'V436']
            
            return df_3B5, df_V436
        
        except Exception as e:
            log(f"Error in function get_calendar. {e}", lte.error)
    
    def close_excel(self):
        import subprocess
        subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])
