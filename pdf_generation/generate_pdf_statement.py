import os
import time

from win32com import client

from excel_report_generation.title_formatting_statement import TitleFormat
from settings import setting_statement


class GeneratePdfStatement(TitleFormat):
    def __init__(self, csv_path):
        super(GeneratePdfStatement, self).__init__(csv_path)
        self.excel = client.Dispatch(setting_statement.PDF_GENERATION_APPLICATION)
        self.excel.Interactive = False
        self.excel.Visible = False

        self.save_excel()
        self.pdf_path = None

        self.work_book = None
        self.work_sheets = None
        self.read_excel()

    def read_excel(self):
        self.work_book = self.excel.Workbooks.Open(self.excel_path)
        self.work_sheets = self.work_book.Worksheets[0]
        return self.work_sheets

    def create_pdf(self):
        self.pdf_path = os.path.splitext(self.excel_path)[0] + setting_statement.PDF_FORMAT
        print(self.pdf_path)
        self.work_sheets.ExportAsFixedFormat(0, self.pdf_path)
        self.work_book.Close(True)
        time.sleep(1.5)
