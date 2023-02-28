import os
import time

from win32com import client

from excel_report_generation.title_formatting_summary import TitleFormat
from settings import setting_summary


class GeneratePdfSummary(TitleFormat):
    def __init__(self, csv_path):
        super(GeneratePdfSummary, self).__init__(csv_path)
        self.excel = client.Dispatch(setting_summary.PDF_GENERATION_APPLICATION)
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
        self.pdf_path = os.path.splitext(self.excel_path)[0] + setting_summary.PDF_FORMAT
        print(self.pdf_path)
        self.work_sheets.ExportAsFixedFormat(0, self.pdf_path)
        self.work_book.Close(True)
        time.sleep(1.5)
