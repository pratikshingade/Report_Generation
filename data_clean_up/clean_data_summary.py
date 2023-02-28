import os
import pandas as pd


class CleanData:

    def __init__(self, data_path):
        self.data = None
        self.data_path = data_path
        self.read_data()

    def read_data(self):
        """Read Crop Statement Data From CSV"""
        if os.path.splitext(self.data_path)[1] not in [".xlsx", ".xls"]:
            raise TypeError("File Format is Not Supported need .xlsx file")
        self.data = pd.read_excel(self.data_path)
        return self.data

    def data_clean(self):
        """Clean Up Data To Desired Requirement"""
        self.data = self.data[["Crop_Type", "FREQUENCY", "SUM_Area_Ha"]]
        self.data['SUM_Area_Ha'] = self.data['SUM_Area_Ha'].apply(lambda x: round(x, 2))

        self.data.loc[self.data["Crop_Type"] == 1, "Crop_Type"] = "ऊस"
        self.data.loc[self.data["Crop_Type"] == 2, "Crop_Type"] = "द्राक्षे"
        self.data.loc[self.data["Crop_Type"] == 3, "Crop_Type"] = "मका"
        self.data.loc[self.data["Crop_Type"] == 4, "Crop_Type"] = "ज्वारी"

        self.data.loc[self.data["Crop_Type"] == 5, "Crop_Type"] = "गहू"
        self.data.loc[self.data["Crop_Type"] == 6, "Crop_Type"] = "फळभाजी"
        self.data.loc[self.data["Crop_Type"] == 7, "Crop_Type"] = "पालेभाजी"
        self.data.loc[self.data["Crop_Type"] == 8, "Crop_Type"] = "केळी"

        self.data.loc[self.data["Crop_Type"] == 9, "Crop_Type"] = "हळद"
        self.data.loc[self.data["Crop_Type"] == 10, "Crop_Type"] = "सोयाबीन"
        self.data.loc[self.data["Crop_Type"] == 11, "Crop_Type"] = "बागायत"
        self.data.loc[self.data["Crop_Type"] == 12, "Crop_Type"] = "इतर"

        self.data.rename(
            columns={"Crop_Type": "पिक", "FREQUENCY": "पिकांची वारंवारता", "SUM_Area_Ha": "मोजणी क्षेत्र (Ha)"},
            inplace=True)
        self.data.index += 1
        self.data.index.name = 'अ. क्र.'

        self.data["ड्रोनद्वारे मोजणी क्षेत्र (Ha)"] = self.data["मोजणी क्षेत्र (Ha)"]
        self.data["आकारणी  क्षेत्र (Ha)"] = self.data["मोजणी क्षेत्र (Ha)"]

        return self.data
