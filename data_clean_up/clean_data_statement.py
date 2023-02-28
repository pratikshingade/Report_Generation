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
            raise TypeError("File Format is Not Supported need .xlsx or .xls file")
        self.data = pd.read_excel(self.data_path, dtype={"SurveyNumber": str})
        return self.data

    def data_clean(self):
        """Clean Up Data To Desired Requirement"""
        if "Source" not in self.data:
            self.data["Source"] = ""

        self.data = self.data[["SurveyNumber", "Crop_Type", "Area_Ha", "Source"]]
        self.data['Area_Ha'] = self.data['Area_Ha'].apply(lambda x: round(x, 2))

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

        self.data.rename(columns={"SurveyNumber": 'भूमापन क्रमांक', "Crop_Type": "पिक", "Area_Ha": "पिक क्षेत्र",
                                  "Source": "स्त्रोत"}, inplace=True)

        self.data['भूमापन क्रमांक'] = self.data['भूमापन क्रमांक'].fillna(" ")
        self.data.loc[self.data['भूमापन क्रमांक'] == "", 'भूमापन क्रमांक'] = " "

        data_melt = pd.melt(self.data, id_vars=['भूमापन क्रमांक', "पिक", "स्त्रोत"], value_name="पिक क्षेत्र (ha)")

        # if data_melt["स्त्रोत"].isna().any():
        #     data_melt["स्त्रोत"] = ""
        data_melt["स्त्रोत"] = ""
        data_pivot_table = data_melt.pivot_table(index=['भूमापन क्रमांक', "स्त्रोत"], columns="पिक", aggfunc='sum',
                                                 margins=True, margins_name='मंजुरी क्षेत्र (ha)')

        return data_pivot_table
