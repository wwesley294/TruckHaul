import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
import warnings


class WaitTime(object):

    def __init__(self, rac_wt, year, month):

        self.rac_wt = rac_wt
        self.year = year
        self.month = month
        return

    def wt_summary(self):

        in_columns = ["Location", "Company", "RR", "TruckTix", "PinkTix", "Material", "StartTime", "EndTime",
                      "TotalTime", "Billable", "Item", "Reason", "Penalty"]
        out_columns = ["Date", "Location", "Company", "RR", "TruckTix", "PinkTix", "Material", "StartTime", "EndTime",
                       "TotalTime", "Billable", "Item", "Reason", "Penalty"]

        path = self.rac_wt + "DWM-WT_" + self.year + self.month + ".xlsx"
        df_wt = pd.DataFrame()
        xlsx = pd.ExcelFile(path)
        sheet_list = xlsx.sheet_names

        for sheet in sheet_list:
            df_temp = pd.read_excel(path, sheet_name=sheet)
            df_temp = df_temp.iloc[3:, 0:13]
            df_temp.columns = in_columns

            time_str = sheet + " " + self.year
            date = datetime.strptime(time_str, "%b %d %Y")
            df_temp.insert(0, "Date", date)
            df_temp.insert(1, "Type", "")
            df_temp["TruckTix"] = pd.to_numeric(df_temp["TruckTix"])
            df_temp = df_temp[df_temp["TruckTix"] > 0]
            df_temp["StartTime"] = pd.to_datetime(df_temp["StartTime"], format="%H:%M:%S")
            df_temp["EndTime"] = pd.to_datetime(df_temp["EndTime"], format="%H:%M:%S")
            df_wt = df_wt.append(df_temp)
            print("Extraction completed: " + sheet)

        df_export = pd.DataFrame()
        df_wt["Location"] = df_wt["Location"].astype(str)
        temp_loc = "no loc"

        for i, row in df_wt.iterrows():

            if row["Location"] != "nan":
                temp_loc = row["Location"]
            else:
                row["Location"] = temp_loc

            temp_start = datetime.combine(row["Date"], row["StartTime"].time())
            temp_end = datetime.combine(row["Date"], row["EndTime"].time())

            if temp_end < temp_start:
                temp_end += timedelta(hours=12)

            row["StartTime"] = temp_start
            row["EndTime"] = temp_end
            row["TotalTime"] = temp_end - temp_start
            df_export = df_export.append(row)

        df_export = df_export.reindex(columns=out_columns)
        df_export = df_export.sort_values(by=["Date", "Location", "StartTime"])

        outpath = self.rac_wt + "DWM-WT_DetailList_" + self.year + self.month + "_working.xlsx"
        writer = pd.ExcelWriter(outpath)
        df_export.to_excel(writer, sheet_name="Sheet1")
        writer.save()


class Breakdown(object):

    def __init__(self, rac_wt, foreman, year, month):
        self.rac_wt = rac_wt
        self.foreman = foreman
        self.year = year
        self.month = month
        return

    def wt_foreman(self):
        df_foreman = pd.read_excel(self.foreman, sheet_name="Sheet1")
        w_list = df_foreman[df_foreman["Type"] == "Water"]["Foreman"].tolist()
        s_list = df_foreman[df_foreman["Type"] == "Sewer"]["Foreman"].tolist()
        return w_list, s_list

    def wt_fuse(self):

        in_columns = ["Location", "Foreman", "BES", "Company", "RR", "TruckTix", "PinkTix", "Material", "StartTime",
                      "EndTime", "TotalTime", "Billable", "Item", "Reason", "Penalty"]

        df_fuse = pd.DataFrame()
        path = self.rac_wt + "DWM-WT_BES_" + self.year + self.month + ".xlsx"
        xlsx = pd.ExcelFile(path)
        sheet_list = xlsx.sheet_names

        for name in sheet_list:
            if name != "Master":
                df_temp = pd.read_excel(path, sheet_name=name)
                df_temp = df_temp.iloc[:, 0:15]
                df_temp.columns = in_columns
                time_str = name + " " + self.year
                date = datetime.strptime(time_str, "%b %d %Y")
                df_temp = df_temp.iloc[3:]
                df_temp.insert(0, "Date", date)
                df_temp.insert(1, "Type", "")
                df_temp["TruckTix"] = pd.to_numeric(df_temp["TruckTix"])
                df_temp = df_temp[df_temp["TruckTix"] > 0]
                df_temp["StartTime"] = pd.to_datetime(df_temp["StartTime"], format="%H:%M:%S")
                df_temp["EndTime"] = pd.to_datetime(df_temp["EndTime"], format="%H:%M:%S")
                df_fuse = df_fuse.append(df_temp)
                print("Extraction completed: " + name)

        return df_fuse

    def wt_cleaner(self, df, w_list, s_list):

        out_columns = ["Date", "Location", "Foreman", "BES", "Company", "RR", "TruckTix", "PinkTix", "Material",
                       "StartTime", "EndTime", "TotalTime", "Billable", "Item", "Reason", "Penalty", "Type", ]

        df_temp = df.fillna(method="ffill")
        df_clean = pd.DataFrame()

        for i, row in df_temp.iterrows():
            temp_start = datetime.combine(row["Date"], row["StartTime"].time())
            temp_end = datetime.combine(row["Date"], row["EndTime"].time())
            if temp_end < temp_start:
                temp_end += timedelta(hours=12)
            row["StartTime"] = temp_start
            row["EndTime"] = temp_end
            temp_total = temp_end - temp_start
            row["TotalTime"] = temp_total.total_seconds() / 60
            # Determines wait time type
            if row["Foreman"] in w_list:
                row["Type"] = "Water"
            elif row["Foreman"] in s_list:
                row["Type"] = "Sewer"
            else:
                warning = "Please check if foreman $" + str(row["Foreman"]) + "$ is in the database!"
                print(warning)
            df_clean = df_clean.append(row)

        warnings.filterwarnings("ignore")
        df_clean = df_clean.reindex(columns=out_columns)
        bes = df_clean["BES"].drop_duplicates().to_list()
        for j in bes:
            df_norm = df_clean[df_clean["BES"] == j]
            loc = df_norm["Location"].mode()[0]
            df_clean["Location"][df_clean["BES"] == j] = loc
        df_clean = df_clean.sort_values(by=["Type", "Date"])

        outpath = self.rac_wt + "MonthlyCleaned_" + self.year + self.month + ".xlsx"
        writer = pd.ExcelWriter(outpath)
        df_clean.to_excel(writer, sheet_name="Sheet1")
        writer.save()
