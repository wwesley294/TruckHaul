import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta


class WaitTime(object):

    def __init__(self, rac_wt, year, month):

        self.rac_wt = rac_wt
        self.year = year
        self.month = month
        return

    def wt_summary(self):
        
        in_columns = ["Location", "Company", "RR", "TruckTix", "PinkTix", "Material", "StartTime", "EndTime", "TotalTime", "Billable", "Item", "Reason", "Penalty"]
        out_columns = ["Date", "Location", "Company", "RR", "TruckTix", "PinkTix", "Material", "StartTime", "EndTime", "TotalTime", "Billable", "Item", "Reason", "Penalty"]
        

        path = self.rac_wt + "DWM-WT_" + self.year + self.month + ".xlsx"
        df_wt = pd.DataFrame()
        xlsx = pd.ExcelFile(path)
        sheet_list = xlsx.sheet_names
        
        for sheet in sheet_list:
            
            df_temp = pd.read_excel(path, sheet_name=sheet)
            df_temp = df_temp.iloc[3:,0:13]
            df_temp.columns = in_columns
            df_temp = df_temp.dropna(subset=["TruckTix"])

            time_str = sheet + " " + self.year
            date = datetime.strptime(time_str, "%b %d %Y")
            df_temp.insert(0, "Date", date)
            df_temp.insert(1, "Type", "")
            df_temp["TruckTix"] = pd.to_numeric(df_temp["TruckTix"])
            df_temp['StartTime'] = pd.to_datetime(df_temp['StartTime'], format="%H:%M:%S")
            df_temp['EndTime'] = pd.to_datetime(df_temp['EndTime'], format="%H:%M:%S")
            df_wt = df_wt.append(df_temp)
            print("Extraction completed: " + sheet)


        df_export = pd.DataFrame()
        df_wt["Location"] = df_wt["Location"].astype(str)
        temp_loc = "no loc"
        
        for i, row in df_wt.iterrows():
            
            if row["Location"]!="nan":
                temp_loc = row["Location"]
            else:
                row["Location"] = temp_loc
                
            temp_start = datetime.combine(row["Date"], row["StartTime"].time())
            temp_end = datetime.combine(row["Date"], row["EndTime"].time())
            
            if temp_end < temp_start:
                temp_end += timedelta(hours = 12)
                
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

    # def wt_breakdown(self):

        
        
