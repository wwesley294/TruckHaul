import numpy as np
import pandas as pd
import os
import re
from datetime import datetime
import warnings


# All_Fighter extracts and transforms All truck haul ticket data so it can then be loaded into the Access database.
class All_Merger(object):

    def __init__(self, all_in, all_out):

        self.all_in = all_in
        self.all_out = all_out
        return

    # all_inspect iterates through all .xlsx files and sheets to ensure consistency.
    def all_inspect(self):

        # To return the numbers of column on a sheet as a preliminary screen.
        direc = os.listdir(self.all_in)
        for file in direc:
            if "summary" in file.lower():
                path = self.all_in + file
                xlsx = pd.ExcelFile(path)
                sheet_list = xlsx.sheet_names
                for bit in sheet_list:
                    df = pd.read_excel(path, sheet_name=bit)
                    print(str(df.shape[1]) + " columns detected in " + file + " - " + bit)

        user_check = input("Please press Enter to continue if no anomaly is detected... \n")

    # all_zipper extracts ticket data, categorizes it, and exports a Merge.
    def all_zipper(self):

        # Temporary column format.
        temp_columns = ["TransNum", "DateIn", "TimeIn", "TruckNum", "DumpTixNum", "Material", "ItemNum", "WeightIn", \
                        "WeightOut", "Quantity", "Source"]
        # Final column format for export.
        master_columns = ["Type", "Source", "TransNum", "DateIn", "TimeIn", "TruckNum", "DumpTixNum", "Material",
                          "ItemNum", \
                          "WeightIn", "WeightOut", "Quantity"]

        # To extract data from individual .xlsx file/sheet under the following directory.
        # To create a new DataFrame where the extracted data will be processed and stored.
        df_og = pd.DataFrame(columns=master_columns)
        direc = os.listdir(self.all_in)
        for file in direc:

            # To only select files with "summary" in their names.
            if "summary" in file.lower():
                lead = self.all_in + file
                xlsx = pd.ExcelFile(lead)
                sheet_list = xlsx.sheet_names

                # To specify the source and ticket type (W/S).
                if "water" in file.lower():
                    for item in sheet_list:
                        df_item = pd.read_excel(lead, sheet_name=item)
                        source = df_item.iloc[0, 0]
                        df_item = df_item.iloc[2:, 0:11]
                        df_item.columns = temp_columns
                        df_item = df_item.dropna(subset=["TransNum"], axis=0)
                        df_item["Source"] = source
                        df_item["Type"] = "W"
                        df_item = df_item.reindex(columns=master_columns)
                        df_og = df_og.append(df_item)
                        print("Extraction completed: " + file + "- " + item)
                elif "sewer" in file.lower():
                    for item in sheet_list:
                        df_item = pd.read_excel(lead, sheet_name=item)
                        source = df_item.iloc[0, 0]
                        df_item = df_item.iloc[2:, 0:11]
                        df_item.columns = temp_columns
                        df_item = df_item.dropna(subset=["TransNum"], axis=0)
                        df_item["Source"] = source
                        df_item["Type"] = "S"
                        df_item = df_item.reindex(columns=master_columns)
                        df_og = df_og.append(df_item)
                        print("Extraction completed: " + file + "- " + item)

        # To remove potential duplicates based on DumpTixNum.
        df_og = df_og.drop_duplicates(subset=["DumpTixNum"])
        df_og = df_og.sort_values(by=["Type", "DateIn", "TransNum"])

        # To export the DataFrame where processed ticket data is stored.
        # To save the results as Merge with today's date.
        # To date and export Merge.
        datestr = datetime.strftime(datetime.now(), "%m%d%Y")
        outpath = self.all_out + "Merge_" + datestr + ".xlsx"
        writer = pd.ExcelWriter(outpath)
        df_og.to_excel(writer, sheet_name="Sheet1")
        writer.save()
        print("A cada cerdo le llega su San Martín. \n")


# RAC_Fighter extracts and transforms Rel truck haul ticket data so it can then be loaded into the Access database.
# RAC_Fighter also summarized Reliable truck haul ticket data to help identify missing/extra ticket data.
class RAC_Merger(object):

    def __init__(self, rac_in, rac_out, unit_price):

        self.rac_in = rac_in
        self.rac_out = rac_out
        self.datestr = datetime.strftime(datetime.now(), "%m%d%Y")
        self.unit_price = unit_price
        return

        # rac_inspect iterates through all .xlsx files to ensure consistency..

    def rac_inspect(self):

        # To return the numbers of column on a sheet as a preliminary screen.
        direc = os.listdir(self.rac_in)
        for file in direc:
            if "summary" in file.lower():
                path = self.rac_in + file
                df = pd.read_csv(path)
                print(str(df.shape[1]) + " columns detected in " + file)

        user_check = input("Please press Enter to continue if no anomaly is detected... \n")

    # rac_zipper extracts ticket data, categorizes it, and exports a Merge.
    def rac_zipper(self):

        # Temporay column format.
        temp_columns = ["Date", "Dept", "PO", "Location", "Material", "ItemNum", "TixNum", "Quantity", "Truck"]
        # Final column format for export.
        final_columns = ["Date", "Dept", "PO", "Location", "Material", "ItemNum", "TixNum", "Quantity", "Time", "Truck"]

        # To extract data from individual .xlsx file under the following directory.
        # To create a new DataFrame where the extracted data will be processed and stored.
        df_og = pd.DataFrame()
        direc = os.listdir(self.rac_in)
        for file in direc:
            if "SUMMARY" in file:
                path = self.rac_in + file
                df_file = pd.read_csv(path)
                df_file = df_file.iloc[:, 0:9]
                df_file.columns = temp_columns
                df_file.insert(8, "Time", "")
                df_og = df_og.append(df_file)
                print("Extraction completed: " + file)

                # To remove potential duplicates based on TixNum and to re-index the columns to final form.
        df_og = df_og.drop_duplicates(subset=["TixNum"])
        df_og = df_og.reindex(columns=final_columns)

        # To set Date to datetime and to remove alphabets in ItemNum.
        df_og["Date"] = pd.to_datetime(df_og["Date"])
        df_og["ItemNum"] = df_og["ItemNum"].apply(lambda x: re.sub("[a-zA-Z]", "", x))

        df_og = df_og.sort_values(by=["Date", "TixNum"])

        # To date and export the processed ticket data.
        outpath = self.rac_out + "Merge_" + self.datestr + ".xlsx"
        writer = pd.ExcelWriter(outpath)
        df_og.to_excel(writer, sheet_name="Sheet1")
        writer.save()
        print("A pan de quince días, hambre de tres semanas. \n")

    # rac_summary summarizes the data in Merge.
    def rac_summary(self):

        # To extract data from Merge.
        ta_merge = self.rac_out + "Merge_" + self.datestr + ".xlsx"
        df_merge = pd.read_excel(ta_merge)
        df_merge['UnitPrice'] = np.nan

        # To extract unit price data of individual items.
        ta_up = self.unit_price
        df_up = pd.read_excel(ta_up)

        # To match items in Merge with corresponding unit prices.
        warnings.filterwarnings('ignore')
        for i, row in df_up.iterrows():
            item = row['ItemNum']
            price = row['UnitPrice']
            df_merge['UnitPrice'][df_merge['ItemNum'] == item] = price

        # To calculate the total cost based on the quantities and unit prices.
        df_merge['Amount'] = np.nan
        df_merge['Amount'] = df_merge['Quantity'] * df_merge['UnitPrice']
        df_merge['Amount'] = round(df_merge['Amount'], 2)
        df_sewer = df_merge[df_merge['Dept'].str.contains('SEWERS')]
        df_sewer = pd.pivot_table(df_sewer, index=['ItemNum'], values=['Quantity', 'Amount'], aggfunc=np.sum)
        df_water = df_merge[~df_merge['Dept'].str.contains('SEWERS')]
        df_water = pd.pivot_table(df_water, index=['ItemNum'], values=['Quantity', 'Amount'], aggfunc=np.sum)

        # To date and export the summary (water and sewer).
        outpath = self.rac_out + "SummaryCheck_" + self.datestr + ".xlsx"
        writer = pd.ExcelWriter(outpath)
        df_water.to_excel(writer, sheet_name='water')
        df_sewer.to_excel(writer, sheet_name='sewer')
        writer.save()
        print("A pan de quince días, hambre de tres semanas. \n")


# Uti_Fighter compiles all Uti truck haul ticket data for import into Access database.
class Uti_Merger(object):

    def __init__(self, uti_in, uti_out):

        self.uti_in = uti_in
        self.uti_out = uti_out
        return

    # uti_inspect iterates through all .xlsx files and sheets to ensure consistency.
    def uti_inspect(self):

        # To return the numbers of column on a sheet as a preliminary screen.
        direc = os.listdir(self.uti_in)
        for file in direc:
            if "report" in file.lower():
                path = self.uti_in + file
                xlsx = pd.ExcelFile(path)
                sheet_list = xlsx.sheet_names
                for bit in sheet_list:
                    df = pd.read_excel(path, sheet_name=bit)
                    print(str(df.shape[1]) + " columns detected in " + file)

        user_check = input("Please press Enter to continue if no anomaly is detected... \n")

    # uti_zipper extracts ticket data, categorizes it, and exports a Merge.
    def uti_zipper(self):

        # Temporary column format.
        temp_columns = ["Date", "DumpSite", "TixNum", "ItemNum", "DebrisTy", "WO", "Voucher", "Location", "Driver", \
                        "Water/Sewer", "Tonage", "Price", "Total"]
        # Final column format for export.
        master_columns = ["Type", "Date", "DumpSite", "TixNum", "ItemNum", "DebrisTy", "WO", "Voucher", "Location", \
                          "Driver", "Tonage", "Price", "Total"]

        # To extract data from individual .xlsx file/sheet under the following directory.
        # To create a new DataFrame where the extracted data will be processed and stored.
        df_og = pd.DataFrame(columns=temp_columns)
        direc = os.listdir(self.uti_in)
        for file in direc:
            if "REPORT" in file:
                path = self.uti_in + file
                xlsx = pd.ExcelFile(path)
                sheet_list = xlsx.sheet_names
                warnings.filterwarnings("ignore")
                for bit in sheet_list:
                    df_bit = pd.read_excel(path, sheet_name=bit)
                    df_bit = df_bit.iloc[2:, 0:13]
                    df_bit.columns = temp_columns
                    df_bit = df_bit.dropna(subset=["Date"], axis=0)
                    df_bit["Water/Sewer"] = df_bit["Water/Sewer"].str.lower()
                    df_bit["Type"] = np.nan
                    df_bit["Type"][df_bit["Water/Sewer"].str.contains("water")] = "W"
                    df_bit["Type"][df_bit["Water/Sewer"].str.contains("sewer")] = "S"
                    df_bit["Type"] = df_bit["Type"].fillna("Unknown")
                    df_og = df_og.append(df_bit)
                    print("Extraction completed: " + file + "- " + bit)

                    # To remove potential duplicates based on TixNum and to re-index the columns to final form.
        df_og = df_og.drop_duplicates(subset=["TixNum"])
        df_og = df_og.reindex(columns=master_columns)
        df_og = df_og.sort_values(by=["Type", "Date", "TixNum"])

        # To date and export the processed ticket data.
        datestr = datetime.strftime(datetime.now(), "%m%d%Y")
        outpath = self.uti_out + "Merge_" + datestr + ".xlsx"
        writer = pd.ExcelWriter(outpath)
        df_og = df_og[master_columns]
        df_og.to_excel(writer, sheet_name="Sheet1")
        writer.save()
        print("Más se consigue lamiendo que mordiendo. \n")


