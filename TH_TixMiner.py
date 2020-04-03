from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBoxHorizontal
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
import pandas as pd
from datetime import datetime
import os
import io
import re


class TixMiner(object):

    def __init__(self, rac_out):
        
        self.rac_out = rac_out
        self.datestr = datetime.strftime(datetime.now(), "%m%d%Y")
        return
    
    def pdf_reader(self):
        
        pdf_list = []
        manager = PDFResourceManager()
        layout = LAParams()
        device = PDFPageAggregator(manager, laparams = layout)
        interpreter = PDFPageInterpreter(manager, device)
        
        file = input("Please provide the complete path to the target PDF file: ")
        target = open(file, "rb")
        for page in PDFPage.get_pages(target):
            interpreter.process_page(page)
            content = device.get_result()
            for element in content:
                if isinstance(element, LTTextBoxHorizontal):
                    temp = element.get_text()
                    temp = re.sub("[\n]", "", temp)
                    if len(temp) == 7 and temp.isdigit():
                        pdf_list.append(temp)

        return pdf_list

    def merge_reader(self):
        
        merge = self.rac_out + "Merge_" + self.datestr + ".xlsx"
        df = pd.read_excel(merge, sheet_name = "Sheet1")
        df["TixNum"] = df["TixNum"].astype(str)
        orion_list = list(df["TixNum"])
        
        return orion_list

    def missing_tix(self, pdf_list, orion_list):
        
        orion_miss = set(pdf_list) - set(orion_list)
        pdf_miss = set(orion_list) - set(pdf_list)
        export = self.rac_out + "MissingTix_" + self.datestr + ".xlsx"
        writer = pd.ExcelWriter(export)
        
        if orion_miss != set():
            df_orion = pd.DataFrame(orion_miss)
            df_orion.to_excel(writer, sheet_name="Missing in Orion")
            print("Missing ticket(s) in Orion detected... \n")
            
        if pdf_miss != set():
            df_pdf = pd.DataFrame(pdf_miss)
            df_pdf.to_excel(writer, sheet_name="Missing in PDF")
            print("Missing ticket(s) in PDF detected... \n")
            
        if (orion_miss == set()) and (pdf_miss == set()):
            print("We either have everything or everything is fucked... \n")
        else:            
            writer.save()
