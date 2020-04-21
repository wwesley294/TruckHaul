import numpy as np
import pandas as pd
import os
from datetime import datetime
from TH_Merger import All_Merger, RAC_Merger, Uti_Merger
from TH_TixMiner import TixMiner
from TH_WaitTime import WaitTime, Breakdown
import sys

if __name__ == "__main__":

    # Space for All links
    all_in = "user_input"
    all_out = "user_input"

    # Space for RAC links
    rac_in = "user_input"
    rac_out = "user_input"
    unit_price = "user_input"

    # Space for RAC-WT links
    foreman = "user_input"
    rac_wt = "user_input"
    year = "2020"
    month = "Jan"

    # Space for Uti links
    uti_in = "user_input"
    uti_out = "user_input"


    def hostess():
        print("Today's special menu for detoxing with Gwyneth Paltrow: \n" + \
              "1) Allied Merge \n" + \
              "2) Reliable Merge \n" + \
              "3) Reliable Merge Summary \n" + \
              "4) Reliable TixMiner \n" + \
              "5) Reliable WT Summary  \n" + \
              "6) Reliable WT Breakdown \n" + \
              "7) Utility Merge \n" + \
              "8) Exit...Stage Left")
        user = input("Please enter your selection: ")

        # Option 1
        if user == "1":
            allied = All_Merger(all_in, all_out)
            allied.all_inspect()
            allied.all_zipper()

        # Option 2
        elif user == "2":
            reliable = RAC_Merger(rac_in, rac_out, unit_price)
            reliable.rac_inspect()
            reliable.rac_zipper()

        # Option 3
        elif user == "3":
            summary = RAC_Merger(rac_in, rac_out, unit_price)
            summary.rac_summary()

        # Option 4
        elif user == "4":
            miner = TixMiner(rac_out)
            pdf_list = miner.pdf_reader()
            orion_list = miner.merge_reader()
            miner.missing_tix(pdf_list, orion_list)

        # Option 5
        elif user == "5":
            wt = WaitTime(rac_wt, year, month)
            wt.wt_summary()

        # Option 6
        elif user == "6":
            bd = Breakdown(rac_wt, foreman, year, month)
            w_list, s_list = bd.wt_foreman()
            df = bd.wt_fuse()
            bd.wt_cleaner(df, w_list, s_list)

        # Option 7
        elif user == "7":
            utility = Uti_Merger(uti_in, uti_out)
            utility.uti_inspect()
            utility.uti_zipper()

        # Option 8
        elif user == "8":
            print("Careful Icarus.")
            sys.exit()

        # Loop back
        hostess()


    hostess()


