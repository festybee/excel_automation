# -*- coding: utf-8 -*-
"""
Created on Sun Aug  1 23:34:47 2021

@author: User
"""
import pandas as pd
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file


def automation(file):
    overdue_days = [31, 32, 33, 34, 35]
    channel = input("Enter Biz Channel (lcash or cashlion or Nairaplus): ")
    biz_channel = [channel]
    # channel = ['Lcash', 'Cashlion', 'Nairaplus']

    excel_file = file
    df1 = pd.read_excel(excel_file)
    df2 = df1.loc[(df1['Overdue Days'].isin(overdue_days))]
    # biz_channel = [biz_channel]
    df3 = df2.loc[(df2['Biz Channel'].str.lower().isin(biz_channel))]
    writer = pd.ExcelWriter(channel + '.xlsx')
    df3.to_excel(writer)
    writer.save()
    print(channel.capitalize() + '.xlsx file has been created.')


if __name__ == '__main__':
    automation(filename)
