import pandas as pd
from numpy.ma.core import append
import openpyxl
from openpyxl import load_workbook
import csv
import os
from future.backports.datetime import datetime
from datetime import datetime
from copy import copy
import time
import os
import PySimpleGUI as sg
import os

if __name__=="__main__":
    layout = [[sg.Text("选择文件 扇区通道查询结果-泰山:")],
              [sg.InputText(key='-FILE_A-', enable_events=True), sg.FilesBrowse()],
              [sg.Text("选择文件 扇区通道查询结果-新网:")],
              [sg.InputText(key='-FILE_B-', enable_events=True), sg.FilesBrowse()],
              [sg.Text("选择文件 MML报文解析-泰山:")],
              [sg.InputText(key='-FILE_C-', enable_events=True), sg.FilesBrowse()],
              [sg.Text("选择文件 MML报文解析-新网:")],
              [sg.InputText(key='-FILE_D-', enable_events=True), sg.FilesBrowse()],
              [sg.Button('提交')]]

    window = sg.Window('选择所需文件', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == 'Cancel':
            break
        if event == '提交':
            sheet4_path_mr = values['-FILE_A-']
            sheet4_path_zl_4 = values['-FILE_B-']
            sheet4_path_zl_5 = values['-FILE_C-']
            sheet4_path_zl_pp_4 = values['-FILE_D-']

            print(path_str1)

            window.close()
        if event == sg.WIN_CLOSED:
            break