import time
import xlwings as xw
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from Test import test3
import pandas as pd

def read_excel(file_path, sheet_name=0):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"读取Excel文件时出错: {str(e)}")
        return None

def read_txt(file_path, encoding='utf-8'):
    try:
        with open(file_path, 'r', encoding=encoding) as file:
            return file.read()
    except Exception as e:
        print(f"读取文件时出错: {str(e)}")
        return None

base_station_5 = read_excel('C:/Users/24253/Desktop/每月例行工作/工参/昌都245G工参总版-1028.xlsx', 'NR').values.tolist()
base_station_4 = read_excel('C:/Users/24253/Desktop/每月例行工作/工参/昌都245G工参总版-1028.xlsx', 'LTE').values.tolist()
base_station=base_station_5+base_station_4
content = read_txt('C:/Users/24253/Desktop/每月例行工作/123/模板/6.txt')
lst=content.split('\n')

for b in base_station_4:
    str=b[0].split('_')[1]+b[0].split('_')[2]
    if '芒康县如美镇政府' == str:
        print(b)

