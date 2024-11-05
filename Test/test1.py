import pandas as pd
import xlwings as xw


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

if __name__ == '__main__':
    content = read_txt('C:/Users/24253/Desktop/每月例行工作/123/模板/5.txt')
    base_station_5 = read_excel('C:/Users/24253/Desktop/每月例行工作/工参/昌都245G工参总版-1028.xlsx',
                                'NR').values.tolist()
    base_station_4 = read_excel('C:/Users/24253/Desktop/每月例行工作/工参/昌都245G工参总版-1028.xlsx',
                                'LTE').values.tolist()
    base_station = base_station_5 + base_station_4
    lst = content.split('\n')
    ind=0
    for i in lst:
        for item in base_station:
            if i in item:
                ind+=1
                break
            else:
                print(i)
                continue