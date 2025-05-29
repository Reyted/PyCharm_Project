import csv
import pandas as pd
import openpyxl
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def read_excel_file(file_path, sheet_name=0):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except FileNotFoundError:
        print(f"错误：文件 '{file_path}' 未找到")
    except Exception as e:
        print(f"读取文件时发生错误: {str(e)}")


def read_csv_with_csv_module(file_path,dic:dict,sheet1,sheet2):
    # sheet.append(['时间', '基站名称', '小区名称', '业务类型（短视频/直播）', '特定业务下行TCP环回总时延', '特定业务下行TCP包数', '特定业务下行TCP环回时延低于门限的包数','特定业务下行总包数'
    #               '特定业务上行总包数', '特定业务下行总数据量', '特定业务上行总数据量', '特定业务平均用户数', '特定业务上行秒级流量低于门限的时长', '特定业务上行持续时长', '特定业务上行秒级流量低于门限的时长',
    #               '特定业务上行持续时长'
    #               ])
    encoding = detect_encoding(file_path)
    print("开始解析短视频直播")
    try:
        with open(file_path, mode='r',encoding=encoding) as file:
            csv_reader = csv.reader(file)
            for row in csv_reader:
                try:
                    if row[3] == '0':
                        lst_temp=row[5:]
                        lst_temp.insert(0,row[0])
                        lst_temp.insert(1,row[1])
                        lst_temp.insert(2,dic[row[1]+row[2]])
                        lst_temp.insert(3,'短视频')
                        lst_temp.append(row[13])
                        lst_temp.append(row[14])
                        # print(lst_temp)
                        sheet1.append(lst_temp)
                except:
                    pass
                try:
                    if row[3] == '1':
                        lst_temp=row[5:]
                        lst_temp.insert(0,row[0])
                        lst_temp.insert(1,row[1])
                        lst_temp.insert(2,dic[row[1]+row[2]])
                        lst_temp.insert(3,'直播')
                        lst_temp.append(row[13])
                        lst_temp.append(row[14])
                        sheet2.append(lst_temp)
                except:
                    pass
    except:
        pass

def read_file2(file_path):
    encoding = detect_encoding(file_path)
    try:
        with open(file_path, mode='r',encoding=encoding) as file:
            csv_reader = csv.reader(file)
            for row in csv_reader:
                try:
                    HWL_DIC[row[0] + row[1] + row[2]] = row[7]
                except:
                    pass
    except:
        pass


def write_KPI(file_path,sheet3):
    print("开始解析KPI")
    encoding = detect_encoding(file_path)
    with open(file_path, mode='r', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            try:
                lst_temp = row[6:]
                lst_temp.insert(0, float(lst_temp[0])+float(lst_temp[1]))
                lst_temp.insert(0,row[0])
                lst_temp.insert(1,row[1])
                lst_temp.insert(2,row[4])
                lst_temp.insert(6,HWL_DIC[row[0]+row[1]+row[2]])
                sheet3.append(lst_temp)
            except:
                pass


def write_LL(file_path1,file_path2,sheet4):
    print("开始解析全网流量")
    encoding1 = detect_encoding(file_path1)
    encoding2 = detect_encoding(file_path2)
    with open(file_path1, mode='r', encoding=encoding1) as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            try:
                LL_DIC[row[0]] = float(row[3]) + float(row[4])
            except:
                pass

    with open(file_path2, mode='r', encoding=encoding2) as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            try:
                lst_temp = [row[0],float(row[3]) + float(row[4])+LL_DIC[row[0]]]
                lst_temp.insert(1,'苏州')
                sheet4.append(lst_temp)
            except:
                pass
# 使用示例
if __name__ == "__main__":
    cell_dic={}
    HWL_DIC={}
    LL_DIC={}
    # 替换为你的Excel文件路径
    excel_file = "C:/Users/24253/Desktop/新建文件夹 (9)/苏州智能板部署清单-20250328.xlsx"
    file_path1='C:/Users/24253/Desktop/新建文件夹 (9)/2288x/0328全量智能板指标提取z-2288_查询结果_(短视频直播).csv'
    file_path1_1='C:/Users/24253/Desktop/新建文件夹 (9)/2288x/0328全量智能板指标提取z-2288_查询结果_(子报表 1).csv'
    file_path1_2='C:/Users/24253/Desktop/新建文件夹 (9)/2288x/0328全量智能板指标提取z-2288_查询结果_(子报表 2).csv'
    file_path1_3='C:/Users/24253/Desktop/新建文件夹 (9)/0328全量智能板指标提取z-2288_查询结果_20250515150950938(子报表 1).csv'
    file_path2='C:/Users/24253/Desktop/新建文件夹 (9)/TS/0328全量智能板指标提取z-泰山_查询结果_(短视频直播).csv'
    file_path2_1='C:/Users/24253/Desktop/新建文件夹 (9)/TS/0328全量智能板指标提取z-泰山_查询结果_(子报表 1).csv'
    file_path2_2='C:/Users/24253/Desktop/新建文件夹 (9)/TS/0328全量智能板指标提取z-泰山_查询结果_(子报表 2).csv'
    file_path2_3='C:/Users/24253/Desktop/新建文件夹 (9)/0328全量智能板指标提取z-泰山_查询结果_20250515151000678(子报表 1).csv'
    file_path3 = 'C:/Users/24253/Desktop/新建文件夹 (9)/地市智能板区域指标收集-20250515-KPI+短视频+上行直播.xlsx'
    workbook = openpyxl.load_workbook(file_path3)
    sheet1 = workbook['短视频']
    sheet2 = workbook['直播']
    sheet3 = workbook['KPI']
    sheet4 = workbook['全网']
    # 读取Excel文件
    data = read_excel_file(excel_file)

    for li in data.values:
        cell_dic[li[5]]=li[6]

    read_csv_with_csv_module(file_path1,cell_dic,sheet1,sheet2)
    read_csv_with_csv_module(file_path2,cell_dic,sheet1,sheet2)

    read_file2(file_path2_2)
    read_file2(file_path1_2)
    write_KPI(file_path2_1,sheet3)
    write_KPI(file_path1_1,sheet3)

    write_LL(file_path2_3,file_path1_3,sheet4)

    workbook.save(file_path3)