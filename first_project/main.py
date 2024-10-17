import csv
import openpyxl
import time
import pandas as pd


#写数据
def write_csv_file(data, filename):
    # 创建一个新的Excel工作簿
    workbook = openpyxl.Workbook()

    #当前活动Sheet
    sheet1 = workbook.active
    sheet1.title = '基带板-全网'
    # 创建一个新的表单
    sheet2 = workbook.create_sheet("RRU-全网")

    #添加一行数据
    sheet2.append(['网元类型','单板名称','单板类型','生产日期','机柜号','机框号','槽位号','特殊信息','资产序列号','网元型号','小区名称'])
    sheet1.append(['网元类型','单板名称','单板类型','生产日期','机柜号','机框号','槽位号','特殊信息','资产序列号'])


    ind1=0
    ind2=0

    for row in data:
        bol=('RRU5' in row[5] or 'AAU5' in row[5] or 'RRU3' in row[5] or 'RRU7' in row[5] or 'AAU3' in row[5] or 'RRN3' in row[5])
        li=[row[0],row[1],row[2],row[3],row[6],row[4],row[8],row[5],row[7]]
        if row[1] in 'UBBPLBBP':
            sheet1.append([row[0],row[1],row[2],row[3],row[6],row[4],row[8],row[5],row[7]])
        if row[1] in 'MRRUAIRUMRFUMPMURRN':
            #sheet2.append([row[0], row[1], row[2], row[3], row[6], row[4], row[8], row[5], row[7], row[5]])

            if not bol:
                sheet2.append([row[0], row[1], row[2], row[3], row[6], row[4], row[8], row[5], row[7], ' '])
            else:
                li = row[5].split(',')
                if  len(li)==1 or not bol:
                    sheet2.append([row[0],row[1],row[2],row[3],row[6],row[4],row[8],row[5],row[7],' '])
                elif len(li)>1 and bol:
                    for i in li:
                        li2=i.split(' ')
                        if len(li2)==1:
                            if 'RRU5' in i or 'AAU5' in i or 'RRU3' in i or 'RRU7' in i or 'AAU3' in i or 'RRN3' in i:
                                sheet2.append([row[0],row[1],row[2],row[3],row[6],row[4],row[8],row[5],row[7],i.split('(')[0]])
                                break
                        else:
                            for j in li2:
                                if 'RRU5' in j or 'AAU5' in j or 'RRU3' in j or 'RRU7' in j or 'AAU3' in j or 'RRN3' in j:
                                    sheet2.append([row[0],row[1],row[2],row[3],row[6],row[4],row[8],row[5],row[7],j.split('(')[0]])
                                    break



    # 保存工作簿
    #获取当前时间
    today = time.strftime("%Y-%m-%d", time.localtime())
    workbook.save(filename+f'硬件信息{today}.xlsx')


#小区查询
def lst_cell(filepath1:str,filepath2:str):
    data = []
    with open(filepath1, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            identify1=row[1]+row[5]
            with open(filepath2, 'r') as f:
                reader = csv.reader(f)
                for row in reader:
                    indetify2=row[1]+row[4]
                    if indetify2 == identify1:
                        data.append(row[5])
                        print(row[5])
    print(data)


#合并函数
def merge_csv(file1, file2, output_file):
    li1 = []
    li2 = []
    #打开第一个CSV文件并读取内容
    with open(file1, 'r') as f1:
        reader1 = csv.reader(f1)
        index=[]
        ind=0
        data1 = list(reader1)
        for i in data1[0]:
            if i in '网元类型单板名称单板类型生产日期机柜号机框号槽位号特殊信息资产序列号':
                index.append(ind)
            ind+=1
        for row in data1:
            li=[]
            for i in index:
                li.append(row[i])
            li1.append(li)

    #打开第二个CSV文件并读取内容
    with open(file2, 'r') as f2:
        reader2 = csv.reader(f2)
        index = []
        ind = 0
        data2 = list(reader2)
        for i in data2[0]:
            if i in '网元类型单板名称单板类型生产日期机柜号机框号槽位号特殊信息资产序列号':
                index.append(ind)
            ind+=1
        for row in data2:
            li = []
            for i in index:
                li.append(row[i])
            li2.append(li)

    # 合并两个CSV文件的内容
    merged_data = li1+li2

    #写入新的CSV文件
    with open(output_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerows(merged_data)

#筛选函数
def read_csv_file(filepath):
    with open(filepath, 'r') as f:
        reader = csv.reader(f)
        data = []
        for row in reader:
            data.append(row)
        write_csv_file(data, 'C:/Users/24253/Desktop/PLAY/')


if __name__ == '__main__':
    lst_cell('/Users/reyted/Desktop/PLAY/查询小区物理单板拓扑关系.csv','/Users/reyted/Desktop/PLAY/查询小区静态参数.csv')

    # 调用函数进行合并
    #merge_csv('C:/Users/24253/Desktop/PLAY/存量_板_20241009_101654.csv', 'C:/Users/24253/Desktop/PLAY/存量_板_20241009_101706.csv', 'C:/Users/24253/Desktop/PLAY/merged_file.csv')

    #read_csv_file('C:/Users/24253/Desktop/PLAY/merged_file.csv')