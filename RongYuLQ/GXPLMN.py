import csv

eNodeBFunction_dic=dict()
cellStaticParm_lst=[]
cellOperator_lst=[]

def read_cellStaticParm(file_path,lst:[]):
    with open(file_path, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        for row in reader:
            row.append(eNodeBFunction_dic[row[1]])
            lst.append(eNodeBFunction_dic[row[1]])

def read_eNodeBFunction(file_path,dic:dict()):
    with open(file_path, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter=',')
        for row in reader:
            dic[row[1]]=row[6]


if __name__ == '__main__':
    eNodeBFunction_path='G:/工作内容/25年每月例行工作/1月工作内容/冗余邻区/2月/解析文件/查询gNodeB功能.csv'
    cellStaticParm_path='G:/工作内容/25年每月例行工作/1月工作内容/冗余邻区/2月/解析文件/查询小区静态参数.csv'
    cellOperator_path='G:/工作内容/25年每月例行工作/1月工作内容/冗余邻区/2月/解析文件/查询小区运营商信息.csv'
    read_eNodeBFunction(eNodeBFunction_path,eNodeBFunction_dic)
    read_cellStaticParm(cellStaticParm_path,cellStaticParm_lst)


    #print(cellStaticParm_lst)