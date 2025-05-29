import pandas as pd

CELL_CJ={}
CJ_CELL={'苏州站':[],'寒山寺':[],'虎丘山风景区':[],'同里古镇':[],'拙政园':[],'汽车南站':[],'金鸡湖景区':[],'木渎古镇':[],'周庄古镇':[],'太湖旅游区':[],'苏州北站':[]}

# 读取Excel文件（自动检测sheet）
file_path1 = "C:/Users/24253/Desktop/新建文件夹/流量/临时查询_查询结果_LTE_LL_2288x.xlsx"  # 替换为你的文件路径
file_path2 = "C:/Users/24253/Desktop/新建文件夹/流量/临时查询_查询结果_LTE_LL_TS.xlsx"  # 替换为你的文件路径

file_path4 = "C:/Users/24253/Desktop/新建文件夹/用户/临时查询_查询结果_LTE_2288x.xlsx"  # 替换为你的文件路径
file_path5 = "C:/Users/24253/Desktop/新建文件夹/用户/临时查询_查询结果_LTE_yh_ts.xlsx"  # 替换为你的文件路径

file_path3 = "C:/Users/24253/Desktop/新建文件夹/场景清单汇总0429-苏州.xlsx"  # 替换为你的文件路径
LTE_LL_XW = pd.read_excel(file_path1, engine='openpyxl')
LTE_LL_TS = pd.read_excel(file_path2, engine='openpyxl')
CJ=pd.read_excel(file_path3, engine='openpyxl',sheet_name='4G小区清单')

for li in CJ.values:
    CELL_CJ[li[2]]=li[3]

for li in LTE_LL_XW.values:
    CJ_CELL[CELL_CJ[li[4]]].append(li[0]+"*"+str(li[7]))
for li in LTE_LL_TS.values:
    CJ_CELL[CELL_CJ[li[4]]].append(li[0]+"*"+str(li[7]))
print(CJ_CELL['苏州北站'])

