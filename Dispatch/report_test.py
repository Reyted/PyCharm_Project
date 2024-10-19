import time
import xlwings as xw
from docxtpl import DocxTemplate

wb=xw.Book('C:/Users/24253/Desktop/每月例行工作/10月工作内容/10月工单报告/10月工单报告附件/昌都市城区LTE&NR测试分析优化问题点10月 - 副本.xlsx')
sht=wb.sheets['Sheet1']
data_list=sht.range(2,3).options(expand='table').value


tpl=DocxTemplate('C:/Users/24253/Desktop/每月例行工作/10月工作内容/10月派单回单/派单/NR;昌都市城关镇北大桥路段;NR弱覆盖;2024年10月_模板.docx')
for data in data_list:
    pos=data[3].split('道路')[0]
    type='弱覆盖' if data[8]=='覆盖问题' else '质差'
    datetime=data[0].strftime('%Y年%m月')
    context={
        'system':data[2],
        'position':pos,
        'question_type':type,
        'length':int(data[6]),
        'datetime':data_list[0][0].strftime('%Y年%m月'),
        'lontitude':data[4],
        'latitude':data[5],
        'serving_cell':data[7],
        'type_define':data[11],
        'problem_analysis':data[9],
        'solution_measures':data[10]
    }

    tpl.render(context)
    tpl.save(f'C:/Users/24253/Desktop/每月例行工作/10月工作内容/10月派单回单/派单/{data[2]};{pos};{data[2]}{type};{datetime}.docx')

wb.close()