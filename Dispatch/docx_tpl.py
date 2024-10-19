import xlwings as xw
from docxtpl import DocxTemplate

wb=xw.Book('C:/Users/24253/Desktop/Python_Excel/t1.xlsx')
sht=wb.sheets['Sheet1']
data_list=sht.range('A2').options(expand='table').value

print(data_list)

tpl=DocxTemplate('C:/Users/24253/Desktop/Python_World/准考证.docx')

for data in data_list:
    context={
        'name':data[1],
        'sno':data[0],
        'gender':data[2],
        'city':data[3],
    }
    tpl.render(context)
    tpl.save(f'C:/Users/24253/Desktop/Python_World/学生准考证/{data[1]}.docx')
wb.close()