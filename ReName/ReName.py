import os
import random
import shutil
import openpyxl

townlist=['卡若区','八宿县','边坝县','察雅县','丁青县','贡觉县','江达县','类乌齐县','洛隆县','芒康县','左贡县']
#townlist1=['卡若区','八宿县','边坝县','察雅县','丁青县','贡觉县','江达县','类乌齐县','洛隆县','芒康县','左贡县']


url='C:/Users/24253/Desktop/每月例行工作/10月工作内容/10月录音/'


def createxl():
    wb = openpyxl.load_workbook(f'{url}test.xlsx')
    sheet = wb['Sheet2']
    rw = 0
    for row in sheet.iter_rows():
        cm = 0
        list: str = []
        for cell in row:
            rw += 1
            cm += 1
            list.append(cell.value)
        if rw != 3:
            if list[2]:
                county = list[2].split("_")[1]
                route = f'{url}phonenum/' + county + '.txt'
                with open(route, 'a', encoding='utf-8') as f:
                    f.write(str(list[1]) +'_'+list[0]+ '\n')


def rename(fn:list,tw:str):
    old_mikdir = f"{url}素材/"+tw+"素材"
    new_mikdir = f"{url}录音/"+tw+"录音"
    list=[]

    if not os.path.exists(new_mikdir):
        os.mkdir(new_mikdir)

    for filename in os.listdir(old_mikdir):
        old_filename = old_mikdir + "/" + filename
        list.append(old_filename)

    index = random.randint(0, len(list) - 1)
    new_filename = new_mikdir + "/" + fn[0] + "首响+用户表示已恢复.m4a"

    while os.path.exists(new_filename):
        new_filename = new_mikdir + "/" + fn[0] + "首响"+"+"+"用户表示已恢复"+"+"+fn[1]+".m4a"
    shutil.copy(list[index], new_filename)


createxl()


for town in townlist:

    with open(f'{url}phonenum/'+town+'.txt', 'r', encoding='utf-8') as f:
        for line in f.readlines():
            fnstrlst = line.strip('\n').split('_')
            rename(fnstrlst,town)

if __name__ == '__main__':
    pass