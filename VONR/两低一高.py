import os
import pandas as pd
import csv

# 5G-4G低切换

date_lst_xw={}
date_lst_ts={}

rrc_successful={}

ind_5G_4G = 0
ind_djt=0
one_ind_djt=0
ind_gdh=0
one_ind_gdh=0
ind_EpsFallBack = 0
cell_num=0
one_ind_5G_4G=0
one_ind_EpsFallBack=0
one_cell_num=0

vinr_djt=0
vinr_gdh=0
one_vinr_djt=0
one_vinr_gdh=0


max_lst=[]


def read_ali(file_path,lst_all:list):
    pf=pd.read_excel(file_path,sheet_name='差小区')
    for li in pf.values:
        str1=str(li[0]).split(' ')
        if str1[0][0]=='2' or str1[0][0]=='2025-02-17':
            try:
                all_cellnum=li[1]+date_lst_xw[str1[0]]['小区数']+date_lst_ts[str1[0]]['小区数']
                djt=(li[2]+date_lst_xw[str1[0]]['vonr低接通']+date_lst_ts[str1[0]]['vonr低接通'])/all_cellnum
                gdh=(li[3]+date_lst_xw[str1[0]]['vonr高掉话']+date_lst_ts[str1[0]]['vonr高掉话'])/all_cellnum
                ffdqh=(li[4]+date_lst_xw[str1[0]]['5G-4G低切换']+date_lst_ts[str1[0]]['5G-4G低切换'])/all_cellnum
                vinrdjt=(li[5]+date_lst_xw[str1[0]]['vinr低接通']+date_lst_ts[str1[0]]['vinr低接通'])/all_cellnum
                vinrgdh=(li[6]+date_lst_xw[str1[0]]['vinr高掉话']+date_lst_ts[str1[0]]['vinr高掉话'])/all_cellnum
                efdqh=(li[7]+date_lst_xw[str1[0]]['EpsFallBack低切换']+date_lst_ts[str1[0]]['EpsFallBack低切换'])/all_cellnum
                lst_all.append([str1[0],all_cellnum,djt,gdh,ffdqh,djt+gdh+ffdqh,vinrdjt,vinrgdh,efdqh])
            except:
                pass


def read_sheet3(file_path,lst_all:list):
    global vinr_djt
    global vinr_gdh
    global one_vinr_djt
    global one_vinr_gdh
    with open(file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        date_time = '2025-02-17'
        for row in reader:
            try:
                lst = row[0].split('-')
                if row[0] != '日期' and len(lst) == 3 and row[0][0] == '2':
                    if row[0] == date_time:
                        try:
                            if float(row[8]) > 1:
                                one_vinr_gdh += 1
                                lst_all['2025-02-17'].update({'vinr高掉话': one_vinr_gdh})
                        except:
                            pass
                        try:
                            str2 = lst[0] + lst[1] + lst[2] + row[4]
                            if rrc_successful[str1] * float(row[7]) < 95:
                                one_vinr_djt += 1
                                lst_all['2025-02-17'].update({'vinr低接通': one_vinr_djt})
                            else:
                                pass
                        except:
                            pass
                    if date_time == row[0]:
                        try:
                            if len(lst[1]) == 2:
                                str3 = lst[0] + '-' + lst[1] + '-' + lst[2]
                            else:
                                str3 = lst[0] + '-' + '0' + lst[1] + '-' + lst[2]
                            try:
                                if float(row[8]) > 1:
                                    vinr_gdh += 1
                                    lst_all[str3].update({'vinr高掉话': vinr_gdh})
                            except:
                                lst_all[str3].update({'vinr高掉话': vinr_gdh})

                            if len(lst[1]) == 2:
                                str2 = lst[0] + lst[1] + lst[2] + row[4]
                            else:
                                str2 = lst[0] + '0' + lst[1] + lst[2] + row[4]
                            try:
                                if rrc_successful[str2] * float(row[7]) < 95:
                                    vinr_djt += 1
                                    if len(lst[1]) == 2:
                                        str3 = lst[0] + '-' + lst[1] + '-' + lst[2]
                                    else:
                                        str3 = lst[0] + '-' + '0' + lst[1] + '-' + lst[2]
                                    lst_all[str3].update({'vinr低接通': vinr_djt})
                            except:
                                lst_all[str3].update({'vinr低接通': vinr_djt})
                        except:
                            pass
                    else:
                        date_time = row[0]
                        vinr_djt = 0
                        vinr_gdh = 0
                        continue
            except:
                pass

def read_sheet2(file_path,lst_all:list):
    global ind_djt
    global ind_gdh
    global one_ind_gdh
    global one_ind_djt
    with open(file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        date_time = '2025-02-17'
        for row in reader:
            try:
                lst=row[0].split('-')
                if row[0]!='日期' and len(lst) == 3 and row[0][0] == '2':
                    if row[0]==date_time:
                        try:
                            if float(row[8])>1:
                                one_ind_gdh+=1
                                lst_all['2025-02-17'].update({'vonr高掉话': one_ind_gdh})
                                # lst_all['2025-02-17'] = {'小区数': lst_all['2025-02-17']['小区数'],
                                #                          '5G-4G低切换': lst_all['2025-02-17']['5G-4G低切换'],
                                #                          'EpsFallBack低切换': lst_all['2025-02-17'][
                                #                              'EpsFallBack低切换'],
                                #                          'vonr高掉话': one_ind_gdh}
                        except:
                            lst_all['2025-02-17'].update({'vonr高掉话': one_ind_gdh})
                        try:
                            str2 = lst[0] + lst[1] + lst[2] + row[4]
                            if rrc_successful[str1] * float(row[7]) < 95:
                                one_ind_djt += 1
                                lst_all['2025-02-17'].update({'vonr低接通': one_ind_djt})
                                # lst_all['2025-02-17'] = {'小区数': lst_all['2025-02-17']['小区数'],
                                #                          '5G-4G低切换': lst_all['2025-02-17']['5G-4G低切换'],
                                #                          'EpsFallBack低切换': lst_all['2025-02-17'][
                                #                              'EpsFallBack低切换'],
                                #                          'vonr低接通': one_ind_djt}
                            else:
                                pass
                        except:
                            lst_all['2025-02-17'].update({'vonr低接通': one_ind_djt})
                    if date_time == row[0]:
                        try:
                            if len(lst[1]) == 2:
                                str3 = lst[0] + '-' + lst[1] + '-' + lst[2]
                            else:
                                str3 = lst[0] + '-' + '0' + lst[1] + '-' + lst[2]
                            try:
                                if float(row[8]) > 1:
                                    ind_gdh += 1
                                    lst_all[str3].update({'vonr高掉话': ind_gdh})
                            except:
                                lst_all[str3].update({'vonr高掉话': ind_gdh})

                            if len(lst[1]) == 2:
                                str2 = lst[0] + lst[1] + lst[2] + row[4]
                            else:
                                str2 = lst[0] + '0' + lst[1] + lst[2] + row[4]
                            try:
                                if len(lst[1]) == 2:
                                    str3 = lst[0] + '-' + lst[1] + '-' + lst[2]
                                else:
                                    str3 = lst[0] + '-' + '0' + lst[1] + '-' + lst[2]
                                if rrc_successful[str2] * float(row[7]) < 95:
                                    ind_djt += 1
                                    lst_all[str3].update({'vonr低接通': ind_djt})
                            except:
                                lst_all[str3].update({'vonr低接通': ind_djt})
                        except:
                            pass
                    else:
                        date_time = row[0]
                        ind_djt=0
                        ind_gdh=0
                        continue
            except:
                pass



def read_sheet1(file_path,lst_all:list):
    global ind_5G_4G
    global ind_EpsFallBack
    global cell_num
    global one_ind_5G_4G
    global one_ind_EpsFallBack
    global one_cell_num

    with open(file_path, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        date_time = '2025-02-17'
        for row in reader:
            try:
                if row[0]!='日期' and len(row[0].split('-')) == 3 and row[0][0]=='2':
                    if row[0]=='2025-02-17':
                        one_cell_num += 1
                        lst = date_time.split('-')
                        str2 = lst[0] + lst[1] + lst[2]
                        rrc_successful[str2+row[4]]=float(row[6])
                        try:
                            if float(row[7]) < 95:
                                one_ind_5G_4G += 1
                            if float(row[8]) < 0.95:
                                one_ind_EpsFallBack += 1
                        except:
                            pass
                        lst_all['2025-02-17'] = {'小区数': one_cell_num, '5G-4G低切换': one_ind_5G_4G,
                                                  'EpsFallBack低切换': one_ind_EpsFallBack}
                    if date_time == row[0]:
                        cell_num+=1
                        lst=date_time.split('-')
                        str2=lst[0]+lst[1]+lst[2]
                        rrc_successful[str2 + row[4]] = float(row[6])
                        try:
                            if float(row[7]) < 95:
                                ind_5G_4G += 1
                        except:
                            pass

                        try:
                            if float(row[8]) < 0.95:
                                ind_EpsFallBack += 1
                        except:
                            pass
                        lst4=date_time.split('-')
                        if len(lst4[1]) == 2:
                            str4 = lst4[0] + '-' + lst4[1] + '-' + lst4[2]
                        else:
                            str4 = lst4[0] + '-' + '0' + lst4[1] + '-' + lst4[2]
                        lst_all[str4] = {'小区数': cell_num, '5G-4G低切换': ind_5G_4G,
                                          'EpsFallBack低切换': ind_EpsFallBack}
                    else:
                        date_time = row[0]
                        cell_num=0
                        ind_5G_4G = 0
                        ind_EpsFallBack = 0
                        continue
            except:
                pass


if __name__ == '__main__':
    file_path_xw = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（2288X） (1)/VONR两低一高小区级（2288X）(NR小区（2288X）).csv'
    file_path_xw_5QI1 = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（2288X） (1)/VONR两低一高小区级（2288X）(NR小区5QI1（2288X）).csv'
    file_path_xw_5QI2 = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（2288X） (1)/VONR两低一高小区级（2288X）(NR小区5QI2（2288X）).csv'
    file_path_xwww = 'C:/Users/24253/Desktop/新建文件夹 (6)/工作簿2.csv'
    file_path_xwww_5QI1 = 'C:/Users/24253/Desktop/新建文件夹 (6)/工作簿3.csv'
    file_path_ts = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（泰山）/VONR两低一高小区级（泰山）(NR小区（泰山）).csv'
    file_path_ts_5QI1 = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（泰山）/VONR两低一高小区级（泰山）(NR小区5QI1（泰山）).csv'
    file_path_ts_5QI2 = 'C:/Users/24253/Desktop/新建文件夹 (2)/VONR两低一高小区级（泰山）/VONR两低一高小区级（泰山）(NR小区5QI2（泰山）).csv'

    file_path_alx = 'C:/Users/24253/Desktop/新建文件夹 (6)/指标.xlsx'

    read_sheet1(file_path_xw,date_lst_xw)
    read_sheet1(file_path_ts,date_lst_ts)
    # read_sheet1(file_path_xwww,date_lst_xw)

    # print(len(rrc_successful))
    read_sheet2(file_path_xw_5QI1,date_lst_xw)
    read_sheet2(file_path_ts_5QI1,date_lst_ts)

    read_sheet3(file_path_xw_5QI2, date_lst_xw)
    read_sheet3(file_path_ts_5QI2, date_lst_ts)

    # read_sheet2(file_path_xwww_5QI1,date_lst_xw)
    print(date_lst_xw)
    print(date_lst_ts)

    read_ali(file_path_alx,max_lst)
    print(max_lst)

    data = max_lst

    # 转换为DataFrame（不指定列名）
    df = pd.DataFrame(data)

    # 写入Excel文件（默认写入Sheet1，不包含索引）
    df.to_excel("C:/Users/24253/Desktop/新建文件夹 (6)/工作簿1.xlsx", index=False)

    print("数据已成功写入")