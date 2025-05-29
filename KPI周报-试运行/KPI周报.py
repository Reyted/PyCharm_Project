import openpyxl


def read_excel_sheet(file_path, sheet_name):
    try:
        # 加载工作簿
        workbook = openpyxl.load_workbook(file_path)

        # 获取指定的 sheet
        sheet = workbook[sheet_name]

        # 读取所有数据
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(list(row))

        return data

    except Exception as e:
        print(f"读取 Excel 文件时出错: {e}")
        return None


# 使用示例
if __name__ == "__main__":
    file_path = "C:/Users/24253/Desktop/新建文件夹/维护宏站KPI（2288X）_查询结果_20250527174528990.xlsx"
    file_path2 = "C:/Users/24253/Desktop/新建文件夹/站点清单.xlsx"
    sheet_name1 = "NR小区-EPSFB及5-4语音指标"
    sheet_name2 = "NR小区-5QI1指标"
    sheet_name3 = "NR DU小区-5QI1指标"
    sheet_name4 = "Sheet1"
    ZDQD_DIC={}
    END_DADALST=[]
    sheet2_data={}
    sheet3_data={}
    all_data_dic={}

    sheet_data4 = read_excel_sheet(file_path2, sheet_name4)
    sheet_data1 = read_excel_sheet(file_path, sheet_name1)
    sheet_data2 = read_excel_sheet(file_path, sheet_name2)
    sheet_data3 = read_excel_sheet(file_path, sheet_name3)

    if sheet_data4:
        for row in sheet_data4:
            ZDQD_DIC[row[1]] = row[11]

    if sheet_data2:
        for row in sheet_data2:
            try:
                sheet2_data[f'{row[0]}${ZDQD_DIC[row[4]]}'] = row[7:]
            except:
                pass

    if sheet_data3:
        for row in sheet_data3:
            try:
                sheet3_data[f'{row[0]}${ZDQD_DIC[row[4]]}'] = [row[10]]
            except:
                pass
        #print(sheet3_data)

    if sheet_data1:
        lst=[]
        for row in sheet_data1:
            try:
                index=f'{row[0]}${ZDQD_DIC[row[4]]}'
                lst = row[6:14]
                lst.append(sheet2_data[index][2])
                lst.append(sheet3_data[index][0])
                lst.append(row[14])
                lst.append(sheet2_data[index][1])
                lst.append(row[15])
                lst.append(row[18]*sheet2_data[index][0])
                lst.append(row[18]*sheet2_data[index][0])
                lst.append(row[16])
                lst.append(row[17])
                if index in all_data_dic:
                    all_data_dic[index]=[a + b for a, b in zip(all_data_dic[index], lst)]
                else:
                    all_data_dic[index] = lst
            except:
                pass
        for i,j in all_data_dic.items():
            all_data_dic[i]=[]
            if j[0]>200 and j[0]<300:
                for li in j:
                    try:
                        all_data_dic[i].append(li / 3)
                    except:
                        all_data_dic[i].append(0)
                all_data_dic[i].pop()
                all_data_dic[i].append(j[len(j) - 1])
            if j[0]>100 and j[0]<200:
                for li in j:
                    try:
                        all_data_dic[i].append(li / 2)
                    except:
                        all_data_dic[i].append(0)
                all_data_dic[i].pop()
                all_data_dic[i].append(j[len(j) - 1])
            if j[0]>300 and j[0]<400:
                for li in j:
                    try:
                        all_data_dic[i].append(li / 4)
                    except:
                        all_data_dic[i].append(0)
                all_data_dic[i].pop()
                all_data_dic[i].append(j[len(j)-1])
            if j[0]<100:
                all_data_dic[i]=j

        print(all_data_dic)