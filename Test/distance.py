import math
import pandas as pd
import openpyxl


def calculate_distance(lat1, lon1, lat2, lon2):
    """
    计算两点之间的大圆距离（单位：米）
    使用Haversine公式

    参数:
    lat1, lon1: 源点的纬度和经度（度）
    lat2, lon2: 目标点的纬度和经度（度）

    返回:
    两点之间的距离（米）
    """
    # 将角度转换为弧度
    lat1_rad = math.radians(lat1)
    lon1_rad = math.radians(lon1)
    lat2_rad = math.radians(lat2)
    lon2_rad = math.radians(lon2)

    # 计算差值
    delta_lat = lat1_rad - lat2_rad
    delta_lon = lon1_rad - lon2_rad

    # Haversine公式
    a = (math.sin(delta_lat / 2) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * (math.sin(delta_lon / 2) ** 2))
    distance = 2 * math.asin(math.sqrt(a)) *6378137  # 地球半径6378137米

    return distance


if __name__ == '__main__':
    lst=[1,2,3,4,5,6,7,8,9]

    #for x in lst:
    file_path = 'F:/邻区核查/0522_NR/单板不受限小区0522.csv'
    file_path2 = 'F:/邻区核查/0522_NR/NR邻区精简0522.xlsx'
    df = pd.read_csv(file_path, encoding='gbk')
    lst = []
    lstt = []
    workbook = openpyxl.load_workbook(file_path2)
    sheet = workbook['Sheet1']
    sheet.append(['基站名称', '本地小区标识', '基站名称+本地小区标识', '日期', '基站名称', '本地小区名称', '源经度',
                  '源纬度', '目标基站标识', '目标小区标识', '目经度', '目纬度', '移动国家码', '移动网络码',
                  '目标小区名称', '完整度', '特定两小区间切换出尝试次数', '特定两小区间切换出成功次数', '小区名称',
                  'CMCC-切换成功率-分母', '特定两小区间切换出尝试次数_Sum', '切换尝试次数占比核查', '差值', '距离'
                  ])
    for li in df.values:
        print(li)
        # LTE
        # distance = calculate_distance(li[7]
        # , li[6], li[11]
        # , li[10]
        # )
        # if distance>1000 and li[16]<50:
        #     lstt=list(li)
        #     lstt.append(distance)
        #     sheet.append(lstt)

        # NR
        # distance = calculate_distance(li[5]
        #                               , li[4], li[14]
        #                               , li[13]
        #                               )
        # if distance > 1000 and li[16] < 30:
        #     lstt = list(li)
        #     lstt.append(distance)
        #     sheet.append(lstt)

    workbook.save(file_path2)
