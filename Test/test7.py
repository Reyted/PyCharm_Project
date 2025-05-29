import pandas as pd
import os


def excel_to_vcf(excel_path, vcf_path, sheet_name=0):
    """
    将Excel文件转换为VCF文件

    参数:
        excel_path (str): Excel文件路径
        vcf_path (str): 输出的VCF文件路径
        sheet_name (str/int): Excel工作表名称或索引，默认为第一个工作表
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        # 检查必要的列是否存在
        required_columns = ['姓名']  # 至少需要姓名列
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Excel文件中缺少必要列: {col}")

        # 创建VCF文件
        with open(vcf_path, 'w', encoding='utf-8') as vcf_file:
            for index, row in df.iterrows():
                # 开始一个vCard
                vcf_file.write('BEGIN:VCARD\n')
                vcf_file.write('VERSION:3.0\n')

                # 姓名 (必须)
                vcf_file.write(f'N:{row["姓名"]}\n')
                vcf_file.write(f'FN:{row["姓名"]}\n')

                # 电话 (可选)
                if '电话' in df.columns and pd.notna(row['电话']):
                    vcf_file.write(f'TEL;TYPE=CELL:{row["电话"]}\n')

                # 电子邮件 (可选)
                if '邮箱' in df.columns and pd.notna(row['邮箱']):
                    vcf_file.write(f'EMAIL;TYPE=INTERNET:{row["邮箱"]}\n')

                # 公司/组织 (可选)
                if '公司' in df.columns and pd.notna(row['公司']):
                    vcf_file.write(f'ORG:{row["公司"]}\n')

                # 职位 (可选)
                if '职位' in df.columns and pd.notna(row['职位']):
                    vcf_file.write(f'TITLE:{row["职位"]}\n')

                # 地址 (可选)
                if '地址' in df.columns and pd.notna(row['地址']):
                    vcf_file.write(f'ADR;TYPE=WORK:{row["地址"]}\n')

                # 备注 (可选)
                if '备注' in df.columns and pd.notna(row['备注']):
                    vcf_file.write(f'NOTE:{row["备注"]}\n')

                # 结束vCard
                vcf_file.write('END:VCARD\n\n')

        print(f"成功转换 {len(df)} 个联系人到 {vcf_path}")

    except Exception as e:
        print(f"转换过程中出错: {str(e)}")
        # 如果出错，删除可能已创建的不完整VCF文件
        if os.path.exists(vcf_path):
            os.remove(vcf_path)


if __name__ == "__main__":
    # 示例用法
    excel_file = 'C:/Users/24253/Desktop/新建文件夹/工作簿1.xlsx'  # 输入的Excel文件
    vcf_file = 'C:/Users/24253/Desktop/新建文件夹/工作簿1.vcf'  # 输出的VCF文件

    excel_to_vcf(excel_file, vcf_file)

    print("转换完成!")