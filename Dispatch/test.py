from docx import Document

doc = Document('C:/Users/24253/Desktop/JobReport/报告/网络测试-网格质量测试-昌都城区网格;LTE,NR;昌都;测试任务工单;2024年9月.docx')

for paragraph in doc.paragraphs:

    for run in paragraph.runs:
        if '<w:drawing' in run._r.xml:
            pic=run._r._add_drawing
            print(pic)
