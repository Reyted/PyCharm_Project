import fitz  # PyMuPDF

import fitz  # PyMuPDF


def read_pdf_to_console(file_path):
    try:
        # 打开PDF文件
        pdf_document = fitz.open(file_path)
        print(f"Number of pages: {len(pdf_document)}")

        # 遍历每一页并读取其内容，然后输出到控制台
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)  # 根据页码加载页
            page_text = page.get_text("text")  # 提取文本
            print(f"Page {page_num + 1}:\n{page_text}\n{'-' * 40}\n")  # 输出文本，并添加分隔线

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # 确保文件在读取后被正确关闭
        pdf_document.close()


# 调用函数并传递PDF文件路径
pdf_file_path = "C:/Users/24253/Desktop/Python/Python_Pdf/111.pdf"
read_pdf_to_console(pdf_file_path)
