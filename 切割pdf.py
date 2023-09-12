import re
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter


def split_single_pdf(read_file, start_page, end_page, pdf_file):
    # 1. 获取原始pdf文件
    fp_read_file = open(read_file, 'rb')
    # 2. 将要分割的PDF内容格式化
    pdf_input = PdfFileReader(fp_read_file)
    # 3. 实例一个 PDF文件编写器
    pdf_output = PdfFileWriter()
    # 4. 把3到4页放到PDF文件编写器
    for i in range(start_page, end_page):
        pdf_output.addPage(pdf_input.getPage(i))
    # 5. PDF文件输出
    with open(pdf_file, 'wb') as pdf_out:
        pdf_output.write(pdf_out)
    print(f'{read_file}分割{start_page}页-{end_page}页完成，保存为{pdf_file}!')

# 获取 PDF 信息
pdfFile = open('数媒动漫短片获奖证书.pdf', 'rb')
in_pdf_name = "数媒动漫短片获奖证书.pdf"
pdfObj = PyPDF2.PdfFileReader(pdfFile)
page_count = pdfObj.getNumPages()
x=[]
#提取文本
for p in range(0, page_count):
    text = pdfObj.getPage(p)
    text = text.extractText()
    rule = r'证书编号：JSJDS20230054(\d{10})'  # 正则规则
    if(re.findall(rule, text)):
        x.append(re.findall(rule, text))
        #保存和发送
        # 切分后文件文件名
        tmp=str(re.findall(rule, text))
        tmp=tmp[2:-2]
        out_pdf_name = tmp+'.pdf'
        # 切分开始页面
        start = p
        # 切分结束页面
        end = p+1
        split_single_pdf(in_pdf_name, start, end, out_pdf_name)

print(x)

