import os
import pdfplumber
import re
import string
import jieba.analyse
import xlsxwriter
from tqdm.autonotebook import tqdm
from fenci import processing
from fenci import seg_depart

file_list = []  # 新建一个空列表用于存放文件路径
file_name = []  # 新建一个空列表用于存放文件名字
file_dir = r'D:/Desktop/双碳政策知识图谱构建/政策原文-20221121'  # 遍历的文件夹路径
for files in os.walk(file_dir):  # 遍历指定文件夹及其下的所有子文件夹
    for file in files[2]:  # 遍历每个文件夹里的所有文件，（files[2]:母文件夹和子文件夹下的所有文件信息，files[1]:子文件夹信息，files[0]:母文件夹信息）
        if os.path.splitext(file)[1] == '.PDF' or os.path.splitext(file)[1] == '.pdf':  # 检查文件后缀名,逻辑判断用==
            filename = re.sub(".pdf", "", file)
            file_name.append(filename)  # 筛选后的文件名为字符串，将得到的文件名放进去列表，方便以后调用
            file_list.append(file_dir + '\\' + file)  # 给文件名加入文件夹路径

f = xlsxwriter.Workbook('D:/Desktop/关联条文-202221215-3.xlsx')
# sheet1 = f.add_worksheet('关联条文1')
sheet2= f.add_worksheet('关联条文2')
# row1=0
# col1=0
row2=0
col2=0
for i in tqdm(range(len(file_list)),colour= 'blue',desc = "已经完成进度",ncols = 100):   # 遍历所提取的文件地址
    pdf = pdfplumber.open(file_list[i])  # 循环执行打开pdf文件命令
    pages = pdf.pages  # pages属性获取页数
    text_all = []  # 新建一个列表，存放PDF文件文本解析
    all_text = ''
    for page in pages:  # 遍历pages里面每一页的信息
        text = page.extract_text()  # 提取当前页内容，赋值给text
        text = re.sub("\n", "", text)  # 去除换行
        text = re.sub("\s", "", text)  # 去除空格
        text_all.append(text)
    k = 0
    while k < len(text_all):
        all_text = all_text + str(text_all[k])
        k += 1
    pdf.close()
    # Related_provisions_1= re.findall(r'《[\u4e00-\u9fa5]+》', all_text)
    # Related_provisions_1 = list(set(Related_provisions_1))
    # for Related_provision_1 in Related_provisions_1:
    #     if (len(Related_provision_1) > 5):
    #         sheet1.write(row1, col1, file_name[i])
    #         col1 += 1
    #         sheet1.write(row1, col1, "关联条文")
    #         col1 += 1
    #         sheet1.write(row1, col1, Related_provision_1)
    #         row1 += 1
    #         col1 = 0
    print(all_text)
    Related_provisions_2= re.findall(r'《(.*?)}', all_text)
    Related_provisions_2 = list(set(Related_provisions_2))
    for Related_provision_2 in Related_provisions_2:
        if (len(Related_provision_2) > 5):
            name_list=file_name[i].split('、', 1)
            if Related_provision_2!=name_list[1]:
                sheet2.write(row2, col2, file_name[i])
                col2 += 1
                sheet2.write(row2, col2, "关联条文")
                col2 += 1
                sheet2.write(row2, col2, Related_provision_2)
                row2 += 1
                col2 = 0
f.close()