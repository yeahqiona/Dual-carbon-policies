import jieba
import re
import jieba.analyse
import csv
import os  # 引用os库
import pdfplumber  # 引进pdfplumber库
import xlrd
import xlwt
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

# 批量解析每一个pdf文件
# f = xlwt.Workbook()
# sheet = f.add_sheet('sheet1',cell_overwrite_ok=True)
# row=0
for i in tqdm(range(len(file_list)),colour= 'blue',desc = "已经完成进度",ncols = 100): # 遍历所提取的文件地址
    col=0
    pdf = pdfplumber.open(file_list[i])  # 循环执行打开pdf文件命令
    pages = pdf.pages  # pages属性获取页数
    text_all = []  # 新建一个列表，存放PDF文件文本解析
    keyword_all = []
    all_text = ''
    #print('-----------------------')  # 用来间隔每个文件之间的内容，好看
    for page in pages:  # 遍历pages里面每一页的信息
        text = page.extract_text()  # 提取当前页内容，赋值给text
        text = re.sub("\n", "", text)  # 去除换行
        text = re.sub("\s", "", text)  # 去除空格
        text = re.sub(r'[0-9]+', '', text)
        text = re.sub(r'[a-z]+', '', text)
        text = re.sub(r'[A-Z]+', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}条', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}篇', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}节', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}章', '', text)
        text = re.sub(r'[\u4e00-\u9fa5]{1,2}省', '', text)
        text = re.sub(r'[\u4e00-\u9fa5]{1,2}市', '', text)
        text = re.sub(r'[\u4e00-\u9fa5]{1,2}区', '', text)
        text = re.sub(r'[\u4e00-\u9fa5]{1,2}县', '', text)
        text = re.sub(r'—[0-9]{1,2}—', '', text)
        text = re.sub(r'-[0-9]{1,2}—', '', text)
        text = re.sub(r'（.+?）', '', text)
        text = re.sub(r'【.+?】', '', text)
        text_all.append(text)
    k = 0
    while k < len(text_all):
        all_text = all_text + str(text_all[k])
        k += 1

    text=processing(all_text)
    text_seg = seg_depart(text)


    jieba.analyse.set_idf_path(r'D:/Desktop/IDF2.txt')
    jieba.analyse.set_stop_words('D:/Desktop/双碳政策知识图谱构建/文件读取/Chinese_stop_words.txt')#停用词表文件路径
    keywords = jieba.analyse.extract_tags(text_seg, topK=5, withWeight=True, allowPOS=('n'))
    sheet.write(row,col, file_name[i])
    col+=1
    for item in keywords:
        keword0=item[0]
        keyword_all.append(keword0)
    sheet.write(row,col,keyword_all[0] + ';' +keyword_all[1] + ';'+keyword_all[2] + ';'+keyword_all[3] + ';'+keyword_all[4] + ';')
    row+=1
f.save('D:/Desktop/keyword.xls')#输出结果保存路径

