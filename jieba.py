import xlwt
import pdfplumber
import re
from fenci import processing
from fenci import seg_depart
from tqdm.autonotebook import tqdm
import sys
import ckpe
sys.path.append("../")
import os
path = os.getcwd()
# 添加用户词典
jieba.load_userdict(path + "\\保留词.txt")

#------------------------------遍历文件夹的所有PDF文件----------------------------#
file_list = []  # 新建一个空列表用于存放文件路径
file_name = []  # 新建一个空列表用于存放文件名字
file_dir = r'D:/Desktop/双碳政策知识图谱构建/政策原文-20221101'  # 遍历的文件夹路径
for files in os.walk(file_dir):  # 遍历指定文件夹及其下的所有子文件夹
    for file in files[2]:  # 遍历每个文件夹里的所有文件，（files[2]:母文件夹和子文件夹下的所有文件信息，files[1]:子文件夹信息，files[0]:母文件夹信息）
        if os.path.splitext(file)[1] == '.PDF' or os.path.splitext(file)[1] == '.pdf':  # 检查文件后缀名,逻辑判断用==
            filename = re.sub(".pdf", "", file)
            file_name.append(filename)  # 筛选后的文件名为字符串，将得到的文件名放进去列表，方便以后调用
            file_list.append(file_dir + '\\' + file)  # 给文件名加入文件夹路径
#--------------------------------批量解析每一个pdf文件----------------------------#
output1 = open("D:\Desktop\双碳政策知识图谱构建\双碳政策结果\双碳政策知识图谱结果20221101\output-202221101-4.csv", 'w', encoding='UTF-8')
for i in tqdm(range(len(file_list)),colour= 'blue',desc = "已经完成进度",ncols = 100):  # 遍历所提取的文件地址
    pdf = pdfplumber.open(file_list[i])  # 循环执行打开pdf文件命令
    pages = pdf.pages  # pages属性获取页数
    key_phrases_all ='' #新建一个列表，存放高质量短语
    text_all = []  # 新建一个列表，存放PDF文件文本解析
    all_text = ''
    for page in pages:  # 遍历pages里面每一页的信息
        text = page.extract_text()  # 提取当前页内容，赋值给text
        text = re.sub("\n", "", text)  # 去除换行
        text = re.sub("\s", "", text)  # 去除空格
        text = re.sub(r'[0-9]+', '', text)
        text = re.sub(r'[a-z]+', '', text)
        text = re.sub(r'[A-Z]+', '', text)
        text = re.sub(r'\ue5cf', '', text)
        text = re.sub(r'\ue5d2', '', text)
        text = re.sub(r'\U001001ba', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}条', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}篇', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}节', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}步', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}师', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}页', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}项', '', text)
        text = re.sub(r'第[\u4e00-\u9fa5]{1,3}号', '', text)
        text = re.sub(r'（.+?）', '', text)
        text = re.sub(r'【.+?】', '', text)
        text_all.append(text)
    k = 0
    while k < len(text_all):
        all_text = all_text + str(text_all[k])
        k+=1
    #print(file_name[i]+all_text + '\n')
    pdf.close()
    #——————-------------------输出分词结果——————------------------#
    line = processing(all_text)
    line_seg = seg_depart(line)
    output1.write(line_seg + '\n')
output1.close()

