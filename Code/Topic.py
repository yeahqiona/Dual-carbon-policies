import builtins
import os
from gensim.models import LdaModel
import pandas as pd
import numpy as np
from gensim.corpora import Dictionary
from gensim import corpora, models
import csv
import pyLDAvis
import re
import xlsxwriter

# ------------------------------遍历文件夹的所有PDF文件----------------------------#
file_list = []  # 新建一个空列表用于存放文件路径
file_name = []  # 新建一个空列表用于存放文件名字
file_dir = r'D:/Desktop/双碳政策知识图谱构建/政策原文-20221012'  # 遍历的文件夹路径
topic_words = [line.strip() for line in open('D:/Desktop/双碳政策知识图谱构建/文件读取/主题总结1.txt', encoding='UTF-8').readlines()]
for files in os.walk(file_dir):  # 遍历指定文件夹及其下的所有子文件夹
    for file in files[2]:  # 遍历每个文件夹里的所有文件，（files[2]:母文件夹和子文件夹下的所有文件信息，files[1]:子文件夹信息，files[0]:母文件夹信息）
        if os.path.splitext(file)[1] == '.PDF' or os.path.splitext(file)[1] == '.pdf':  # 检查文件后缀名,逻辑判断用==
            filename = re.sub(".pdf", "", file)
            file_name.append(filename)  # 筛选后的文件名为字符串，将得到的文件名放进去列表，方便以后调用
            file_list.append(file_dir + '\\' + file)  # 给文件名加入文件夹路径
# ——————————————————————————————准备数据————————————————————————————#
PATH = "D:\Desktop\双碳政策知识图谱构建\双碳政策结果\双碳政策知识图谱结果20221101\output-202221101-3.csv"
file_object2 = [line.strip() for line in open(PATH, encoding='utf-8', errors='ignore').readlines()]  # 一行行的读取内容
f = xlsxwriter.Workbook('D:/Desktop/双碳政策知识图谱构建/LDA主题模型/政策主题结果（10个主题50轮）-20221103-1.xlsx')
sheet1 = f.add_worksheet('文档-主题分布')
sheet2 = f.add_worksheet('文档-概率分布')
sheet3 = f.add_worksheet('概率阈值-主题数量分布')
sheet4 = f.add_worksheet()
row1 = 0
col1 = 0
row2 = 0
col2 = 0
row3 = 0
col3 = 0
row4 = 0
col4 = 0
data_set = []  # 建立存储分词的列表
for i in range(len(file_object2)):
    result = []
    seg_list = file_object2[i].split()
    for w in seg_list:  # 读取每一行分词
        result.append(w)
    data_set.append(result)
dictionary = corpora.Dictionary(data_set)  # 构建词典
corpus = [dictionary.doc2bow(text) for text in data_set]
# ——————————————————主题模型训练并输出各个主题的words———————————————————#
lda = LdaModel(corpus=corpus, id2word=dictionary, num_topics=10, passes=50, random_state=1, minimum_probability=0)
topic_list = lda.print_topics(num_words=15)
for i in range(len(topic_list)):
    sheet1.write(row1, col1, '主题{}'.format(i + 1))
    col1 += 1
    sheet1.write(row1, col1, topic_list[i][1])
    row1 += 1
    col1 = 0
row1 += 1
z = 0
# ————————————————————输出每个文档最有可能对应的主题—————————————————————#
for i in lda.get_document_topics(corpus)[:]:
    a = [0] * 10
    listj = []
    for j in i:
        listj.append(j[1])
    sheet2.write(row2, 0, file_name[z])
    sheet4.write(row2, 0, file_name[z])
    for l in range(len(listj)):
        sheet2.write(row2, l + 1, listj[l])  # 文档-概率分布
        if (listj[l] >= 0.1):
            a[0] += 1
            col4 += 1
            sheet4.write(row2, col4, topic_words[l])
            # sheet1.write(row1, col1, file_name[z])
            # col1 += 1
            # sheet1.write(row1, col1, '主题{}'.format(l + 1))  # 输出主题序号
            # col1 += 1
            # sheet1.write(row1, col1, topic_list[l][1])
            # row1 += 1
            # col1 = 0
        if (listj[l] >= 0.2):
            a[1] += 1
        if (listj[l] >= 0.3):
            a[2] += 1
        if (listj[l] >= 0.4):
            a[3] += 1
        if (listj[l] >= 0.5):
            a[4] += 1
        if (listj[l] >= 0.6):
            a[5] += 1
        if (listj[l] >= 0.7):
            a[6] += 1
        if (listj[l] >= 0.8):
            a[7] += 1
        if (listj[l] >= 0.9):
            a[8] += 1
    sheet3.write(row3, 0, file_name[z])
    col3 += 1
    for x in range(len(a)):
        sheet3.write(row3, col3, a[x])  # 概率阈值-主题数量分布
        col3 += 1
    # —————————————————————————文档-主题分布——————————————————————————#
    bz = listj.index(max(listj))
    sheet1.write(row1, 0, file_name[z])
    sheet1.write(row1, 1, '主题{}'.format(i[bz][0] + 1))  # 输出主题序号
    sheet1.write(row1, 2, topic_list[i[bz][0]][1])  # 输出主题
    row1 += 1

    z += 1
    row2 += 1
    row3 += 1
    col3 = 0
    col4 = 0
f.close()
# ————————————————————————————气泡图—————————————————————————————#
import pyLDAvis.gensim_models as gensims
import pyLDAvis

data = gensims.prepare(lda, corpus, dictionary)
pyLDAvis.save_html(data, 'D:/Desktop/双碳政策知识图谱构建/LDA主题模型/topic-20221103(11个主题).html')
