import os
import jieba
import jieba.posseg as pseg
import sys
import string
import pandas  as pd
import collections
import re
import xlsxwriter
from tqdm.autonotebook import tqdm
from sklearn import feature_extraction
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.feature_extraction.text import CountVectorizer

file_list = []  # 新建一个空列表用于存放文件路径
file_name = []  # 新建一个空列表用于存放文件名字
file_dir = r'D:/Desktop/双碳政策知识图谱构建/政策原文-20221101'  # 遍历的文件夹路径
for files in os.walk(file_dir):  # 遍历指定文件夹及其下的所有子文件夹
    for file in files[2]:  # 遍历每个文件夹里的所有文件，（files[2]:母文件夹和子文件夹下的所有文件信息，files[1]:子文件夹信息，files[0]:母文件夹信息）
        if os.path.splitext(file)[1] == '.PDF' or os.path.splitext(file)[1] == '.pdf':  # 检查文件后缀名,逻辑判断用==
            filename = re.sub(".pdf", "", file)
            file_name.append(filename)  # 筛选后的文件名为字符串，将得到的文件名放进去列表，方便以后调用
            file_list.append(file_dir + '\\' + file)  # 给文件名加入文件夹路径
PATH = "D:\Desktop\output-20221122.csv"
data_content = [line.strip() for line in open(PATH, encoding='utf-8', errors='ignore').readlines()]
vectorizer = CountVectorizer()
transformer = TfidfTransformer()
tfidf = transformer.fit_transform(vectorizer.fit_transform(data_content))
word = vectorizer.get_feature_names()  # 所有文本的关键字
weight = tfidf.toarray()  # 对应的tfidf矩阵
# df_word_idf = list(zip(vectorizer.get_feature_names(),transformer.idf_))
# f = open('D:/Desktop/IDF20221028-2.txt', 'w+', encoding='utf-8')
# for i in range(len(df_word_idf)) :
#     f.write(df_word_idf[i][0]+" "+str(df_word_idf[i][1])+'\n')
f = xlsxwriter.Workbook('D:/Desktop/关键词结果-20221122.xlsx')
sFilePath = 'D:/Desktop/tf-idf'
sheet2 = f.add_worksheet('关键词结果')
row2 = 0
col2 = 0
for i in tqdm(range(len(weight)), colour='blue', desc="已经完成进度", ncols=100):
    dic1 ={}
    for j in range(len(word)):
        dic1.setdefault(word[j],weight[i][j])
    keyword_all=[]
    sheet2.write(row2, col2, file_name[i])
    col2 += 1
    sheet2.write(row2,col2,"关键词")
    col2+=1
    count=0
    fp = open(sFilePath + '/' + file_name[i] + '.txt', 'w+', encoding='UTF-8')
    for k in sorted(dic1, key=dic1.__getitem__, reverse=True):
        #print(k, d[k])
        if  dic1[k]!=0:
            fp.write(k+ "  " + str(dic1[k]) + "\n")
    for k in sorted(dic1, key=dic1.__getitem__, reverse=True):
        keyword_all.append(k)
        count+=1
        if count>=15:                            #输出关键词个数
            sheet2.write(row2, col2,keyword_all[0] + ';' + keyword_all[1] + ';' + keyword_all[2] + ';' + keyword_all[3] + ';'+ keyword_all[4] + ';'+keyword_all[5]+';'+keyword_all[6]+';'+keyword_all[7]+';'+keyword_all[8]+';'+keyword_all[9]+';'+keyword_all[10] + ';'+keyword_all[11]+';'+keyword_all[12]+';'+keyword_all[13]+';'+keyword_all[14])
            row2 += 1
            col2 = 0
            break
        else:
            continue
f.close()


