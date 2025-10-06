def analysis(sentences):
    support = []
    forbid = []
    support_words = [line.strip() for line in open('D:/Desktop/双碳政策知识图谱构建/文件读取/support_words.txt', encoding='UTF-8').readlines()]  # 读取支持条文词典
    forbid_words = [line.strip() for line in open('D:/Desktop/双碳政策知识图谱构建/文件读取/forbid_words.txt', encoding='UTF-8').readlines()]  # 读取禁止词典
    sentences = split_sents(sentences)
    stop_words = [line.strip() for line in open('D:/Desktop/双碳政策知识图谱构建/文件读取/stop_word.txt', encoding='UTF-8').readlines()]  # 数据清洗词典
    # 根据支持性词语对分句结果进行检索
    global row4
    global col4
    col4 = 0
    for stop_word in stop_words:
        for sentence in sentences:
            if stop_word in sentence:
                sentences.remove(sentence)
    for support_word in support_words:
        for sentence in sentences:
            if (support_word in sentence)and(len(sentence)>10):
                sheet4.write(row4, col4, file_name[i])
                col4+=1
                sheet4.write(row4, col4,"扶持条文")
                col4+=1
                sentences.remove(sentence)
                sentence=sentence.lstrip(string.digits)
                sentence = sentence.lstrip('．—一123456789')
                sentence = sentence.lstrip('、')
                sentence = sentence.lstrip('一二三四五六七八九是')
                sheet4.write(row4,col4,'{}。'.format(sentence))
                row4+=1
                col4=0
    # if (len(support) == 0):
    #     sheet4.write(row4, col4, file_name[i])
    #     col4+=1
    #     sheet4.write(row4,col4,'该政策文本暂未检索到支持性条文')
    #     row4+=1
    # 根据禁止性词语对分句结果进行检索
    global row5
    global col5
    col5 = 0
    for forbid_word in forbid_words:
        for sentence in sentences:
            if (forbid_word in sentence)and(len(sentence)>10):
                forbid.append(sentence)
                sheet5.write(row5,col5,file_name[i])
                col5+=1
                sheet5.write(row5, col5, '禁止政策')
                col5+=1
                sentences.remove(sentence)
                sentence=sentence.lstrip(string.digits)
                sentence = sentence.lstrip('．—一123456789')
                sentence = sentence.lstrip('、')
                sentence = sentence.lstrip('一二三四五六七八九是')
                sheet5.write(row5, col5,'{}。'.format(sentence))
                row5 += 1
                col5 = 0