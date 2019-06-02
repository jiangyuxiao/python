# -*- coding: utf-8 -*-

import xlrd
import codecs
import os
import math

# 读取Excel中关键词并去除首行元素
def read_xls(path,col):
    xl = xlrd.open_workbook(path)
    sheet = xl.sheets()[0]
    raw_data = list(sheet.col_values(col-1)[1:])
    return raw_data

# 读取txt文档
def read_txt(path):
    with codecs.open(path,'r','utf-8','ignore') as f:
        raw_txt = f.read()
        return raw_txt

def write_txt(data,path,title):
    with open(path+'/'+title,'w') as f:
        f.write(data)

# 批量读取文件夹里路径
def read_path(path):
    path_list = []
    for root, dirs, files in os.walk(path):
        for fn in files:
            path = str(root+'/'+fn)
            path_list.append(path)
    return path_list

# 将一篇文章按照'/'进行分割，转换为数组形式
def to_words(text,stopwords):
    words=[]
    data=text.split('/')
    for word in data:
        if word not in stopwords:
            words.append(word)
    return words

# 获取语料库
def get_all_words(txt_path,stopwords):
    allwords=[]
    for path in txt_path:
        words=to_words(read_txt(path),stopwords)
        allwords.append(words)
    return allwords

# 统计包含某一关键词的文档数
def file_count(word,allwords):
    count=0
    for words in allwords:
        if word in set(words):
            count+=1
        else:
            continue
    return count

# 统计某一篇文章的词频，传入数据为文章经分割后的数组形式
def tf_count(words):
    tf={}
    word_dic={}
    for word in words:
        if word in word_dic.keys():
            word_dic[word] += 1
        else:
            word_dic[word] = 1
    for word in word_dic:
        tf[word]=word_dic[word]/len(words)
    return tf

# 统计某一篇文章的TF-IDF
def tfidf_count(words,allwords):
    tf_idf={}
    tf=tf_count(words)
    for word in tf.keys():
        idf = math.log(len(allwords) / (file_count(word, allwords) + 1))
        tf_idf_value= tf[word]*idf
        tf_idf[word]=tf_idf_value
    return tf_idf


def main():
    stopwords_path='/users/apple/desktop/python/第5次训练/词表/stopwords.xlsx'
    stopwords=read_xls(stopwords_path,2)

    data_folder_path = '/users/apple/desktop/python/第6次训练/分词结果'
    write_folder_path = '/users/apple/desktop/python/第6次训练/蒋雨肖_任务六/TF-IDF统计结果'
    txt_path = read_path(data_folder_path)

    allwords=get_all_words(txt_path,stopwords)

    for path in txt_path:
        # read_txt(path)
        words=to_words(read_txt(path),stopwords)
        tfidf=tfidf_count(words,allwords)
        title=os.path.basename(path)
        write_txt(str(tfidf),write_folder_path,title)

if __name__=="__main__":
    main()