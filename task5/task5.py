# -*- coding: utf-8 -*-

import xlrd
import codecs
import os

# 读取Excel中关键词并去除首行元素
def read_xls(path,col):
    xl = xlrd.open_workbook(path)
    sheet = xl.sheets()[0]
    raw_data = list(sheet.col_values(col-1)[1:])
    return raw_data

# 读取txt文档
def read_txt(path):
    with codecs.open(path,'r','utf8') as f:
        raw_txt = f.read()
        return raw_txt

def write_txt(data,path,title):
    with open(path+'/'+title,'w') as f:
        f.write(data)

# 批量读取文件夹里路径
def read_path(path):  #输入文件夹路径
    path_list = []
    for root, dirs, files in os.walk(path):
        for fn in files:
            path = str(root+'/'+fn)
            path_list.append(path)
    return path_list  #输出文件夹内所有文件的路径

def RMM(data,words,stopwords):
    RMM_result=''
    max_length=12
    while len(data)>0:
        if len(data)<=max_length:
            text=data
        else:
            text=data[-max_length:]
        for i in range(len(text)):
            if text in stopwords or text=='\r\n':
                data=data[:-len(text)]
                break
            elif text in words or len(text)==1:
                RMM_result= str(text)+'/'+RMM_result
                data=data[:-len(text)]
                break
            else:
                text=text[1:]
    return RMM_result

def main():
    word_path='/users/apple/desktop/python/第5次训练/词表/words.xlsx'
    stopwords_path='/users/apple/desktop/python/第5次训练/词表/stopwords.xlsx'

    folder_path = '/users/apple/desktop/python/第5次训练/分词文本'
    write_folder_path='/users/apple/desktop/python/第5次训练/蒋雨肖_任务五/分词结果'

    words=read_xls(word_path,2)
    stopwords=read_xls(stopwords_path,2)

    txt_path=read_path(folder_path)
    for path in txt_path:
        text=read_txt(path)
        RMM_result=RMM(text,words,stopwords)
        title=os.path.basename(path)
        write_txt(RMM_result,write_folder_path,title)

if __name__=="__main__":
    main()