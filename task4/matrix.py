# -*- coding: utf-8 -*-
import xlrd
import datetime

def readxlsx(path):
    file = xlrd.open_workbook(path)
    sheet = file.sheets()[0]
    data_keywords = sheet.col_values(2)[1:]
    keywords_list=[]
    words_list=[]
    for item in data_keywords:
        words=item.split('/')
        words_list.append(words)
        for keyword in words:
            keywords_list.append(keyword)
    keywords=list(set(keywords_list))
    return keywords,words_list

def buildmatrix(keywords,words_list):
    length=len(keywords)+1
    matrix=[['0'] * length for y in range(length)]
    matrix[0][0]='列/行'
    for i in range(0,length-1):
        matrix[i+1][0]=keywords[i]
        matrix[0][i+1]=keywords[i]
    for i in range(1,length):
        for j in range(i+1,length):
            num=0
            for item in words_list:
                if matrix[0][j] in item and matrix[i][0] in item:
                    num+=1
            matrix[i][j]=num
            matrix[j][i]=num
    return matrix

def writetxt(path,matrix):
    with open(path,'a','utf8') as f:
        for i in range(0,len(matrix)):
            for j in range(0,len(matrix)):
                f.write(str(matrix[i][j]) + '\t')
            f.write('\r\n')

def main():
    startTime = datetime.datetime.now()
    path_xls = '/users/apple/desktop/python/第4次训练/材料1.xlsx'
    path_txt = '/users/apple/desktop/python/第4次训练/共现矩阵.txt'
    keywords,words_list=readxlsx(path_xls)
    result = buildmatrix(keywords,words_list)
    writetxt(path_txt,result)
    endTime = datetime.datetime.now()
    print("得到共现矩阵共用时：",(endTime - startTime).seconds ,"秒")

if __name__=='__main__':
    main()
