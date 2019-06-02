# -*- coding: utf-8 -*-
import xlrd

#统计词频的函数
def count(list):
    dic = {}
    for item in list:
        if item in dic:
            dic[item] += 1
        else:
            dic[item] = 1
    data=sorted(dic.items(),key=lambda d:d[1],reverse=True)
    return data

#写入文件函数
def write_txt(path,data):
    with open(path,'w',encoding = 'utf8') as f:
        for item in data:
            f.write(item[0]+'\t'+str(item[1])+'\n')


def read_xlsx(path):
    xl = xlrd.open_workbook(path)
    sheet = xl.sheets()[0]
    gjc=sheet.col_values(2)[1:]
    data_gjc=[]
    for item in gjc:
        item_list=item.split('/')
        for i in item_list:
            data_gjc.append(i)
    data_qk=sheet.col_values(3)[1:]
    return data_gjc,data_qk

def main():
    path = '/users/apple/desktop/python/第2次训练/'
    data_gjc,data_qk=read_xlsx(path+'材料1.xlsx')
    keyword=count(data_gjc)
    journal=count(data_qk)
    write_txt(path+'关键词词频.txt',keyword)
    write_txt(path+'期刊频次.txt',journal)


if __name__ == '__main__':
    main()