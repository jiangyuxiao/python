# -*- coding: utf-8 -*-
import xlrd

def write_txt(path,data):
    with open(path,'w',encoding = 'utf8') as f:
        for item in data:
            f.write(item[0]+'\t'+item[1]+'\t'+item[2]+'\n')
    f.close()

def read_xlsx(path):
    xl=xlrd.open_workbook(path)
    sheet=xl.sheets()[0]
    data=[]
    for i in range(1,sheet.nrows):
        data.append(sheet.row_values(i))
    return data

def merge(data):
    ID=[]
    list=[]
    for item in data:
        if item[0] in ID:
            list[ID.index(item[0])][2]+=';'+item[2]
        else:
            ID.append(item[0])
            list.append(item)
    return list


def main():
    path='/users/apple/desktop/python/第2次训练/'
    data=read_xlsx(path+'材料2.xlsx')
    list=merge(data)
    write_txt(path+'作者合并.txt',list)

if __name__ == '__main__':
    main()