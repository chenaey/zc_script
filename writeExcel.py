from docx import Document
import os
import xlrd
import xlwt
from win32com import client as wc
import time
import re

def find_data(stu_num,stu_name,list):
    # print(list)
    for i in list:
        if i[2]==stu_num and i[0]==stu_name:
            # print(i[2], stu_num)
            # print(i)
            # print("##############")
            return i

    # print(stu_name,stu_num)
    # print(stu_name+"_"+stu_num+"_"+"查找错误")
    print(stu_num+"_"+stu_name+"_"+"学号与其他同学重复.请修改后重试")
    return [stu_name, '学号错误或与其他同学学号重复', stu_num, '学号错误或与其他同学学号重复', '信息丢失', '信息丢失', '信息丢失', '信息丢失', '信息丢失']

def read_excel():
    '''
    获取Excle表格基础信息
    :return: 学号姓名等数据
    '''
    path='./initTable/'+os.listdir('./initTable')[0]
    book = xlrd.open_workbook(path)
    sheet1 = book.sheets()[0]
    rows=sheet1.nrows
    list=[]
    for j in range(1,rows):
        stu_name = sheet1.cell(j, 0).value
        cardId = sheet1.cell(j, 1).value
        stu_num = sheet1.cell(j, 2).value
        sex = sheet1.cell(j, 3).value
        xueyuan = sheet1.cell(j, 4).value
        calssName = sheet1.cell(j, 5).value
        year = sheet1.cell(j, 6).value
        isCader = sheet1.cell(j, 7).value
        cader_name = sheet1.cell(j, 8).value
        list.append([stu_name,cardId,stu_num,sex,xueyuan,calssName,year,isCader,cader_name])
    return list

#2
def get_string():
    print("开始写入Excel")

    datalist=[]
    list =read_excel()
    cwd=os.getcwd()

    paths=os.listdir(cwd+'/output/')
    for path in paths:
        stu_name = path.split('_')[1].replace('.docx','')
        stu_num = path.split('_')[0]
        data=find_data(stu_num,stu_name,list)# [stu_name,cardId,stu_num,sex,xueyuan,calssName,year,isCader,cader_name]
        # print(stu_num,stu_name)
        # print(data)
        # print("##################")
        path = cwd+"/output/"+path #文件路径
        document = Document(path) #读入文
        tables = document.tables #获取文件中的表格集

        cardId = data[1]
        sex=data[3]
        xueyuan=data[4]
        calssName=data[5]
        year=data[6]
        isCader=data[7]
        cader_name=data[8]
        think=''
        think_rank=''
        study=''
        study_rank=''
        wt=''
        wt_rank=''
        zh=''
        zh_rank=''
        xy_num=''
        leibie=''
        is_good=''
        # tou = ["姓名", "身份证号码", "学工号", "姓别", "学院", "专业班级", "入学年度", "是否干部", "干部名称","思想分", "思想分排名","学业分","学业分排名",
        # "文体分", "文体分排名","综合分","综合分专业年级排名","专业年级总人数","评优类别","是否评为优秀毕业生"]
        table = tables[0]#获取文件中f的第9个表格
        think =table.cell(9,3).text
        study =table.cell(13,3).text
        wt =table.cell(18,3).text
        zh=table.cell(19,3).text
        if think.__len__()==0:
            think=table.cell(9,2).text
        if study.__len__() == 0:
            study = table.cell(13, 2).text
        if wt.__len__() == 0:
            wt = table.cell(18, 2).text
        if zh.__len__() == 0:
            zh = table.cell(19, 2).text
        datalist.append([ stu_name, cardId,stu_num,sex,xueyuan,calssName,year,isCader,cader_name,think,think_rank,study,study_rank,wt,
                             wt_rank,zh,zh_rank, xy_num, leibie,is_good])
        # print(stu_num,stu_name,think,study,wt,zh)
        # print("#################")

    #写入Excle
    '''
    datalist:数据
    excelList：
    '''
    tou = ["姓名", "身份证号码", "学工号", "姓别", "学院", "专业班级", "入学年度", "是否干部", "干部名称","思想分", "思想分排名", "学业分", "学业分排名",
           "文体分", "文体分排名", "综合分", "综合分专业年级排名", "专业年级总人数", "评优类别", "是否评为优秀毕业生"]
    datalist.insert(0,tou)
    book = xlwt.Workbook()  # 新建一个excel
    sheet = book.add_sheet('sheet')  # 添加一个sheet页
    row = 0  # 控制行
    for stu in datalist:
        col = 0  # 控制列
        for s in stu:  # 再循环里面list的值，每一列
            sheet.write(row, col, s)
            col += 1
        row += 1
    name =os.listdir('./initTable')[0]
    name=name.split(".")[0]+'.xls'
    book.save(name)  # 保存到当前目录下
    print("已保存文件到当前目录下")


#1
def rename():
    w = wc.Dispatch('Word.Application')
    paths = os.listdir('./input')
    cwd=os.getcwd()
    print("开始重命名文件")
    for path in paths:
        in_path = os.listdir('./input/' + path)
        origin_word = in_path[0]
        word = origin_word.replace('-', '_')
        list = word.split('_')
        stu_num = list[0]
        stu_name = list[1]
        word_path =cwd+'\\'+'input\\' + path + '\\' + origin_word
        doc = w.Documents.Open(word_path)
        print(stu_num,stu_name)
        save_path=cwd+"\output\\"+stu_num+"_"+stu_name+'.docx'
        doc.SaveAs(save_path, 16)
        doc.Close()
    print("重命名完成")


def check():
    print("开始校验文件命名是否符合格式...")
    paths = os.listdir('./input')
    error_num=0
    for path in paths:
        try:
            in_path = os.listdir('./input/' + path)
            origin_word = in_path[0]
            origin_word=origin_word.replace('-','_')
            status =re.match('(\d.*)_(.*)_',origin_word)
            if status is None:
                print( "文件名命名错误：" + path)
            if ('.doc' or '.docx') not in origin_word :
                print( "文件错误：" + path)
                error_num=error_num+1
            # if '加分证明' in origin_word :
            #     print( "文件命名错误：" + path)
            #     error_num=error_num+1
        except:
            print("请删除该目录下文件:"+path)
            error_num = error_num + 1
    return error_num


def main():
    err_num= check()
    if err_num==0:
        print("命名校验完成")
        rename()
        get_string()
        time.sleep(10000)

    else:
        print("###################")
        print("请修改命名错误的文件后重试,或重写命名匹配")
        print("运行结束")
        time.sleep(1000)

if __name__ == '__main__':
     main()

