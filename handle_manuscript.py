#!/usr/bin/python

import glob
import re
import docx
import math
import os
import shutil
from win32com.client import Dispatch

org_special_list = {
    "交大医",
    "华东师大",
    "中医大",
    "电机学院",
    "上师大",
    "东方国际",
    "社科院",
    "应技大",
    "立信金融",
    "电力大学",
    "民建市委",
    "民建上海市委",
}

org_replace_dict = {
    "民建市委": "机关",
    "民建上海市委": "机关",
}

docx_list = glob.glob(r'*.doc*')
docx_list.sort(key = lambda x: x.encode("gbk"))
app = Dispatch('Word.Application')
app.visible = False

# open output csv file
fo = open("output.csv", "w")
fo.write("日期,题目,组织,文,图,文稿费,图稿费\n")

for f in docx_list:
    # pick up *.docx/*.doc
    if not re.search(r"\.docx$", f) and not re.search(r"\.doc$", f):
        continue
    # filter temporary files for word
    if re.search(r"^~\$", f) or f == "temp_doc.docx":
        continue

    # print info
    print('正在处理:' + f)

    # date
    date_obj = re.search(r"^[0-9]{8}", f)
    if date_obj:
        date = date_obj.group()
    else:
        date = ""

    '''
    # title
    title_obj = re.search(r"）.*（", f)
    if title_obj:
        title = title_obj.group().strip("）（ ")
    else:
        title = ""
    '''

    # accept all revisions and save it to temp_doc.docx
    doc = app.Documents.Open(os.getcwd() + '/' + f)
    doc.AcceptAllRevisions()
    doc.SaveAs(os.getcwd() + '/temp_doc.docx', 16)
    doc.Close()

    # get full text for getting further info, use first paragraph for title
    text = ""
    title = ""
    file = docx.Document("temp_doc.docx")
    if file:
        for para in file.paragraphs:
            text += para.text + " "
            if not title:
                title = para.text.strip(" ")

    # author
    author_all_obj = re.search(r"[ \t\n\r]文[、/／\＼ 　]{0,1}图[/／\＼:： 　]{1,10}[^ \t\n\r]{2,20}[ \t\n\r]", text)
    if not author_all_obj:
        author_all_obj = re.search(r"[ \t\n\r]图[、/／\＼ 　]{0,1}文[/／\＼:： 　]{1,10}[^ \t\n\r]{2,20}[ \t\n\r]", text)
    if author_all_obj:
        author_all = author_all_obj.group().replace('/',' ').replace('／',' ').replace(':',' ').replace('：',' ').replace('\t',' ').replace('\r',' ').strip(" ").split(" ")[-1]
        author_text = author_all
        author_photo = author_all
    else:
        author_all = ""
        author_text = ""
        author_photo = ""

    if not author_all:
        author_text_obj = re.search(r"[ \t\n\r]文[/／\＼:： 　]{1,10}[^ \t\n\r]{2,20}[ \t\n\r]", text)
        if author_text_obj:
            author_text = author_text_obj.group().replace('/',' ').replace('／',' ').replace(':',' ').replace('：',' ').replace('\t',' ').replace('\r',' ').strip(" ").split(" ")[-1]
        else:
            author_text = ""

        author_photo_obj = re.search(r"[ \t\n\r]图[/／\＼:： 　]{1,10}[^ \t\n\r]{2,20}[ \t\n\r]", text)
        if author_photo_obj:
            author_photo = author_photo_obj.group().replace('/',' ').replace('／',' ').replace(':',' ').replace('：',' ').replace('\t',' ').replace('\r',' ').strip(" ").split(" ")[-1]
        else:
            author_photo = ""

    #print("文/"+author_text+" 图/"+author_photo)

    # charater count and fee
    char_count = len(text)
    if char_count < 300:
        fee = 10
    elif char_count < 500:
        fee = 15
    else:
        fee = math.floor(char_count / 250) * 5 + 10

    if fee > 100:
        fee = 100

    # photo count
    photo_count = len(glob.glob(re.sub('(\([^\(\)]*?\)|（[^（）]*?）)\..*', '', f)+"*.*")) - 1
    
    # special organization
    for special_org in org_special_list:
        if title.startswith(special_org):
            org = special_org
            break
    else:
        org = title[0:2]

    # replace organization
    if org in org_replace_dict:
        org = org_replace_dict[org]
        fee = 0
        photo_count = 0
       
    # write csv
    try:
        fo.write((date + "," + title + "," + org + "," + author_text + "," + author_photo + "," + str(fee) + "," + str(photo_count*10) + "\n").replace('\u2022', ' ').replace('\u2003', ' '))
    except:
        print("================处理文件" + f + "出错！================")
        fo.write("处理出错！")

    # remove temp_doc.docx
    os.remove('temp_doc.docx')

# close output csv file
fo.close()
