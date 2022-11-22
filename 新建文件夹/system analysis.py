import xml.etree.ElementTree as ET
import zipfile
import os

import numpy as np
import pandas as pd
import shutil

# import xlwings as xw

# 代码目录下创建一个文件名为《文件》的文件夹，把答案直接放进去并且改名成《标准答案》，并新建一个考生文件夹放入所有的考生文件，不要放入除docx以外的文件
# 流程：读取文件->解压文件->读取文件xml->分析-》输出成绩->下一个文件循环

# certifi==2022.9.24
# charset-normalizer==2.1.1
# requests==2.28.1
# idna==3.4
# pipreqs==0.4.11
# docopt==0.6.2
# urllib3==1.26.12
# xlrd==2.0.1
# yarg==0.1.9
# openpyxl
# pandas
# append写入表格


# 若没有文件夹要创建文件夹，读取文件相对路径
path_answer = '文件\\标准答案.docx'  # 标准答案存放路径
path_zipfile_answer = '文件\\解压缓存\\标准答案'  # 标准答案解压存放路径
path_exercises = "文件\\考生文件夹"  # 考生文件存放路径
path_zip_exercises = '文件\\解压缓存\\考生文件'  # 考生文件解压存放路径
all_nums = []  # 得分
names = []  # 名字
false_all = []  # 如有错误，说明错误原因

# 标准答案文件解析出来的列表存入这里，用列表的方式存储xml文件中的格式内容
text_tltle_amswer = []  # 标题样式
text_jc_amswer = []  # 居中
pStyle_answer = []  # 标题格式
text_answer = []  # 正文内容
instrText_answer = []  # 图标题注
Font_answer = []  # 字体格式
bsize_answer = []  # 字体加粗
color_answer = []  # 字体颜色
size_answer = []  # 字体大小
paragraph_alignment_answer = []  # 双下划线
spacing_answer = []  # 第二段前后间距
ind_answer = []  # 第二段缩进
headers_answer = []  # 页眉
header_numb_answer = []  # 页码
rFonts_answer = []  # 脚注字体

# 考生文件解析出来的列表存入这里，用列表的方式存储xml文件中的格式内容
text_tltle_exercises = []  # 标题样式
text_jc_exercises = []  # 居中
pStyle_exercises = []  # 标题格式
text_exercises = []  # 正文内容
instrText_exercises = []  # 图标题注
Font_exercises = []  # 字体格式
bsize_exercises = []  # 字体加粗
color_exercises = []  # 字体颜色
size_exercises = []  # 字体大小
paragraph_alignment_exercises = []  # 双下划线
spacing_exercises = []  # 段前后间距
ind_exercises = []  # 缩进
headers_exercises = []  # 页眉
header_numb_exercises = []  # 页码
rFonts_exercises = []  # 脚注字体


# 读取考生所有文件并获取文件名。获取后缀名为.docx的所有文件
def os_student(path):
    dirs = os.listdir(path)
    dirss = []
    for dir1 in dirs:
        if dir1.split('.')[1] == 'docx':
            dirss.append(os.path.join(path, dir1))
        else:
            continue
    return dirss


# 读取标准答案文件解压的xml
def read_xml_answer():  # 解析标准答案解压出来的xml，用element tree将document.xml中的格式内容从根节点获取
    tree = ET.parse("文件\\解压缓存\\标准答案\\word\\document.xml")
    root = tree.getroot()  # 获取根节点
    for child in root:
        # 读取内容   从根节点获取其他节点名字
        for node in child:
            for v2 in node:
                texts = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')  # 匹配出文字内容
                spacings = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')  # 段前段后间距
                for spacing in spacings:
                    for s in list(spacing.attrib.values()):
                        spacing_answer.append(s)
                instrText = v2.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')  # 段前段后间距
                for instrTexts in instrText:
                    instrText_answer.append(instrTexts.text)
                pStyle = v2.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')  # 匹配标题格式
                for style in pStyle:
                    pStyle_answer.append(list(style.attrib.values())[0])
                inds = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')  # 段前段后间距
                for ind in inds:
                    ind_answer.append(list(ind.attrib.values()))
                for Fonts1 in v2:
                    Fonts = Fonts1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')  # 匹配出字体
                    Fonts_list = []
                    for Font in Fonts:
                        Fonts_list.append(list(Font.attrib.values()))
                    for Fonts2 in Fonts_list:
                        for Fonts3 in Fonts2[:2]:
                            Font_answer.append(Fonts3)
                for sz1 in v2:
                    sz = sz1.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')  # 匹配字体大小
                    for size in sz:
                        for i in size.attrib.values():
                            size_answer.append(i)
                if not texts:
                    continue
                else:
                    for text in texts:
                        # print(text.text)
                        if not text.text.split():
                            continue
                        else:
                            text_answer.append(text.text)
                            for v3 in v2:
                                bcs = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bCs')  # 匹配字体加粗
                                for bsize in bcs:
                                    bsize_answer.append(list(bsize.attrib.values()))
                                u = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u')  # 双下划线
                                for paragraph_alignment in u:
                                    paragraph_alignment_answer.append(list(paragraph_alignment.attrib.values()))
                                colors = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')  # 匹配字体颜色
                                for color in colors:
                                    color_answer.append(list(color.attrib.values())[0])
                            # for v4 in v2:
                            #     print(v4.tag)
    # 标题的格式，从numbering.xml中读取，用element tree将document.xml中的格式内容从根节点获取
    try:
        tree_title = ET.parse("文件\\解压缓存\\标准答案\\word\\numbering.xml")
        root_title = tree_title.getroot()  # 获取根节点
        # 读取内容
        for child_title in root_title:
            child1_title = child_title.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl')
            for child2_title in child1_title:
                jc = child2_title.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlJc')
                for jcs in jc:
                    text_jc_amswer.append(list(jcs.attrib.values())[0])
                lvlText = child2_title.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText')
                for text_tltle in lvlText:
                    if list(text_tltle.attrib.values())[0] != '':
                        text_tltle_amswer.append(list(text_tltle.attrib.values())[0])
    except FileNotFoundError:  # 遇到错误中断循环
        pass
    # 页眉，从header.xml中读取
    try:
        for i in range(1, 5):
            tree1 = ET.parse(f"文件\\解压缓存\\标准答案\\word\\header{i}.xml")
            root1 = tree1.getroot()
            for child1 in root1:
                header1 = child1.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                headers = child1.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                for headers2 in header1:
                    headers3 = headers2.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
                    for headers4 in headers3:
                        headers_answer.append(list(headers4.attrib.values())[0])
                for node1 in headers:
                    text5 = node1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    node4 = node1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                    for node5 in node4:
                        node6 = node5.findall(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                    for text6 in text5:
                        headers_answer.append(text6.text)
                        for child4 in child1:
                            pStyle1 = child4.findall(
                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                            for style1 in pStyle1:
                                headers_answer.append(list(style1.attrib.values())[0])
    except FileNotFoundError:  # 遇到错误中断循环
        pass
    # 页码，从footer.xml中读取
    try:
        for i in range(1, 4):  # 支持四种不同页码
            tree2 = ET.parse(f"文件\\解压缓存\\标准答案\\word\\footer{i}.xml")
            root2 = tree2.getroot()
            for child2 in root2:
                for prps in child2:
                    # docPartObj = jcs.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr')
                    # print(jcs.attrib)
                    for jc1 in prps:
                        for jc2 in jc1:
                            jc = jc2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
                            pStyle_numb = jc2.findall(
                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                            for pStyle_numbs in pStyle_numb:
                                header_numb_answer.append(list(pStyle_numbs.attrib.values())[0])
                            for jc3 in jc:
                                header_numb_answer.append(list(jc3.attrib.values())[0])
    except FileNotFoundError:  # 遇到错误中断循环
        pass
    # 脚注，从footnote.xml中读取
    try:
        tree3 = ET.parse("文件\\解压缓存\\标准答案\\word\\footnotes.xml")
        root3 = tree3.getroot()
        for child3 in root3:
            for footnote in child3:
                for p in footnote:
                    rFonts = p.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                    t = p.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    for text in t:
                        if text.text != ' ':
                            # print(text.text)
                            text_answers = ''.join(text.text)
                            rFonts_answer.append(text_answers)
                            for rFont in rFonts:
                                sz1 = rFont.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
                                for sz2 in sz1:
                                    rFonts_answer.append(list(sz2.attrib.values())[0])
                                rFontss = rFont.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                                for rf in rFontss:
                                    rFonts_answer.append(list(rf.attrib.values())[0])
                        else:
                            continue
    except FileNotFoundError:  # 遇到错误中断循环
        pass


# 读取考生文件解压的xml
def read_xml_exercises():  # 解析考生文件解压出来的xml，用element tree将document.xml中的格式内容从根节点获取
    # 标题以及正文内容
    tree = ET.parse(f"{path_zip_exercises}\\word\\document.xml")
    root = tree.getroot()  # 获取根节点
    for child in root:
        # 读取内容 从根节点获取其他节点名字
        for node in child:
            for v2 in node:
                pStyle = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')  # 匹配标题格式
                for style in pStyle:
                    pStyle_exercises.append(list(style.attrib.values())[0])
                    # print(list(style.attrib.values())[0])
                texts = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')  # 匹配出文字内容
                instrText = v2.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')  # 段前段后间距
                for instrTexts in instrText:
                    instrText_exercises.append(instrTexts.text)
                spacings = v2.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')  # 段前段后间距
                for spacing in spacings:
                    for s in list(spacing.attrib.values()):
                        spacing_exercises.append(s)
                inds = v2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')  # 段前段后间距
                for ind in inds:
                    ind_exercises.append(list(ind.attrib.values()))
                for Fonts1 in v2:
                    Fonts = Fonts1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')  # 匹配出字体
                    Fonts_list = []
                    for Font in Fonts:
                        Fonts_list.append(list(Font.attrib.values()))
                    for Fonts2 in Fonts_list:
                        for Fonts3 in Fonts2[:2]:
                            Font_exercises.append(Fonts3)
                for sz1 in v2:
                    sz = sz1.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')  # 匹配字体大小
                    for size in sz:
                        for i in size.attrib.values():
                            size_exercises.append(i)
                if not texts:
                    continue
                else:
                    for text in texts:
                        # print(text.text)
                        if not text.text.split():
                            continue
                        else:
                            text_answer.append(text.text)  # text_exercises.append(text.text)
                            for v3 in v2:
                                bcs = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bCs')  # 匹配字体加粗
                                for bsize in bcs:
                                    bsize_exercises.append(list(bsize.attrib.values()))
                                u = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u')  # 双下划线
                                for paragraph_alignment in u:
                                    paragraph_alignment_exercises.append(
                                        list(paragraph_alignment.attrib.values()))
                                colors = v3.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')  # 匹配字体颜色
                                for color in colors:
                                    color_exercises.append(list(color.attrib.values())[0])
    # 标题的格式
    try:
        tree_title = ET.parse(f"{path_zip_exercises}\\word\\numbering.xml")
        root_title = tree_title.getroot()
        for child_title in root_title:
            child1_title = child_title.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvl')
            for child2_title in child1_title:
                jc = child2_title.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlJc')
                for jcs in jc:
                    text_jc_exercises.append(list(jcs.attrib.values())[0])
                lvlText = child2_title.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lvlText')
                for text_tltle in lvlText:
                    if list(text_tltle.attrib.values())[0] != '':
                        text_tltle_exercises.append(list(text_tltle.attrib.values())[0])
                    else:
                        continue
    except FileNotFoundError:
        pass  # 中断循环
    # 页眉
    try:
        for i in range(1, 5):
            tree1 = ET.parse(f"{path_zip_exercises}\\word\\header{i}.xml")
            root1 = tree1.getroot()
            for child1 in root1:
                header1 = child1.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                headers = child1.findall(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
                for headers2 in header1:
                    headers3 = headers2.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
                    for headers4 in headers3:
                        headers_exercises.append(list(headers4.attrib.values())[0])
                for node1 in headers:
                    text5 = node1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    node4 = node1.findall(
                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                    for node5 in node4:
                        node6 = node5.findall(
                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                        # for node7 in node6:
                        # headers_answer.append(list(node7.attrib.values())[0:3])
                    for text6 in text5:
                        headers_exercises.append(text6.text)
                        for child4 in child1:
                            pStyle1 = child4.findall(
                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                            for style1 in pStyle1:
                                headers_exercises.append(list(style1.attrib.values())[0])
    except FileNotFoundError:
        pass
    # 页码
    try:
        for i in range(1, 4):  # 支持四种不同页码
            tree2 = ET.parse(f"{path_zip_exercises}\\word\\footer{i}.xml")
            root2 = tree2.getroot()
            for child2 in root2:
                for prps in child2:
                    for jc1 in prps:
                        for jc2 in jc1:
                            jc = jc2.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
                            pStyle_numb = jc2.findall(
                                '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
                            for pStyle_numbs in pStyle_numb:
                                header_numb_exercises.append(list(pStyle_numbs.attrib.values())[0])
                            for jc3 in jc:
                                header_numb_exercises.append(list(jc3.attrib.values())[0])
    except FileNotFoundError:
        pass
    # 脚注
    try:
        tree3 = ET.parse(f"{path_zip_exercises}\\word\\footnotes.xml")
        root3 = tree3.getroot()
        for child3 in root3:
            for footnote in child3:
                for p in footnote:
                    rFonts = p.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                    t = p.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                    for text in t:
                        if text.text != ' ':
                            # print(text.text)
                            text_exercises = ''.join(text.text)  # 获取脚注内容
                            rFonts_exercises.append(text_exercises)
                            for rFont in rFonts:
                                sz1 = rFont.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
                                for sz2 in sz1:
                                    rFonts_exercises.append(list(sz2.attrib.values())[0])
                                rFontss = rFont.findall(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
                                for rf in rFontss:
                                    rFonts_exercises.append(list(rf.attrib.values())[0])
                        else:
                            continue
    except FileNotFoundError:
        pass


# 对比xml文件内容
def fraction():
    ten_num = 0  # 分数参数
    false_content = []  # 单个文件错误内容
    num_false = 0  # 错误信息序号
    # text_exercise = ''.join(text_exercises)
    # text_answers = ''.join(text_answer)
    # if text_exercise == text_answers:  # 对比文字内容
    #     print('内容一致')
    #     ten_num += 2
    # else:
    #     print('文章与原内容不一致')

    # 居中
    cuowu1 = str('未居中或左对齐')
    cuowu2 = str('章名或节名样式不正确')
    cuowu3 = str('标题格式与原标题格式不一致')
    cuowu4 = str('字体样式不正确')
    cuowu5 = str('题注和目录不正确')
    cuowu6 = str('字体大小不正确')
    cuowu7 = str('字体未加粗')
    cuowu8 = str('未增加双线')
    cuowu9 = str('颜色不正确')
    cuowu10 = str('段前段后间距不正确')
    cuowu11 = str('缩进不正确')
    cuowu12 = str('页眉不一致')
    cuowu13 = str('页码不正确')
    cuowu14 = str('脚注格式不正确')
    if text_jc_amswer[:2] == text_jc_exercises[:2]:
        print('标题居中，节名左对齐')
        ten_num += 2
    else:
        print(cuowu1)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu1)

    # 对比标题样式
    if text_tltle_amswer[:2] == text_tltle_exercises[:2]:
        print('章名，节名样式正确')
        ten_num += 2
    else:
        print(cuowu2)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu2)

    # 对比标题格式
    if pStyle_answer[:3] == pStyle_exercises[:3]:
        print('标题格式一致')
        ten_num += 2
    else:
        print(cuowu3)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu3)

    # 对比字体
    if Font_answer[:2] == Font_exercises[:2]:
        print('字体样式正确')
        ten_num += 2
    else:
        print(cuowu4)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu4)

    # 表格题注
    if instrText_exercises == instrText_answer:
        # print('目录')
        # print(instrText_answer)
        print('题注和目录正确')
        ten_num += 2
    else:
        print(cuowu5)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu5)

    # 对比字体大小
    # if set(size_exercises) == set(size_answer):
    #   print('字体大小正确')
    #   ten_num += 1
    # else:
    #   print(cuowu6)
    # num_false += 1
    # false_content.append(str(num_false)+'.'+cuowu6)

    # 对比加粗内容
    # if bsize_exercises == bsize_answer:
    #     print('字体加粗')
    #     ten_num += 1
    # else:
    #     print(cuowu7)
    #   num_false+=1
    # false_content.append(str(num_false)+'.'+cuowu7)

    # 下划线样式对比
    # if paragraph_alignment_exercises == paragraph_alignment_answer:
    #     print('增加双线')
    #     ten_num += 1
    # else:
    #     print(cuowu8)
    #    num_false+=1
    # false_content.append(str(num_false)+'.'+cuowu8)

    # 字体颜色对比
    # if set(color_exercises) == set(color_answer):
    #     print('颜色正确')
    #     ten_num += 1
    # else:
    #     print(cuowu9)
    #    num_false+=1
    # false_content.append(str(num_false)+'.'+cuowu9)

    # 段前段后间距和缩进对比
    if set(spacing_answer) == set(spacing_exercises):
        print('段前段后间距正确')
        ten_num += 2
    else:
        print(cuowu10)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu10)

    # 缩进对比
    if ind_exercises[5:] == ind_answer[5:]:
        print('缩进正确')
        ten_num += 2
    else:
        print(cuowu11)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu11)

    # 页眉对比
    if set(headers_answer) == set(headers_exercises):
        print('页眉正确')
        # print(headers_answer)
        ten_num += 2
    else:
        print(cuowu12)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu12)

    # 页码
    if header_numb_exercises == header_numb_answer:
        print('页码正确')
        # print(header_numb_answer)
        ten_num += 2
    else:
        print(cuowu13)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu13)

    # 脚注格式
    # if set(rFonts_answer) == set(rFonts_exercises):
    if rFonts_answer == rFonts_exercises:
        print('脚注正确')
        # print(rFonts_exercises)
        ten_num += 2
    else:
        print(cuowu14)
        num_false += 1
        false_content.append(str(num_false) + '.' + cuowu14)

    # 打印出效果
    # print('页码')
    # print(header_numb_answer)
    # print(header_numb_exercises)
    # print('脚注')
    # print(rFonts_answer)
    # print(rFonts_exercises)
    # print('页眉')
    # print(headers_answer)
    # print(headers_exercises)
    # print('目录')
    # print(instrText_answer)
    # print(instrText_exercises)

    # 错误信息汇总
    false_text = ','.join(false_content)  # 把错误信息拼接在一起
    print(false_text)  # 测试
    false_all.append(false_text)  # 加入全局变量

    # 判断完之后需要对全局变量进行清除
    # 标准答案文件解析出来的列表
    text_tltle_amswer.clear()  # 标题样式
    text_jc_amswer.clear()  # 居中
    pStyle_answer.clear()  # 标题格式
    text_answer.clear()  # 正文内容
    instrText_answer.clear()  # 图标题注
    Font_answer.clear()  # 字体格式
    bsize_answer.clear()  # 字体加粗
    color_answer.clear()  # 字体颜色
    size_answer.clear()  # 字体大小
    paragraph_alignment_answer.clear()  # 双下划线
    spacing_answer.clear()  # 第二段前后间距
    ind_answer.clear()  # 第二段缩进
    headers_answer.clear()  # 页眉
    header_numb_answer.clear()  # 页码
    rFonts_answer.clear()  # 脚注字体

    # 考生文件解析出来的列表
    text_tltle_exercises.clear()  # 标题样式
    text_jc_exercises.clear()  # 居中
    pStyle_exercises.clear()  # 标题格式
    text_exercises.clear()  # 正文内容
    instrText_exercises.clear()  # 图标题注
    Font_exercises.clear()  # 字体格式
    bsize_exercises.clear()  # 字体加粗
    color_exercises.clear()  # 字体颜色
    size_exercises.clear()  # 字体大小
    paragraph_alignment_exercises.clear()  # 双下划线
    spacing_exercises.clear()  # 段前后间距
    ind_exercises.clear()  # 缩进
    headers_exercises.clear()  # 页眉
    header_numb_exercises.clear()  # 页码
    rFonts_exercises.clear()  # 脚注字体
    return ten_num  # 返回分数的参数


# zipfile解压word文件

# 总函数、读取标准答案文件
def answer_docx():
    path = path_answer
    zip_path = path_zipfile_answer
    # file = docx.Document(path)
    zfile = zipfile.ZipFile(path, "r")
    zfile.extractall(path=zip_path, members=zfile.namelist())
    print('解压完毕')
    exercises_docx(os_student(path_exercises))


# 读取考生文件夹文件
def exercises_docx(exercises_name):
    for path_exercises in exercises_name:
        path = path_exercises
        zip_path = path_zip_exercises
        file = path_exercises.split('\\')[2]
        print('读取考生文件:\n' + file)
        zfile = zipfile.ZipFile(path, "r")
        zfile.extractall(path=zip_path, members=zfile.namelist())
        print('解压完毕')
        read_xml_answer()
        read_xml_exercises()
        all_num = (fraction() * 5)
        all_nums.append(all_num)
        del_data("文件\\解压缓存\\考生文件")
    del_data("文件\\解压缓存")
    num_excel()


# 汇总分数输出excel
def num_excel():
    for i in os_student(path_exercises):
        names.append(i.split('\\')[2].split('.')[0])
    excel_all = {'名字': names, '分数': all_nums, '错误': false_all}
    # df = pd.DataFrame(excel_all)  # 构造原始数据文件
    # df = pd.DataFrame.from_dict(excel_all,orient='index')
    # 把不等长的value值输出到excel表格中,并将行和列的数值转换
    df = pd.DataFrame(pd.DataFrame.from_dict(excel_all, orient='index').values.T, columns=list(excel_all.keys()))
    df.to_excel('文件\\成绩.xlsx')  # 生成Excel文件，并存到指定文件路径下
    print('批阅完毕')
    # print(false)


# 删除解压后的word文件夹
def del_data(path):
    shutil.rmtree(path)


# 调用answer_docx()模块
if __name__ == '__main__':
    answer_docx()
