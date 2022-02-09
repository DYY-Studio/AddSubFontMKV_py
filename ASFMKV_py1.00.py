# -*- coding: UTF-8 -*-
# *************************************************************************
#
# 请使用支持 UTF-8 NoBOM 并最好带有 Python 语法高亮的文本编辑器
# Windows 7 的用户请最好不要使用 写字板/记事本 打开本脚本
#
# *************************************************************************

# 调用库，请不要修改
import shutil
from fontTools import ttLib
from fontTools import subset
import chardet
import os
from os import path
import sys
import re
import winreg
import zlib
import json
from colorama import init
from datetime import datetime

# 初始化环境变量
# *************************************************************************
# 自定义变量
# 修改注意: Python的布尔类型首字母要大写 True 或 False，在名称中有单引号的，需要输入反斜杠转义 \'
# *************************************************************************
# extlist 可输入的视频媒体文件扩展名
extlist = 'mkv;mp4;mts;mpg;flv;mpeg;m2ts;avi;webm;rm;rmvb;mov;mk3d;vob'
# *************************************************************************
# mkvout 媒体文件输出目录(封装)
# 在最前方用"?"标记来表示这是一个子目录
# 注意: 在Python中需要在左侧引号前加 r 来保留 Windows 路径中的反斜杠，路径末尾不需要反斜杠
mkvout = ''
# *************************************************************************
# assout 字幕文件输出目录
# 在最前方用"?"标记来表示这是一个子目录
# 注意: 在Python中需要在左侧引号前加 r 来保留 Windows 路径中的反斜杠，路径末尾不需要反斜杠
assout = '?subs'
# *************************************************************************
# fontout 字体文件输出目录
# 在最前方用"?"标记来表示这是一个子目录
# 注意: 在Python中需要在左侧引号前加 r 来保留 Windows 路径中的反斜杠，路径末尾不需要反斜杠
fontout = '?Fonts'
# *************************************************************************
# fontin 自定义字体文件夹，可做额外字体源，必须是绝对路径
# 可以有多个路径，路径之间用"?"来分隔
# 注意: 在Python中需要在左侧引号前加 r 来保留 Windows 路径中的反斜杠，路径末尾不需要反斜杠
fontin = r''
# *************************************************************************
# notfont (封装)字体嵌入
#   True    始终嵌入字体
#   False   不嵌入字体，不子集化字体，不替换字幕中的字体信息
notfont = False
# *************************************************************************
# sublang (封装)字幕语言
#   会按照您所输入的顺序给字幕赋予语言编码，如果字幕数多于语言数，多出部分将赋予最后一种语言编码
#   IDX+SUB 的 DVDSUB 由于 IDX 文件一般有语言信息，对于DVDSUB不再添加语言编码
#   各语言之间应使用半角分号 ; 隔开，如 'chi;chi;chi;und'
#   可以在 mkvmerge -l 了解语言编码
#   可以只有一个 # 号让程序在运行时再询问您，如 '#'
sublang = ''
# *************************************************************************
# matchStrict 严格匹配
#   True   媒体文件名必须在字幕文件名的最前方，如'test.mkv'的字幕可以是'test.ass'或是'test.sub.ass'，但不能是'sub.test.ass'
#   False  只要字幕文件名中有媒体文件名就行了，不管它在哪
matchStrict = True
# *************************************************************************
# rmAssIn (封装)如果输入文件是mkv，删除mkv文件中原有的字幕
rmAssIn = True
# *************************************************************************
# rmAttach (封装)如果输入文件是mkv，删除mkv文件中原有的附件
rmAttach = True
# *************************************************************************
# v_subdir 视频的子目录搜索
v_subdir = False
# *************************************************************************
# s_subdir 字幕的子目录搜索
s_subdir = False
# *************************************************************************
# copyfont (ListAssFont)拷贝字体到源文件夹
copyfont = False
# *************************************************************************
# resultw (ListAssFont)打印结果到源文件夹
resultw = False
# *************************************************************************

# 以下变量谨慎更改
# subext 可输入的字幕扩展名，按照python列表语法
subext = ['ass', 'ssa', 'srt', 'sup', 'idx']


# 以下环境变量不应更改
# 编译 style行 搜索用正则表达式
style_read = re.compile('.*\nStyle:.*')
cillegal = re.compile(r'[\\/:\*"><\|]')
# 切分extlist列表
extlist = [s.strip(' ').lstrip('.').lower() for s in extlist.split(';') if len(s) > 0]
# 切分sublang列表
sublang = [s.strip(' ').lower() for s in sublang.split(';') if len(s) > 0]
# 切分fontin列表
fontin = [s.strip(' ') for s in fontin.split('?') if len(s) > 0]
fontin = [s for s in fontin if path.isdir(s)]
langlist = []
extsupp = []
dupfont = {}
init()

# ASS分析部分
# 需要输入
#   asspath: ASS文件的绝对路径
# 可选输入
#   fontlist: 可以更新fontlist，用于多ASS同时输入的情况，结构见下
#   onlycheck: 只确认字幕中的字体，仅仅返回fontlist
# 将会返回
#   fullass: 完整的ASS文件内容，以行分割
#   fontlist: 字体与其所需字符 { 字体 : 字符串 }
#   styleline: 样式内容的起始行
#   font_pos: 字体在样式中的位置
def assAnalyze(asspath: str, fontlist: dict = {}, onlycheck: bool = False):
    global style_read
    eventline = 0
    style_pos = 0
    style_pos2 = 0
    text_pos = 0
    font_pos = 0
    styleline = 0
    event_read = re.compile('.*\nDialogue:.*')
    style = re.compile(r'^\[V4.*Styles\]$')
    event = re.compile(r'^\[Events\]$')
    #识别文本编码并读取整个SubtitleStationAlpha文件到内存
    ass = open(asspath, mode='rb')
    ass_b = ass.read()
    ass_code = chardet.detect(ass_b)['encoding'].lower()
    ass.close()
    if not ass_code is None:
        ass = open(asspath, encoding=ass_code, mode='r')
    fullass = ass.readlines()
    ass.close()

    #在文件中搜索Styles标签和Events标签来确认起始行
    for s in range(0, len(fullass)):
        if re.match(style, fullass[s]) is not None:
            styleline = s
        elif re.match(event, fullass[s]) is not None:
            eventline = s
        if styleline != 0 and eventline != 0:
            break

    #获取Style的 Format 行，并用半角逗号分割
    style_format = ''.join(fullass[styleline + 1].split(':')[1:]).strip(' ').split(',')
    #确定Style中 Name 和 Fontname 的位置
    for i in range(0, len(style_format)):
        if style_format[i].lower().strip(' ').replace('\n', '') == 'name':
            style_pos = i
        elif style_format[i].lower().strip(' ').replace('\n', '') == 'fontname':
            font_pos = i
        if style_pos != 0 and font_pos != 0:
            break

    #获取 字体表 与 样式字体对应表
    style_font = {}
    #style_font 词典内容:
    # { style : font }
    #fontlist 词典内容:
    # { font : text }
    # Text: 使用该字体的文本
    for i in range(styleline + 2, len(fullass)):
        if len(fullass[i].split(':')) < 2: 
            if re.search(style_read, '\n'.join(fullass[i + 1:])) is None: 
                break
            else:
                continue
        styleStr = ''.join(fullass[i].split(':')[1:]).strip(' ').split(',')
        font_key = styleStr[font_pos].lstrip('@')
        fontlist.setdefault(font_key, '')
        style_font[styleStr[style_pos]] = styleStr[font_pos]

    if onlycheck:
        fullass = []
        style_font.clear()
        return None, fontlist, None, None
    #print(fontlist)

    #提取Event的 Format 行，并用半角逗号分割
    event_format = ''.join(fullass[eventline + 1].split(':')[1:]).strip(' ').split(',')
    #确定Event中 Style 和 Text 的位置
    for i in range(0, len(event_format)):
        if event_format[i].lower().replace('\n', '').strip(' ') == 'style':
            style_pos2 = i
        elif event_format[i].lower().replace('\n', '').strip(' ') == 'text':
            text_pos = i
        if style_pos2 != 0 and text_pos != 0:
            break

    #获取 字体的字符集
    #先获取 Style，用style_font词典查找对应的 Font
    #再将字符串追加到 fontlist 中对应 Font 的值中
    for i in range(eventline + 2, len(fullass)):
        eventline_sp = fullass[i].split(':')
        if len(eventline_sp) < 2: 
            if re.search(event_read, '\n'.join(fullass[i + 1:])) is None: 
                break
            else:
                continue
        #print(fullass[i])
        if eventline_sp[0].strip(' ').lower() == 'comment': continue
        eventline_sp = ''.join(eventline_sp[1:]).split(',')
        eventfont = style_font.get(eventline_sp[style_pos2].lstrip('*'))
        if not eventfont is None and not fontlist.get(eventfont) is None:
            eventtext = re.sub(r'(\{.*?\})|(\\[hnN])|(m .*\w*.*)|(\s)', '', ','.join(eventline_sp[text_pos:]))
            #print(eventfont, eventtext)
            #print(eventtext, ','.join(eventline_sp[text_pos:]))
            for i in range(0,len(eventtext)):
                if not eventtext[i] in fontlist[eventfont]:
                    fontlist[eventfont] = fontlist[eventfont] + eventtext[i]

    print('\033[1m字幕所需字体\033[0m')
    fl_popkey = []
    for s in fontlist.keys():
        if len(fontlist[s]) == 0:
            fl_popkey.append(s)
            #print('跳过没有字符的字体\"{0}\"'.format(s))
        else: print('\033[1m\"{0}\"\033[0m: 字符数[\033[1;33m{1}\033[0m]'.format(s, len(fontlist[s])))
    if len(fl_popkey) > 0:
        for s in fl_popkey:
            fontlist.pop(s)
    return fullass, fontlist, styleline, font_pos

# 获取字体文件列表
# 接受输入
#   customPath: 用户指定的字体文件夹
#   font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）
#   noreg: 只从用户提供的customPath获取输入
# 将会返回
#   filelist: 字体文件清单 [[ 字体绝对路径, 读取位置（'': 注册表, '0': 自定义目录） ], ...]
#   font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）
def getFileList(customPath: list = [], font_name: dict = {}, noreg: bool = False):
    filelist = []

    if not noreg:
        # 从注册表读取
        fontkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
        fontkey_num = winreg.QueryInfoKey(fontkey)[1]
        #fkey = ''
        try:
            fontkey10 = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
            fontkey10_num = winreg.QueryInfoKey(fontkey10)[1]
            if fontkey10_num > 0:
                for i in range(fontkey10_num):
                    p = winreg.EnumValue(fontkey10, i)[1]
                    #n = winreg.EnumValue(fontkey10, i)[0]
                    if path.exists(p): 
                        # test = n.split('&')
                        # if len(test) > 1:
                        #     for i in range(0, len(test)):
                        #         font_name[re.sub(r'\(.*?\)', '', test[i].strip(' '))] = [p, i]
                        # else: font_name[re.sub(r'\(.*?\)', '', n.strip(' '))] = [p, 0]
                        filelist.append([p, ''])
                        # test = path.splitext(path.basename(p))[0].split('&')
        except:
            pass
        for i in range(fontkey_num):
            k = winreg.EnumValue(fontkey, i)[1]
            #n = winreg.EnumValue(fontkey, i)[0]
            pk = path.join(r'C:\Windows\Fonts', k)
            if path.exists(pk): 
                # test = n.split('&')
                # if len(test) > 1:
                #     for i in range(0, len(test)):
                #         font_name[re.sub(r'\(.*?\)', '', test[i].strip(' '))] = [pk, i]
                # else: font_name[re.sub(r'\(.*?\)', '', n.strip(' '))] = [pk, 0]
                filelist.append([pk, ''])

    # 从定义的文件夹读取
    # fontspath = [r'C:\Windows\Fonts', path.join(os.getenv('USERPROFILE'),r'AppData\Local\Microsoft\Windows\Fonts')]
    if customPath is None: customPath == []
    if len(customPath) > 0:
        print('\033[1;33m请稍等，正在获取自定义文件夹中的字体\033[0m')
        for s in customPath:
            if not path.isdir(s): continue
            for r, d, f in os.walk(s):
                for p in f:
                    p = path.join(r, p) 
                    if path.splitext(p)[1][1:].lower() not in ['ttf', 'ttc', 'otc', 'otf']: continue
                    filelist.append([path.join(s, p), 'xxx'])

    #print(font_name)
    #os.system('pause')
    return filelist, font_name

#字体处理部分
# 需要输入
#   fl: 字体文件列表
# 可选输入
#   f_n: 默认新建一个，可用于更新font_name
# 将会返回
#   font_name: 字体内部名称与绝对路径的索引词典
# 会对以下全局变量进行变更
#   dupfont: 重复字体的名称与其路径词典

# font_name 词典结构
# { 字体名称 : [ 字体绝对路径 , 字体索引 (仅用于TTC/OTC; 如果是TTF/OTF，默认为0) ] }
# dupfont 词典结构
# { 重复字体名称 : [ 字体1绝对路径, 字体2绝对路径, ... ] }
def fontProgress(fl: list, f_n: dict = {}) -> dict:
    global dupfont
    #print(fl)
    flL = len(fl)
    print('\033[1;32m正在读取字体信息...\033[0m')
    for si in range(0, flL):
        s = fl[si][0]
        fromCustom = False
        if len(fl[si][1]) > 0: fromCustom = True
        ext = path.splitext(s)[1][1:]
        if ext.lower() not in ['ttf','ttc','otf','otc']: continue
        print('\r' + '\033[1;32m{0}/{1} {2:.2f}% \033[0m'.format(si + 1, flL, ((si + 1)/flL)*100, ), end='', flush=True)
        if ext.lower() in ['ttf', 'otf']:
            try:
                tc = [ttLib.TTFont(s, lazy=True)]
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTC/OTC模式读取\033[0m'.format(s, sys.exc_info()))
                try:
                    tc = ttLib.TTCollection(s, lazy=True)
                    print('\033[1;34m[WARNING] 错误的字体扩展名\"{0}\" \033[0m'.format(s))
                except:
                    print('\033[1;31m\n[ERROR] \"{0}\": {1}\033[0m'.format(s, sys.exc_info()))
                    continue
        else:
            try:
                tc =  ttLib.TTCollection(s, lazy=True)
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTF/OTF模式读取\033[0m'.format(s, sys.exc_info()))
                try:
                    tc = ttLib.TTCollection(s, lazy=True)
                    print('\033[1;34m[WARNING] 错误的字体扩展名\"{0}\" \033[0m'.format(s))
                except:
                    print('\033[1;31m\n[ERROR] \"{0}\": {1}\033[0m'.format(s, sys.exc_info()))
                    continue
            #f_n[path.splitext(path.basename(s))[0]] = [s, 0]
        for ti in range(0, len(tc)):
            t = tc[ti]
            for ii in range(0, len(t['name'].names)):
                name = t['name'].names[ii]
                if name.nameID == 4:
                    namestr = ''
                    try:
                        namestr = name.toStr()
                    except:
                        namestr = name.toBytes().decode('utf-16-be', errors='ignore')
                        try:
                            if len([i for i in name.toBytes() if i == 0]) > 0:
                                nnames = t['name']
                                namebyte = b''.join([bytes.fromhex('{:0>2}'.format(hex(i)[2:])) for i in name.toBytes() if i > 0])
                                nnames.setName(namebyte,
                                    name.nameID, name.platformID, name.platEncID, name.langID)
                                namestr = nnames.names[ii].toStr()
                                print('\n\033[1;33m已修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(s), namestr))
                                #os.system('pause')
                            else: namestr = name.toBytes().decode('utf-16-be', errors='ignore')
                        except:
                            #print('\n\033[1;34m{0}\033[0m'.format(name.toBytes().decode('utf-16-be', errors='ignore')))
                            print('\n\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(s), namestr))
                    if namestr is None: continue
                    namestr = namestr.strip(' ')
                    #print(namestr, path.basename(s))
                    if f_n.get(namestr) is not None:
                        dupp = f_n[namestr][0]
                        if dupp != s and path.splitext(path.basename(dupp))[0] != path.splitext(path.basename(s))[0] and not fromCustom:
                            print('\033[1;35m[WARNING] 字体\"{0}\"与字体\"{1}\"的名称\"{2}\"重复！\033[0m'.format(path.basename(f_n[namestr][0]), 
                            path.basename(s), namestr))
                            if dupfont.get(namestr) is not None:
                                if s not in dupfont[namestr]:
                                    dupfont[namestr].append(s)
                            else:
                                dupfont[namestr] = [dupp, s]
                    else: f_n[namestr] = [s, ti]
                    #f_n[namestr] = [s, ti]
                    # if f_n.get(fname) is None: f_n[fname] = [[namestr], s]
                    # #print(fname, name.toStr(), f_n.get(fname))
                    # if namestr not in f_n[fname][0]: 
                    #     f_n[fname][0] = f_n[fname][0] + [namestr]
        tc[0].close()
    return f_n

#print(filelist)
#if path.exists(fontspath10): filelist = filelist.extend(os.listdir(fontspath10))

#for s in font_name.keys(): print('{0}: {1}'.format(s, font_name[s]))

# 系统字体完整性检查，检查是否有ASS所需的全部字体，如果没有，则要求拖入
# 需要以下输入
#   fontlist: 字体与其所需字符（只读字体部分） { ASS内的字体名称 : 字符串 }
#   font_name: 字体名称与字体路径对应词典 { 字体名称 : [ 字体绝对路径, 字体索引 ] }
# 以下输入可选
#   assfont: 结构见下，用于多ASS文件时更新列表
#   onlycheck: 缺少字体时不要求输入
# 将会返回以下
#   assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称 ]}
#   font_name: 同上，用于有新字体拖入时对该词典的更新
def checkAssFont(fontlist: dict, font_name: dict, assfont: dict = {}, onlycheck: bool = False):
    for s in fontlist.keys():
        cok = False
        if s not in font_name:
            font_name_cache = {}
            for ss in font_name.keys():
                if ss.lower() == s or ss.upper() == s:
                    font_name_cache[ss.lower()] = font_name[ss]
                    cok = True
            font_name.update(font_name_cache)
        else: cok = True
        if not cok:
            if not onlycheck:
                print('\033[1;31m[ERROR] 缺少字体\"{0}\"\n请输入追加的字体文件或其所在字体目录的绝对路径\033[0m'.format(s))
                addFont = {}
                inFont = ''
                while inFont == '':
                    inFont = input().strip('\"').strip(' ')
                    if path.exists(inFont):
                        if path.isdir(inFont):
                            addFont = fontProgress(getFileList([inFont], noreg=True)[0])
                        else:
                            addFont = fontProgress([[inFont, '0']])
                        if s not in addFont.keys():
                            if path.isdir(inFont):
                                print('\033[1;31m[ERROR] 输入路径中\"{0}\"没有所需字体\"{1}\"\033[0m'.format(inFont, s))
                            else: print('\033[1;31m[ERROR] 输入字体\"{0}\"不是所需字体\"{1}\"\033[0m'.format('|'.join(addFont.keys()), s))
                            inFont = ''
                        else:
                            font_name.update(addFont)
                            cok = True
                    else:
                        print('\033[1;31m[ERROR] 您没有输入任何字符！\033[0m')
                        inFont = ''
            else:
                assfont['?'.join([s, s])] = ['', s]
        if cok:
            font_path = font_name[s][0]
            font_index = font_name[s][1]
            dict_key = '?'.join([font_path, str(font_index)])
            if assfont.get(dict_key) is None:
                assfont[dict_key] = [fontlist[s], s]
            else:
                if assfont[dict_key][1] == font_index:
                    newfname = assfont[dict_key][2]
                    if s != newfname:
                        newfname = '|'.join([s, newfname])
                    newstr = assfont[dict_key][1]
                    newstr2 = ''
                    for i in range(0, len(newstr)):
                        if newstr[i] not in fontlist[s]:
                            newstr2 = newstr2 + newstr[i]
                    assfont[dict_key] = [fontlist[s] + newstr2, newfname]
                else:
                    assfont[dict_key] = [fontlist[s], s]
            #print(assfont[dict_key])

    return assfont, font_name

# print('正在输出字体子集字符集')
# for s in fontlist.keys():
#     logpath = '{0}_{1}.log'.format(path.join(os.getenv('TEMP'), path.splitext(path.basename(asspath))[0]), s)
#     log = open(logpath, mode='w', encoding='utf-8')
#     log.write(fontlist[s])
#     log.close()

# 字体内部名称变更
def getNameStr(name, subfontcrc: str) -> str:
    namestr = ''
    nameID = name.nameID
    if nameID in [1,3,4,6]:
        namestr = subfontcrc
    else:
        try:
            namestr = name.toStr()
        except:
            namestr = name.toBytes().decode('utf-16-be', errors='ignore')  
    return namestr

# 字体子集化
# 需要以下输入:
#   assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称 ]}
#   fontdir: 新字体存放目录
# 将会返回以下:
#   newfont_name: { 原字体名 : [ 新字体绝对路径, 新字体名 ] }
def assFontSubset(assfont: dict, fontdir: str) -> dict:
    newfont_name = {}
    # print(fontdir)
    print('正在子集化……')
    for k in assfont.keys():
        # 偷懒没有变更该函数中的assfont解析到新的词典格式
        # 在这里会将assfont词典转换为旧的assfont列表形式
        # assfont: [ 字体绝对路径, 字体索引, 字符串, ASS内的字体名称 ]
        s = k.split('?') + [assfont[k][0], assfont[k][1]]
        subfontext = ''
        fontname, fontext = path.splitext(path.basename(s[0]))
        if fontext[1:].lower() in ['otc', 'ttc']:
            subfontext = fontext[:3] + 'f'
        else: subfontext = fontext
        #print(fontdir, path.exists(path.dirname(fontdir)), path.exists(fontdir))
        if path.exists(path.dirname(fontdir)):
            if not path.isdir(fontdir): 
                try:
                    os.mkdir(fontdir)
                except:
                    print('\033[1;31m[ERROR] 创建文件夹\"{0}\"失败\033[0m'.format(fontdir))
            if not path.isdir(fontdir): fontdir = path.dirname(fontdir)
        else: fontdir = os.getcwd()
        subfontdir = re.sub(cillegal, '_', s[3].upper())
        subfontpath = path.join(fontdir, subfontdir, fontname + subfontext)
        #print(fontdir, subfontpath)
        if not path.exists(path.dirname(subfontpath)): 
            try:
                os.mkdir(path.dirname(subfontpath))
            except:
                subfontpath = path.join(fontdir, fontname + subfontext)
                print('\033[1;31m[ERROR] 创建文件夹\"{0}\"失败\033[0m'.format(fontdir))
        subsetarg = [s[0], '--text={0}'.format(s[2]), '--output-file={0}'.format(subfontpath), '--font-number={0}'.format(s[1]), '--passthrough-tables']
        try:
            subset.main(subsetarg)
        except PermissionError:
            print('\033[1;31m[ERROR] 文件\"{0}\"访问失败[0m'.format(path.basename(subfontpath)))
        except:
            print('\033[1;31m[ERROR] {0}[0m'.format(sys.exc_info()))
            print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
            crcnewf = subfontpath
            shutil.copy(s[0], crcnewf)
            subfontcrc = None
            newfont_name[s[3]] = [crcnewf, subfontcrc]
            continue
        #os.system('pyftsubset {0}'.format(' '.join(subsetarg)))
        if path.exists(subfontpath): 
            subfontbyte = open(subfontpath, mode='rb')
            subfontcrc = str(hex(zlib.crc32(subfontbyte.read())))[2:].upper()
            if len(subfontcrc) < 8: subfontcrc = '0' + subfontcrc
            # print('CRC32: {0} \"{1}\"'.format(subfontcrc, path.basename(s[0])))
            subfontbyte.close()
            if fontext[1:].lower() in ['otc', 'ttc']:
                rawf = ttLib.TTCollection(s[0], lazy=True)[int(s[1])]
            else: rawf = ttLib.TTFont(s[0], lazy=True)
            newf = ttLib.TTFont(subfontpath, lazy=False)
            if len(newf['name'].names) == 0:
                for i in range(0,7):
                    if len(rawf['name'].names) - 1 >= i:
                        name = rawf['name'].names[i]
                        namestr = getNameStr(name, subfontcrc)
                        newf['name'].addName(namestr, minNameID=-1)
            else:
                for i in range(0, len(rawf['name'].names)):
                    name = rawf['name'].names[i]
                    nameID = name.nameID
                    platID = name.platformID
                    langID = name.langID
                    platEncID = name.platEncID
                    namestr = getNameStr(name, subfontcrc)
                    newf['name'].setName(namestr ,nameID, platID, platEncID, langID)
            if len(newf.getGlyphOrder()) == 1 and '.notdef' in newf.getGlyphOrder():
                print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
                crcnewf = subfontpath
                newf.close()
                if not subfontpath == s[0]: os.remove(subfontpath)
                shutil.copy(s[0], crcnewf)
                subfontcrc = None
            else:
                crcnewf = '.{0}'.format(subfontcrc).join(path.splitext(subfontpath))
                newf.save(crcnewf)
                newf.close()
            rawf.close()
            if path.exists(crcnewf):
                if not subfontpath == crcnewf: os.remove(subfontpath)
                newfont_name[s[3]] = [crcnewf, subfontcrc]
    #print(newfont_name)
    return newfont_name

# 更改ASS样式对应的字体
# 需要以下输入
#   fullass: 完整的ass文件内容，以行分割为列表
#   newfont_name: { 原字体名 : [ 新字体路径, 新字体名 ] }
#   asspath: 原ass文件的绝对路径
#   styleline: [V4/V4+ Styles]标签在SSA/ASS中的行数，对应到fullass列表的索引数
#   font_pos: Font参数在 Styles 的 Format 中位于第几个逗号之后
# 以下输入可选
#   outdir: 新字幕的输出目录，默认为源文件目录
#   ncover: 为True时不覆盖原有文件，为False时覆盖
# 将会返回以下
#   newasspath: 新ass文件的绝对路径
def assFontChange(fullass: list, newfont_name: dict, asspath: str, styleline: int, 
                        font_pos: int, outdir: str = '', ncover: bool = False) -> str:
    for i in range(styleline + 2, len(fullass)):
        if len(fullass[i].split(':')) < 2: 
            if re.search(style_read, '\n'.join(fullass[i + 1:])) is None: 
                break
            else:
                continue
        styleStr = ''.join(fullass[i].split(':')[1:]).strip(' ').split(',')
        fontstr = styleStr[font_pos].lstrip('@')
        if not newfont_name.get(fontstr) is None:
            if not newfont_name[fontstr][1] is None:
                fullass[i] = fullass[i].replace(fontstr, newfont_name[fontstr][1])
    if path.exists(path.dirname(outdir)):
        try:
            if not path.isdir(outdir):
                os.mkdir(outdir)
        except:
            print('\033[1;31m[ERROR] 创建文件夹\"{0}\"失败\033[0m'.format(outdir))
    if path.isdir(outdir): 
        newasspath = path.join(outdir, '.subset'.join(path.splitext(path.basename(asspath))))
    else: newasspath = '.subset'.join(path.splitext(asspath))
    if path.exists(newasspath) and ncover:
        testpathl = path.splitext(newasspath)
        testc = 1
        testpath = '{0}#{1}{2}'.format(testpathl[0], testc, testpathl[1])
        while path.exists(testpath):
            testc += 1
            testpath = '{0}#{1}{2}'.format(testpathl[0], testc, testpathl[1])
        newasspath = testpath
    ass = open(newasspath, mode='w', encoding='utf-8')
    ass.writelines(fullass)
    ass.close()
    #print('ASS样式转换完成: {0}'.format(path.basename(newasspath)))
    return newasspath

# ASFMKV，将媒体文件、字幕、字体封装到一个MKV文件，需要mkvmerge命令行支持
# 需要以下输入
#   file: 媒体文件绝对路径
#   outfile: 输出文件的绝对路径，如果该选项空缺，默认为 输入媒体文件.muxed.mkv
#   asslangs: 赋值给字幕轨道的语言，如果字幕轨道多于asslangs的项目数，超出部分将全部应用asslangs的末项
#   asspaths: 字幕绝对路径列表
#   fontpaths: 字体列表，格式为 [[字体1绝对路径], [字体1绝对路径], ...]，必须嵌套一层，因为主函数偷懒了
# 将会返回以下
#   mkvmr: mkvmerge命令行的返回值
def ASFMKV(file: str, outfile: str = '', asslangs: list = [], asspaths: list = [], fontpaths: list = []) -> int:
    #print(fontpaths)
    global rmAssIn, rmAttach, mkvout, notfont
    if file is None: return 4
    elif file == '': return 4
    elif not path.exists(file) or not path.isfile(file): return 4
    if outfile is None: outfile = ''
    if outfile == '' or not path.exists(path.dirname(outfile)): 
        outfile = '.muxed'.join([path.splitext(file)[0], '.mkv']) 
    outfile = path.splitext(outfile)[0] + '.mkv'
    if path.exists(outfile):
        checkloop = 1
        while path.exists('#{0}'.format(checkloop).join(path.splitext(outfile))):
            checkloop += 1
        outfile = '#{0}'.format(checkloop).join([path.splitext(outfile)[0], '.mkv'])
    mkvargs = []
    if rmAssIn: mkvargs.append('-S')
    if rmAttach: mkvargs.append('-M')
    mkvargs.extend(['(', file, ')'])
    fn = path.splitext(path.basename(file))[0]
    if len(asspaths) > 0:
        for i in range(0, len(asspaths)):
            s = asspaths[i]
            assfn = path.splitext(path.basename(s))[0]
            assnote = assfn[(assfn.find(fn) + len(fn)):].replace('.subset', '')
            #print(assfn, fn, assnote)
            if len(assnote) > 1:
                mkvargs.extend(['--track-name', '0:{0}'.format(assnote[1:])])
            if len(asslangs) > 0 and path.splitext(s)[1][1:].lower() not in ['idx']:
                mkvargs.append('--language')
                if i < len(asslangs):
                    mkvargs.append('0:{0}'.format(asslangs[i]))
                else:
                    mkvargs.append('0:{0}'.format(asslangs[len(asslangs) - 1]))
            mkvargs.extend(['(', s, ')'])
    if len(fontpaths) > 0:
        for s in fontpaths:
            fext = path.splitext(s[0])[1][1:].lower()
            if fext in ['ttf', 'ttc']:
                mkvargs.extend(['--attachment-mime-type', 'application/x-truetype-font'])
            elif fext in ['otf', 'otc']:
                mkvargs.extend(['--attachment-mime-type', 'application/vnd.ms-opentype'])
            mkvargs.extend(['--attach-file', s[0]])
    mkvargs.extend(['--title', fn])
    mkvjsonp = path.splitext(file)[0] + '.mkvmerge.json'
    mkvjson = open(mkvjsonp, mode='w', encoding='utf-8')
    json.dump(mkvargs, fp=mkvjson, sort_keys=True, indent=2, separators=(',', ': '))
    mkvjson.close()
    mkvmr = os.system('mkvmerge @\"{0}\" -o \"{1}\"'.format(mkvjsonp, outfile))
    if mkvmr > 1:
        print('\n\033[1;31m[ERROR] 检测到不正常的mkvmerge返回值，重定向输出...\033[0m')
        os.system('mkvmerge -r \"{0}\" @\"{1}\" -o NUL'.format('{0}.{1}.log'
            .format(path.splitext(file)[0], datetime.now().strftime('%Y-%m%d-%H%M-%S_%f')), mkvjsonp))
    elif not notfont:
        for p in asspaths:
            print('\033[1;32m成功: \"{0}\"\033[0m'.format(p))
            try:
                os.remove(p)
            except:
                print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(p))
        for f in fontpaths:
            print('\033[1;32m成功: \"{0}\"\033[0m'.format(f[0]))
            try:
                os.remove(f[0])
            except:
                print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(f[0]))
        try:
            os.remove(mkvjsonp)
        except:
            print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(mkvjsonp))
        print('\033[1;32m输出成功: \"{0}\"\033[0m'.format(outfile))
    else:
        print('\033[1;32m输出成功: \"{0}\"\033[0m'.format(outfile))
    return mkvmr

# 从输入的目录中获取媒体文件列表
# 需要以下输入
#   dir: 要搜索的目录
# 返回以下结果
#   medias: 多媒体文件列表
#       结构: [[ 文件名（无扩展名）, 绝对路径 ], ...]
def getMediaFilelist(dir: str) -> list:
    medias = []
    global v_subdir, extlist
    if path.isdir(dir):
        if v_subdir:
            for r,ds,fs in os.walk(dir):
                for f in fs:
                    if path.splitext(f)[1][1:].lower() in extlist:
                        medias.append([path.splitext(path.basename(f))[0], path.join(r, f)])
        else:
            for f in os.listdir(dir):
                if path.isfile(path.join(dir, f)):
                    if path.splitext(f)[1][1:].lower() in extlist:
                        medias.append([path.splitext(path.basename(f))[0], path.join(dir, f)])
        return medias

# 在目录中找到与媒体文件列表中的媒体文件对应的字幕
# 遵循以下原则
#   媒体文件在上级目录，则匹配子目录中的字幕；媒体文件的字幕只能在媒体文件的同一目录或子目录中，不能在上级目录和其他同级目录
# 需要以下输入
#   medias: 媒体文件列表，结构见 getMediaFilelist
#   cpath: 开始搜索的顶级目录
# 将会返回以下
#   media_ass: 媒体文件与字幕文件的对应词典
#       结构: { 媒体文件绝对路径 : [ 字幕1绝对路径, 字幕2绝对路径, ...] }
def getSubtitles(medias: list, cpath: str) -> dict:
    media_ass = {}
    global s_subdir, matchStrict
    if s_subdir:
        for r,ds,fs in os.walk(cpath):
            for f in [path.join(r, s) for s in fs if path.splitext(s)[1][1:].lower() in subext]:
                if '.subset' in path.basename(f): continue
                for l in medias:
                    vdir = path.dirname(l[1])
                    sdir = path.dirname(f)
                    sext = path.splitext(f)
                    if (l[0] in f and not matchStrict) or (l[0] == path.basename(f)[:len(l[0])] and matchStrict):
                        if((vdir in sdir and sdir not in vdir) or (vdir == sdir)):
                            if sext[1][:1].lower() == 'idx':
                                if not path.exists(sext[1] + '.sub'):
                                    continue
                            if media_ass.get(l[1]) is None:
                                media_ass[l[1]] = [f]
                            else: media_ass[l[1]].append(f)
    else:
        for f in [path.join(cpath, s) for s in os.listdir(cpath) if not path.isdir(s) and 
        path.splitext(s)[1][1:].lower() in subext]:
            # print(f, cpath)
            if '.subset' in path.basename(f): continue
            for l in medias:
                # print(path.basename(f)[len(l[0]):], l)
                sext = path.splitext(f)
                if (l[0] in f and not matchStrict) or (l[0] == path.basename(f)[:len(l[0])] and matchStrict):
                    if path.dirname(l[1]) == path.dirname(f):
                        if sext[1][:1].lower() == 'idx':
                            if not path.exists(sext[1] + '.sub'):
                                continue
                        if media_ass.get(l[1]) is None:
                            media_ass[l[1]] = [f]
                        else: media_ass[l[1]].append(f)
    return media_ass

# 主函数，负责调用各函数走完完整的处理流程
# 需要以下输入
#   font_name: 字体名称与字体路径对应词典，结构见 fontProgress
#   asspath: 字幕绝对路径列表
# 以下输入可选
#   outdir: 输出目录，格式 [ 字幕输出目录, 字体输出目录, 视频输出目录 ]，如果项数不足，则取最后一项；默认为 asspaths 中每项所在目录
#   mux: 不要封装，只运行到子集化完成
#   vpath: 视频路径，只在 mux = True 时生效
#   asslangs: 字幕语言列表，将会按照顺序赋给对应的字幕轨道，只在 mux = True 时生效
# 将会返回以下
#   newasspath: 列表，新生成的字幕文件绝对路径
#   newfont_name: 词典，{ 原字体名 : [ 新字体绝对路径, 新字体名 ] }
#   ??? : 数值，mkvmerge的返回值；如果 mux = False，返回-1
def main(font_name: dict, asspath: list, outdir: list = ['', '', ''], mux: bool = False, vpath: str = '', asslangs: list = []):
    outdir_temp = outdir[:3]
    outdir = ['', '', '']
    for i in range(0, len(outdir_temp)):
        s = outdir_temp[i]
        # print(s)
        if s is None:
            outdir[i] = ''
        elif s == '':
            outdir[i] = s
        else:
            try:
                if not path.isdir(s) and path.exists(path.dirname(s)):
                    os.mkdir(s)
                if path.isdir(s): outdir[i] = s
            except:
                print('\033[1;31m[ERROR] 创建输出文件夹错误\n[ERROR] {0}\033[0m'.format(sys.exc_info()))
                if '\\' in s:
                    outdir[i] = path.join(os.getcwd(), path.basename(s.rstrip('\\')))
                else: outdir[i] = path.join(os.getcwd(), s)
    # print(outdir)
    # os.system('pause')
    global notfont
    # multiAss 多ASS文件输入记录词典
    # 结构: { ASS文件绝对路径 : [ 完整ASS文件内容(fullass), 样式位置(styleline), 字体在样式行中的位置(font_pos) ]}
    multiAss = {}
    assfont = {}
    fontlist = {}
    newasspath = []
    fo = ''
    if not notfont:
        print('\n字体名称总数: {0}'.format(len(font_name.keys())))
        noass = False
        for i in range(0, len(asspath)):
            s = asspath[i]
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                multiAss[s] = [[], 0, 0]
            else:
                print('正在分析字幕文件: \"{0}\"'.format(path.basename(s)))
                fullass, fontlist, styleline, font_pos = assAnalyze(s, fontlist)
                multiAss[s] = [fullass, styleline, font_pos]
                assfont, font_name = checkAssFont(fontlist, font_name, assfont)
        sn = path.splitext(path.basename(asspath[0]))[0]
        fn = path.join(path.dirname(asspath[0]), 'Fonts')
        if outdir[1] == '': outdir[1] = fn
        if not path.isdir(outdir[1]):
            try:
                os.mkdir(outdir[1])
                fo = path.join(outdir[1], sn)
            except:
                fo = path.join(path.dirname(outdir[1]), sn)
        else:
            fo = path.join(outdir[1], sn)
        newfont_name = assFontSubset(assfont, fo)
        for s in asspath:
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                newasspath.append(s)
            elif len(multiAss[s][0]) == 0 or multiAss[s][1] == multiAss[s][2]: 
                continue
            else: newasspath.append(assFontChange(multiAss[s][0], newfont_name, s, multiAss[s][1], multiAss[s][2], outdir[0]))
    else:
        newasspath = asspath
        newfont_name = {}
    if mux: 
        if outdir[2] == '': outdir[2] = path.dirname(vpath)
        if not path.isdir(outdir[2]):
            try:
                os.mkdir(outdir[2])
            except:
                outdir[2] = path.dirname(outdir[2])
        mkvr = ASFMKV(vpath, path.join(outdir[2], path.splitext(path.basename(vpath))[0] + '.mkv'), 
            asslangs=asslangs, asspaths=newasspath, fontpaths=list(newfont_name.values()))
        if not notfont:
            for ap in newasspath:
                if path.exists(path.dirname(ap)):
                    try:
                        os.rmdir(path.dirname(ap))
                    except:
                        break
            for fp in newfont_name.keys():
                if path.exists(path.dirname(newfont_name[fp][0])):
                    try:
                        os.rmdir(path.dirname(newfont_name[fp][0]))
                    except:
                        continue
            if path.isdir(fo):
                try:
                    os.rmdir(fo)
                except:
                    pass
        return newasspath, newfont_name, mkvr
    else:
        return newasspath, newfont_name, -1

def cls():
    os.system('cls')


# 初始化字体列表 和 mkvmerge 相关参数
os.system('title ASFMKV Python Remake 1.00 ^| (c) 2022 yyfll ^| Apache-2.0')
fontlist, font_name = getFileList(fontin)
font_name = fontProgress(fontlist, font_name)
no_mkvm = False
no_cmdc = False
mkvmv = '\n\033[1;33m没有检测到 mkvmerge\033[0m'
if os.system('mkvmerge -V 1>nul 2>nul') > 0:
    no_mkvm = True
else: 
    print('\n\033[1;33m正在获取 mkvmerge 语言编码列表和支持格式列表...\033[0m')
    mkvmv = '\n' + os.popen('mkvmerge -V --ui-language en', mode='r').read().replace('\n', '')
    extget = re.compile(r'\[.*\]')
    langmkv = os.popen('mkvmerge --list-languages', mode='r')
    for s in langmkv.buffer.read().decode('utf-8').splitlines()[2:]:
        s = s.replace('\n', '').split('|')
        for si in range(1, len(s)):
            ss = s[si]
            if len(ss.strip(' ')) > 0:
                langlist.append(ss.strip(' '))
    langmkv.close()
    for s in os.popen('mkvmerge -l --ui-language en', mode='r').readlines()[1:]:
        for ss in re.search(r'\[.*\]', s).group().lstrip('[').rstrip(']').split(' '):
            extsupp.append(ss)
    extl_c = extlist
    extlist = []
    print('')
    for i in range(0, len(extl_c)):
        s = extl_c[i]
        if s in extsupp:
            extlist.append(s)
        else:
            print('\033[1;31m[WARNING] 您设定的媒体扩展名 {0} 无效，已从列表移除\033[0m'.format(s))
    if len(extlist) != len(extl_c):
        print('\n\033[1;33m当前的媒体扩展名列表: \"{0}\"\033[0m\n'.format(';'.join(extlist)))
        os.system('pause')
    del extl_c
    if len(sublang) > 0:
        print('')
        sublang_c = sublang
        sublang = []
        for i in range(0, len(sublang_c)):
            s = sublang_c[i]
            if s in langlist:
                sublang.append(s)
            else:
                print('\033[1;31m[WARNING] 您设定的语言编码 {0} 无效，已从列表移除\033[0m'.format(s))
        if len(sublang) != len(sublang_c):
            print('\n\033[1;33m当前的语言编码列表: \"{0}\"\033[0m\n'.format(';'.join(sublang)))
            os.system('pause')
        del sublang_c


def cListAssFont(font_name):
    global resultw, s_subdir, copyfont
    leave = True
    while leave:
        cls()
        print('''ASFMKV-ListAssFontPy
注意: 本程序由于设计原因，列出的是字体文件与其内部字体名的对照表

选择功能:
[L] 回到上级菜单
[A] 列出全部字体
[B] 检查并列出字幕所需字体

切换开关:
[1] 将结果写入文件: \033[1;33m{0}\033[0m
[2] 拷贝所需字体: \033[1;33m{1}\033[0m
[3] 搜索子目录(字幕): \033[1;33m{2}\033[0m
'''.format(resultw, copyfont, s_subdir))
        work = os.system('choice /M 请输入 /C AB123L')
        if work == 1:
            cls()
            wfilep = path.join(os.getcwd(), datetime.now().strftime('ASFMKV_FullFont_%Y-%m%d-%H%M-%S_%f.log'))
            if resultw:
                wfile = open(wfilep, mode='w', encoding='utf-8-sig')
            else: wfile = None
            fn = ''
            print('FontList', file=wfile)
            for s in font_name.keys():
                nfn = path.basename(font_name[s][0])
                if fn != nfn: 
                    if wfile is not None: 
                        print('>\n{0} <{1}'.format(nfn, s), end='', file=wfile)
                    else: print('>\033[0m\n\033[1;36m{0}\033[0m \033[1m<{1}'.format(nfn, s), end='')
                else:
                    print(', {0}'.format(s), end='', file=wfile)
                fn = nfn
            if wfile is not None: 
                print(wfilep)
                print('>', file=wfile)
                wfile.close()
            else: print('>\033[0m')
        elif work == 2:
            cls()
            cpath = ''
            directout = False
            while not path.exists(cpath) and not directout:
                directout = True
                cpath = input('请输入目录路径或字幕文件路径: ').strip('\"')
                if cpath == '' : print('没有输入，回到上级菜单')
                elif not path.exists(cpath): print('\033[1;31m[ERROR] 找不到路径: \"{0}\"\033[0m'.format(cpath))
                elif not path.isfile(cpath) and not path.isdir(cpath): print('\033[1;31m[ERROR] 输入的必须是文件或目录！: \"{0}\"\033[0m'.format(cpath))
                elif not path.isabs(cpath): print('\033[1;31m[ERROR] 路径必须为绝对路径！: \"{0}\"\033[0m'.format(cpath))
                elif path.isdir(cpath): directout = False
                elif not path.splitext(cpath)[1][1:].lower() in ['ass', 'ssa']: 
                    print('\033[1;31m[ERROR] 输入的必须是ASS/SSA字幕文件！: \"{0}\"\033[0m'.format(cpath))
                else: directout = False
            #clist = []
            fontlist = {}
            assfont = {}
            if not directout:
                if path.isdir(cpath):
                    #print(cpath)
                    dir = ''
                    if s_subdir:
                        for r,ds,fs in os.walk(cpath):
                            for f in fs:
                                if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                                    a, fontlist, b, c = assAnalyze(path.join(r, f), fontlist, onlycheck=True)
                                    assfont, font_name = checkAssFont(fontlist, font_name, assfont, onlycheck=True)
                    else:
                        for f in os.listdir(cpath):
                            if path.isfile(path.join(cpath, f)):
                                #print(f, path.splitext(f)[1][1:].lower())
                                if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                                    #print(f, 'pass')
                                    a, fontlist, b, c = assAnalyze(path.join(cpath, f), fontlist, onlycheck=True)
                                    assfont, font_name = checkAssFont(fontlist, font_name, assfont, onlycheck=True)
                    fd = path.join(cpath, 'Fonts')
                else:
                    a, fontlist, b, c = assAnalyze(cpath, fontlist, onlycheck=True)
                    assfont, font_name = checkAssFont(fontlist, font_name, assfont, onlycheck=True)
                    fd = path.join(path.dirname(cpath), 'Fonts')
                if len(assfont.keys()) < 1:
                    print('\033[1;31m[ERROR] 目标路径没有ASS/SSA字幕文件\033[0m')
                else:
                    wfile = None
                    print('')
                    if copyfont or resultw: 
                        if not path.isdir(fd): os.mkdir(fd)
                    if resultw: 
                        wfile = open(path.join(cpath, 'Fonts', 'fonts.txt'), mode='w', encoding='utf-8-sig')
                    for s in assfont.keys():
                        ssp = s.split('?')
                        if not ssp[0] == ssp[1]:
                            fp = ssp[0]
                            fn = path.basename(fp)
                            ann = ''
                            errshow = False
                            if copyfont: 
                                try:
                                    shutil.copy(fp, path.join(fd, fn))
                                    ann = ' - copied'
                                except:
                                    print('[ERROR]', sys.exc_info())
                                    ann = ' - copy error'
                                    errshow = True
                            if resultw:
                                print('{0} <{1}>{2}'.format(assfont[s][1], path.basename(fn), ann), file=wfile)
                            if errshow:
                                print('\033[1;36m{0}\033[0m \033[1m<{1}>\033[1;31m{2}\033[0m'.format(assfont[s][1], path.basename(fn), ann))
                            else: print('\033[1;36m{0}\033[0m \033[1m<{1}>\033[1;32m{2}\033[0m'.format(assfont[s][1], path.basename(fn), ann))
                        else:
                            if resultw:
                                print('{0} - No Found'.format(ssp[0]), file=wfile)
                            print('\033[1;36m{0}\033[1;31m - No Found\033[0m'.format(ssp[0]))
                    if resultw: wfile.close()
                    print('')
        elif work == 3:
            if resultw: resultw = False 
            else: resultw = True
        elif work == 4:
            if copyfont: copyfont = False
            else: copyfont = True
        elif work == 5:
            if s_subdir: s_subdir = False
            else: s_subdir = True
        else:
            leave = False
        if work < 3: os.system('pause')


def checkOutPath(op: str, default: str) -> str:
    if op == '': return default
    if op[0] == '?': return '?' + re.sub(cillegal, '_', op[1:])
    if not path.isabs(op): 
        print('\033[1;31m[ERROR] 输入的必须是绝对路径或子目录名称！: \"{0}\"\033[0m'.format(op))
        os.system('pause')
        return default
    if path.isdir(op): return op
    if path.isfile(op): return path.dirname(op)
    print('\033[1;31m[ERROR] 输入的必须是目录路径或子目录名称！: \"{0}\"\033[0m'.format(op))
    os.system('pause')
    return default

def cFontSubset(font_name):
    global extlist, v_subdir, s_subdir, rmAssIn, rmAttach, \
    mkvout, assout, fontout, matchStrict, no_mkvm, notfont
    leave = True
    while leave:
        cls()
        print('''ASFMKV & ASFMKV-FontSubset

选择功能:
[L] 回到上级菜单
[A] 子集化字体
[B] 子集化并封装

切换开关:
[1] 检视媒体扩展名列表 及 语言编码列表
[2] 搜索子目录(视频): \033[1;33m{0}\033[0m
[3] 搜索子目录(字幕): \033[1;33m{1}\033[0m
[4] (封装)移除内挂字幕: \033[1;33m{2}\033[0m
[5] (封装)移除原有附件: \033[1;33m{3}\033[0m
[6] (封装)不封装字体: \033[1;33m{8}\033[0m
[7] 严格字幕匹配: \033[1;33m{7}\033[0m
[8] 媒体文件输出文件夹: \033[1;33m{4}\033[0m
[9] 字幕文件输出文件夹: \033[1;33m{5}\033[0m
[0] 字体文件输出文件夹: \033[1;33m{6}\033[0m
'''.format(v_subdir, s_subdir, rmAssIn, rmAttach, mkvout, assout, 
fontout, matchStrict, notfont))
        work = 0
        work = os.system('choice /M 请输入 /C AB1234567890L')
        if work == 2 and no_mkvm:
            print('[ERROR] 在您的系统中找不到 mkvmerge, 该功能不可用')
            work = -1
        if work in [1, 2]:
            cls()
            if work == 1: print('''子集化字体

搜索子目录(视频): \033[1;33m{0}\033[0m
搜索子目录(字幕): \033[1;33m{1}\033[0m
严格字幕匹配: \033[1;33m{2}\033[0m
字幕文件输出文件夹: \033[1;33m{3}\033[0m
字体文件输出文件夹: \033[1;33m{4}\033[0m
'''.format(v_subdir, s_subdir, matchStrict, assout, fontout))
            else: print('''子集化字体并封装

搜索子目录(视频): \033[1;33m{4}\033[0m
搜索子目录(字幕): \033[1;33m{5}\033[0m
移除内挂字幕: \033[1;33m{0}\033[0m
移除原有附件: \033[1;33m{1}\033[0m
不封装字体: \033[1;33m{2}\033[0m
严格字幕匹配: \033[1;33m{6}\033[0m
媒体文件输出文件夹: \033[1;33m{3}\033[0m
'''.format(rmAssIn, rmAttach, notfont, mkvout, v_subdir, s_subdir, matchStrict))
            cpath = ''
            directout = False
            while not path.exists(cpath) and not directout:
                directout = True
                cpath = input('不输入任何值 直接回车回到上一页面\n请输入文件或目录路径: ').strip('\"')
                if cpath == '':  print('没有输入，回到上级菜单')
                elif not path.isabs(cpath): print('\033[1;31m[ERROR] 输入的必须是绝对路径！: \"{0}\"\033[0m'.format(cpath))
                elif not path.exists(cpath): print('\033[1;31m[ERROR] 找不到路径: \"{0}\"\033[0m'.format(cpath))
                elif path.isfile(cpath): 
                    if path.splitext(cpath)[1][1:] not in extlist:
                        print('\033[1;31m[ERROR] 扩展名不正确: \"{0}\"\033[0m'.format(cpath))
                    else: directout = False
                elif not path.isdir(cpath):
                    print('\033[1;31m[ERROR] 输入的应该是目录或媒体文件！: \"{0}\"\033[0m'.format(cpath))
                else: directout = False
            # print(directout)
            if not directout:
                if path.isfile(cpath): 
                    medias = [[path.splitext(path.basename(cpath))[0], cpath]]
                    cpath = path.dirname(cpath)
                else: medias = getMediaFilelist(cpath)
                # print(medias)
                if not medias is None:
                    media_ass = getSubtitles(medias, cpath)
                    # print(media_ass)
                    for k in media_ass.keys():
                        #print(k)
                        if assout == '':
                            assout_cache = path.join(cpath, 'Subtitles')
                        elif assout[0] == '?':
                            assout_cache = path.join(cpath, assout[1:])
                        else: assout_cache = assout
                        if fontout == '':
                            fontout_cache = path.join(cpath, 'Fonts')
                        elif fontout[0] == '?':
                            fontout_cache = path.join(cpath, fontout[1:])
                        else: fontout_cache = fontout
                        if mkvout == '':
                            mkvout_cache = ''
                        elif mkvout[0] == '?':
                            mkvout_cache = path.join(cpath, mkvout[1:])
                        else: mkvout_cache = mkvout

                        #print([assout_cache, fontout_cache, mkvout_cache])
                        if work == 1:
                            newasspaths, newfont_name, mkvr = main(font_name, media_ass[k], 
                            mux = False, outdir=[assout_cache, fontout_cache, mkvout_cache])
                        else:
                            newasspaths, newfont_name, mkvr = main(font_name, media_ass[k], 
                            mux = True, outdir=[assout_cache, fontout_cache, mkvout_cache], vpath=k)

                        for ap in newasspaths:
                            if path.exists(ap): 
                                print('\033[1;32m成功:\033[1m \"{0}\"\033[0m'.format(path.basename(ap)))
                        for nf in newfont_name.keys():
                            if path.exists(newfont_name[nf][0]):
                                if newfont_name[nf][1] is None:
                                    print('\033[1;31m失败:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\"\033[0m'.format(path.basename(nf), path.basename(newfont_name[nf][0])))
                                else:
                                    print('\033[1;32m成功:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\" ({2})\033[0m'.format(path.basename(nf), path.basename(newfont_name[nf][0]), newfont_name[nf][1]))
        elif work == 3:
            cls()
            print('ExtList Viewer 1.00-Final\n')
            for i in range(0, len(extlist)):
                s = extlist[i]
                print('[Ext{0:>3d}] {1}'.format(i, s))
            print('\n')
            if not no_mkvm:
                os.system('pause')
                cls()
                os.system('mkvmerge --list-languages')
                #print('\033[1m mkvmerge 语言编码列表 \033[0m')
            else:
                print('没有检测到mkvmerge，无法输出语言编码列表')
        elif work == 4:
            if v_subdir: v_subdir = False
            else: v_subdir = True
        elif work == 5:
            if s_subdir: s_subdir = False
            else: s_subdir = True
        elif work == 6:
            if rmAssIn: rmAssIn = False
            else: rmAssIn = True
        elif work == 7:
            if rmAttach: rmAttach = False
            else: rmAttach = True
        elif work == 8:
            if notfont: notfont = False
            else: notfont = True
        elif work == 9:
            if matchStrict: matchStrict = False
            else: matchStrict = True
        elif work in [10, 11, 12]:
            cls()
            print('''请输入目标目录路径或子目录名称\n(若要输入子目录，请在目录名前加\"?\")
(不支持多层子目录，会自动将\"\\\"换成下划线)''')
            if work == 10:
                mkvout = checkOutPath(input(), mkvout)
            elif work == 11:
                assout = checkOutPath(input(), assout)
            elif work == 12:
                fontout = checkOutPath(input(), fontout)
        else:
            leave = False
        if work < 4:
            os.system('pause')

def cLicense():
    cls()
    print('''AddSubFontMKV Python Remake 1.00
Apache-2.0 License
Copyright(c) 2022 yyfll

依赖:
fontTools    |  MIT License
chardet      |  LGPL-2.1 License
mkvmerge     |  GPL-2 License
''')
    print('for more information:\nhttps://www.apache.org/licenses/')
    os.system('pause')

cls()
        
if os.system('choice /? 1>nul 2>nul') > 0:
    no_cmdc = True

if __name__=="__main__":
    while 1:
        work = 0
        print('''ASFMKV Python Remake 1.00 | (c) 2022 yyfll{0}
字体名称数: [\033[1;33m{2}\033[0m]（包含乱码的名称）

请选择功能:
[A] ListAssFont
[B] 字体子集化 & MKV封装
[C] 检视重复字体: 重复名称[\033[1;33m{1}\033[0m]

其他:
[D] 依赖与许可证
[L] 直接退出'''.format(mkvmv, len(dupfont.keys()), len(font_name.keys())))
        print('')
        work = os.system('choice /M 请选择: /C ABCDL')
        if work == 1:
            cListAssFont(font_name)
        elif work == 2:
            cFontSubset(font_name)
        elif work == 3:
            cls()
            if len(dupfont.keys()) > 0:
                for s in dupfont.keys():
                    print('\033[1;31m[{0}]\033[0m'.format(s))
                    for ss in dupfont[s]:
                        print('\"{0}\"'.format(ss))
                    print('')
            else:
                print('没有名称重复的字体')
            os.system('pause')
        elif work == 4:
            cLicense()
        elif work == 5:
            exit()
        else: exit()
        cls()