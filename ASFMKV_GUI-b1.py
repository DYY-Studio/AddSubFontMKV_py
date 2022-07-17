# -*- coding: UTF-8 -*-
# *************************************************************************
#
# 请使用支持 UTF-8 NoBOM 并最好带有 Python 语法高亮的文本编辑器
# Windows 7 的用户请最好不要使用 记事本 打开本脚本
#
# *************************************************************************

# 调用库，请不要修改
import shutil
from traceback import print_tb
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QColor
from PyQt5.QtCore import QRect
from typing import Optional, Any
from fontTools import ttLib
from fontTools import subset
from chardet.universaldetector import UniversalDetector
import os
from os import path
import sys
import re
import winreg
import zlib
import json
import copy
from colorama import init
from datetime import datetime

# 初始化环境变量
# *************************************************************************
# 自定义变量
# 修改注意: Python的布尔类型首字母要大写 True 或 False；在名称中有单引号的，需要输入反斜杠转义如 \'
# *************************************************************************
# extlist 可输入的视频媒体文件扩展名
extlist = 'mkv;mp4;mts;mpg;flv;mpeg;m2ts;avi;webm;rm;rmvb;mov;mk3d;vob'
# *************************************************************************
# no_extcheck 关闭对extlist的扩展名检查来添加一些可能支持的格式
no_extcheck = False
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
# fontload 从视频所在文件夹及其子文件夹临时载入字体（优先于系统字体）
#   True    载入
#   False   不载入
fontload = True
# *************************************************************************
# o_fontload 只从视频所在文件夹及其子文件夹临时载入字体（不再读取系统字体）
# 可以提高启动速度，如果 o_fontload = True 且 fontload = False，那么 fontload 会被自动设定为 True
#   True    不再载入系统字体
#   False   仍然载入系统字体
o_fontload = False
# *************************************************************************
# s_fontload 是否从视频所在文件夹的子文件夹载入字体
#   True    载入
#   False   不载入（次一级目录中名称带有font的除外）
s_fontload = False
# *************************************************************************
# f_priority 字体优先级，主要处理同名字体的取舍方式
# 注意：在开启 fontload 时工作目录中的字体始终是最优先使用
#   True    优先使用自定义目录及要求用户输入的目录中的字体
#   False   优先使用系统注册表中的字体
f_priority = False
# *************************************************************************
# notfont (封装)字体嵌入
#   True    不嵌入字体，不子集化字体，不替换字幕中的字体信息
#   False   始终嵌入字体
notfont = False
# *************************************************************************
# sublang (封装)字幕语言
#   会按照您所输入的顺序给字幕赋予语言编码，如果字幕数多于语言数，多出部分将赋予最后一种语言编码
#   IDX+SUB 的 DVDSUB 由于 IDX 文件一般有语言信息，对于DVDSUB不再添加语言编码
#   各语言之间应使用半角分号 ; 隔开，如 'chi;chi;chi;und'
#   可以在 mkvmerge -l 了解语言编码
sublang = ''
# *************************************************************************
# matchStrict 严格匹配
#   True   媒体文件名必须在字幕文件名的最前方，如'test.mkv'的字幕可以是'test.ass'或是'test.sub.ass'，但不能是'sub.test.ass'
#   False  只要字幕文件名中有媒体文件名就行了，不管它在哪
matchStrict = True
# *************************************************************************
# warningStop 对有可能子集化失败的字体的处理方式
#   True   不子集化
#   False  子集化（如果失败，会保留原始字体；如果原始字体是TTC/OTC，则会保留其中的TTF/OTF）
warningStop = False
# errorStop 子集化失败后的处理方式
#   True   终止批量子集化
#   False  继续运行（运行方式参考warningStop）
errorStop = True
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
# lcidfil
# LCID过滤器，用于选择字体名称所用语言
# 结构为：{ platformID(Decimal) : { LCID : textcoding } }
# 目前Textcoding没有实际作用，仅用于让这个词典可读性更强
# 详情可在 https://docs.microsoft.com/en-us/typography/opentype/spec/name#platform-encoding-and-language-ids 查询
lcidfil = {3: {
    2052: 'gbk',
    1042: 'euc-kr',
    1041: 'shift-jis',
    1033: 'utf-16be',
    1028: 'big5',
    0: 'utf-16be'
}, 2: {
    0: 'ascii',
    1: 'utf-8',
    2: 'iso-8859-1'
}, 1: {
    33: 'gbk',
    23: 'euc-kr',
    19: 'big5',
    11: 'shift-jis',
    0: 'mac-roman'
}, 0: {
    0: 'utf-16be',
    1: 'utf-16be',
    2: 'utf-16be',
    3: 'utf-16be',
    4: 'utf-16be',
    5: 'utf-16be'
}}

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
if o_fontload and not fontload:
    fontload = True
langlist = []
extsupp = []
dupfont = {}
init()

# font_info 列表结构
#   [ font_name, font_n_lower, font_family, warning_font ]
# font_name 词典结构
#   { 字体名称 : [ 字体绝对路径 , 字体索引 (仅用于TTC/OTC; 如果是TTF/OTF，默认为0) , 字体样式 ] }
# dupfont 词典结构
#   { 重复字体名称 : [ 字体1绝对路径, 字体2绝对路径, ... ] }


def fontlistAdd(s: str, fn: str, fontlist: dict) -> dict:
    if len(s) > 0:
        si = 0
        fn = fn.lstrip('@')
        # print(fn, s)
        if fontlist.get(fn) is None:
            fontlist[fn] = s[0]
            si += 1
        for i in range(si, len(s)):
            if not s[i] in fontlist[fn]:
                fontlist[fn] = fontlist[fn] + s[i]
    return fontlist


def updateAssFont(assfont: dict = {}, assfont2: dict = {}) -> Optional[dict]:
    if assfont2 is None:
        return None
    for key in assfont2.keys():
        if assfont.get(key) is not None:
            assfont[key] = [
                assfont[key][0] + ''.join([ks for ks in assfont2[key][0] if ks not in assfont[key][0]]),
                assfont[key][1] + ''.join(
                    ['|' + ks for ks in assfont2[key][1].split('|') if ks.lower() not in assfont[key][1].lower()]),
                assfont[key][2] + ''.join(
                    ['|' + ks for ks in assfont2[key][2].split('|') if ks.lower() not in assfont[key][2].lower()])
            ]
        else:
            assfont[key] = assfont2[key]
    del assfont2
    return assfont


def assAnalyze(asspath: str, fontlist: dict = {}, onlycheck: bool = False):
    """
ASS分析部分

需要输入
  asspath: ASS文件的绝对路径
  font_name: 字体名称与字体绝对路径词典，用于查询 Bold、Italic 字体

可选输入
  fontlist: 可以更新fontlist，用于多ASS同时输入的情况，结构见下
  onlycheck: 只确认字幕中的字体，仅仅返回fontlist

将会返回
  fullass: 完整的ASS文件内容，以行分割
  fontlist: 字体与其所需字符 { 字体 : 字符串 }
  styleline: 样式内容的起始行
  font_pos: 字体在样式中的位置
  fn_lines: 带有fn标签的行数与该行的完整特效标签，一项一个 [ [行数, { 标签1 : ( ASS内部字体名称, Italic, Bold )}], ... ]
    """
    global style_read
    # 初始化变量
    eventline = 0
    style_pos = 0
    style_pos2 = 0
    text_pos = 0
    font_pos = 0
    bold_pos = 0
    italic_pos = 0
    infoline = 0
    styleline = 0
    ssRecover = {}
    fn_lines = []
    # 编译分析用正则表达式
    style = re.compile(r'^\[[Vv]4.*[Ss]tyles\]$')
    event = re.compile(r'^\[[Ee]vents\]$')
    sinfo = re.compile(r'^\[[Ss]cript [Ii]nfo\]$')
    # 识别文本编码并读取整个SubtitleStationAlpha文件到内存
    print('\033[1;33m正在分析字幕: \033[1;37m\"{0}\"\033[0m'.format(path.basename(asspath)))
    ass = open(asspath, mode='rb')
    # 使用 UniversalDetector 确定字幕文本编码（防止非标准 UTF-16LE 字幕）
    detector = UniversalDetector()
    for dt in ass:
        detector.feed(dt)
        if detector.done:
            ass_code = detector.result['encoding']
            break
    detector.reset()
    ass.close()
    ass = open(asspath, encoding=ass_code, mode='r')
    fullass = ass.readlines()
    ass.close()
    asslen = len(fullass)
    # 在文件中搜索Styles标签、Events标签、Script Info标签来确认起始行
    for s in range(0, asslen):
        if re.match(sinfo, fullass[s]) is not None:
            infoline = s
        if re.match(style, fullass[s]) is not None:
            styleline = s
        elif re.match(event, fullass[s]) is not None:
            eventline = s
        if styleline != 0 and eventline != 0:
            break

    # 子集化字体信息还原
    # 如果在 [Script Info] 中找到字体子集化的后缀信息，则将字体名称还原并移除后缀
    ssRemove = []
    for s in range(infoline + 1, asslen):
        ss = fullass[s]
        if ss.strip(' ') == '':
            continue
        elif ss.strip(' ')[0] == ';' and ss.lower().find('font subset') > -1 and len(ss.split(':')) > 1:
            fr = ss.split(':')[1].split('-')
            rawname = '-'.join(fr[1:]).strip(' ').rstrip('\n')
            nowname = fr[0].strip(' ').rstrip('\n')
            if not rawname == nowname: ssRecover[nowname] = rawname
            ssRemove.append(s)
        elif ss.strip(' ')[0] == '[' and ss.strip(' ')[-1] == ']':
            break
        else:
            continue

    if len(ssRemove) > 0:
        ssRemove.reverse()
        for s in ssRemove:
            fullass.pop(s)
            asslen -= 1
            infoline -= 1
            styleline -= 1
            eventline -= 1
    del ssRemove

    # 获取Style的 Format 行，并用半角逗号分割
    style_format = ''.join(fullass[styleline + 1].split(':')[1:]).strip(' ').split(',')
    # 确定Style中 Name 、 Fontname 、 Bold 、 Italic 的位置
    style_pos, font_pos, bold_pos, italic_pos = -1, -1, -1, -1
    for i in range(0, len(style_format)):
        s = style_format[i].lower().strip(' ').replace('\n', '')
        if s == 'name':
            style_pos = i
        elif s == 'fontname':
            font_pos = i
        elif s == 'bold':
            bold_pos = i
        elif s == 'italic':
            italic_pos = i
        if style_pos != -1 and font_pos != -1 and bold_pos != -1 and italic_pos != -1:
            break
    if style_pos == font_pos == bold_pos == italic_pos == -1:
        style_pos, font_pos, bold_pos, italic_pos = 0, 0, 0, 0

    # 获取 字体表 与 样式字体对应表
    style_font = {}
    # style_font 词典内容:
    #   { 样式 : 字体名 }
    # fontlist 词典内容:
    #   { 字体名?斜体?粗体 : 使用该字体的文本 }
    for i in range(styleline + 2, asslen):
        # 如果该行不是样式行，则跳过，直到文件末尾或出现其他标记
        if len(fullass[i].split(':')) < 2:
            continue
        if fullass[i].split(':')[0].strip(' ').lower() != 'style':
            break
        styleStr = ''.join([s.strip(' ') for s in fullass[i].split(':')[1:]]).strip(' ').split(',')
        font_key = styleStr[font_pos].lstrip('@')
        # 如果有字体还原信息，则将样式行对应的字体还原
        if len(ssRecover) > 0:
            if ssRecover.get(font_key) is not None:
                font_key = ssRecover[font_key]
                styleStr[font_pos] = font_key
                fullass[i] = fullass[i].split(':')[0] + ': ' + ','.join(styleStr)
        # 获取样式行中的粗体斜体信息
        isItalic = - int(styleStr[italic_pos])
        isBold = - int(styleStr[bold_pos])
        fontlist.setdefault('{0}?{1}?{2}'.format(font_key, isItalic, isBold), '')
        style_font[styleStr[style_pos]] = '{0}?{1}?{2}'.format(font_key, isItalic, isBold)

    # print(fontlist)

    # 提取Event的 Format 行，并用半角逗号分割
    event_format = ''.join(fullass[eventline + 1].split(':')[1:]).strip(' ').split(',')
    style_pos2, text_pos = -1, -1
    # 确定Event中 Style 和 Text 的位置
    for i in range(0, len(event_format)):
        if event_format[i].lower().replace('\n', '').strip(' ') == 'style':
            style_pos2 = i
        elif event_format[i].lower().replace('\n', '').strip(' ') == 'text':
            text_pos = i
        if style_pos2 != -1 and text_pos != -1:
            break
    if style_pos2 == -1 and text_pos == -1:
        style_pos2, text_pos == 0, 0

    # 获取 字体的字符集
    # 先获取 Style，用style_font词典查找对应的 Font
    # 再将字符串追加到 fontlist 中对应 Font 的值中
    for i in range(eventline + 2, asslen):
        eventline_sp = fullass[i].split(':')
        if len(eventline_sp) < 2:
            continue
        # print(fullass[i])
        if eventline_sp[0].strip(' ').lower() == 'comment':
            continue
        elif eventline_sp[0].strip(' ').lower() != 'dialogue':
            break
        eventline_sp = ''.join(eventline_sp[1:]).split(',')
        eventftext = ','.join(eventline_sp[text_pos:])
        effectDel = r'(\{.*?\})|(\\[hnN])|(\s)'
        textremain = ''
        # 矢量绘图处理，如果发现有矢量表达，从字符串中删除这一部分
        # 这一部分的工作未经详细验证，作用也不大
        if re.search(r'\{.*?\\p[1-9]\d*.*?\}([\s\S]*?)\{.*?\\p0.*?\}', eventftext) is not None:
            vecpos = re.findall(r'\{.*?\\p[1-9]\d*.*?\}[\s\S]*?\{.*?\\p0.*?\}', eventftext)
            nexts = 0
            for s in vecpos:
                vecfind = eventftext.find(s)
                textremain += eventftext[nexts:vecfind]
                nexts = vecfind
                s = re.sub(r'\\p\d+', '', re.sub(r'}.*?{', '}{', s))
                textremain += s
        elif re.search(r'\{.*?\\p[1-9]\d*.*?\}', eventftext) is not None:
            eventftext = re.sub(r'\\p\d+', '',
                                eventftext[:re.search(r'\{.*?\\p[1-9]\d*.*?\}', eventftext).span()[0]])
        if len(textremain) > 0:
            eventftext = textremain
        eventfont = style_font.get(eventline_sp[style_pos2].lstrip('*'))

        # 粗体、斜体标签处理
        splittext = []
        splitpos = []
        # 首先查找蕴含有启用粗体/斜体标记的特效标签
        if re.search(r'\{.*?(?:\\b[7-9]00|\\b1|\\i1).*?\}', eventftext) is not None:
            lastfind = 0
            allfind = re.findall(r'\{.*?\}', eventftext)
            # 在所有特效标签中寻找
            # 启用粗体斜体
            # 禁用粗体斜体
            # 分成两种情况再进入下一层嵌套if
            # 启用粗体/启用斜体
            # 禁用粗体/禁用斜体
            # 然后分别确认该特效标签的适用范围，以准确将字体子集化
            for sti in range(0, len(allfind)):
                st = allfind[sti]
                ibclose = re.search(r'(\\b[1-4]00|\\b0|\\i0|\\b[\\\}]|\\i[\\\}])', st)
                ibopen = re.search(r'(\\b[7-9]00|\\b1|\\i1)', st)
                if ibopen is not None:
                    stfind = eventftext.find(st)
                    for stii in range(sti, -1, -1):
                        if len(splitpos) > 0:
                            if splitpos[-1][1] >= eventftext.find(allfind[stii]):
                                break
                        if allfind[stii].find('\\fn') > -1:
                            stfind = eventftext.find(allfind[stii])
                            break
                    addbold = '0'
                    additalic = '0'
                    if re.search(r'(\\b[7-9]00|\\b1)', st) is not None:
                        if re.search(r'(\\b[7-9]00|\\b1)', st).span()[0] > max(
                                [st.find('\\b0'), st.find('\\b\\'), st.find('\\b}')]):
                            addbold = '1'
                    if st.find('\\i1') > st.find('\\i0'): additalic = '1'
                    if len(splittext) == 0:
                        if stfind > 0:
                            splittext.append([eventftext[:stfind], '0', '0'])
                        splittext.append([eventftext[stfind:], additalic, addbold])
                    else:
                        if splittext[-1][1] != additalic or splittext[-1][2] != addbold:
                            if stfind > 0:
                                splittext[-1][0] = eventftext[lastfind:stfind]
                            splittext.append([eventftext[stfind:], additalic, addbold])
                    lastfind = stfind
                elif ibclose is not None:
                    stfind = eventftext.find(st)
                    for stii in range(sti, -1, -1):
                        if len(splitpos) > 0:
                            if splitpos[-1][1] >= eventftext.find(allfind[stii]):
                                break
                        if allfind[stii].find('\\fn') > -1:
                            stfind = eventftext.find(allfind[stii])
                            break
                    if len(splittext) > 0:
                        ltext = splittext[-1]
                        readytext = [eventftext[stfind:], ltext[1], ltext[2]]
                        if re.search(r'(\\i0|\\i[\\\}])', st) is not None:
                            readytext[1] = '0'
                        if re.search(r'(\\b[1-4]00|\\b0|\\b[\\\}])', st) is not None:
                            readytext[2] = '0'
                        if ltext[1] != readytext[1] or ltext[2] != readytext[2]:
                            if stfind > 0:
                                splittext[-1][0] = eventftext[lastfind:stfind]
                            splittext.append(readytext)
                            lastfind = stfind
                    else:
                        splittext.append([eventftext, '0', '0'])
                else:
                    continue
        if len(splittext) == 0:
            splittext.append([eventftext, '0', '0'])
        # elif len(splittext) > 1:
        #     print(splittext)
        splitpos = [[0, len(splittext[0][0]), int(splittext[0][1]), int(splittext[0][2])]]
        if len(splittext) > 1:
            for spi in range(1, len(splittext)):
                spil = len(splittext[spi][0])
                lspil = splitpos[spi - 1][1]
                splitpos.append([lspil, spil + lspil, int(splittext[spi][1]), int(splittext[spi][2])])
            # print(splitpos)
            # print(eventftext)
            # os.system('pause')
        del splittext
        # print(eventftext)
        # 字体标签分析
        eventftext2 = eventftext
        # 在全部特效标签中寻找 fn 字体标签
        itl = [its for its in re.findall(r'\{.*?\}', eventftext) if its.find('\\fn') > -1]
        # 如果有发现 fn 字体标签，利用 while 循环确定该标签的适用范围和所包含的文本
        if len(itl) > 0:
            fn_line = [i]
            newstrl = []
            # it = Index Font
            while len(itl) > 0:
                if len(newstrl) == 0:
                    newstrl.append([eventftext, 0, ''])
                itpos = eventftext.find(itl[0])
                it = (itpos, itpos + len(itl[0]))
                s = eventftext[(it[0] + 1):(it[1] - 1)]
                cuttext = eventftext[it[0]:]
                fnpos = eventftext2.find(cuttext)
                itl = itl[1:]
                if len(itl) > 0:
                    itpos = eventftext.find(itl[0])
                    newstrl[-1][0] = newstrl[-1][0][:itpos]
                else:
                    newstrl[-1][0] = newstrl[-1][0][:it[0]]
                newstrl.append([cuttext, fnpos, s])
                eventftext = eventftext[it[0]:]
            newstrl = [nst for nst in newstrl if len(nst[0].strip(' ')) > 0]
            # 然后将新增的字体所对应的文本添加到fn_line中
            fn = ''
            for sp in splitpos:
                for fi in range(0, len(newstrl)):
                    fs = newstrl[fi]
                    s = fs[2]
                    if int(fs[1]) in range(sp[0], sp[1]):
                        fss = re.sub(effectDel, '', fs[0])
                        li = [sf.strip(' ') for sf in s.split('\\') if len(s.strip(' ')) > 0]
                        li.reverse()
                        for sf in li:
                            if 'fn' in sf.lower():
                                fn = sf[2:].strip(' ').lstrip('@')
                                if len(ssRecover) > 0:
                                    if ssRecover.get(fn) is not None:
                                        fullassSplit = fullass[i].split(',')
                                        fullass[i] = ','.join(fullassSplit[0:text_pos] + [
                                            ','.join(fullassSplit[text_pos:]).replace(fn, ssRecover[fn])])
                                        s = s.replace(fn, ssRecover[fn])
                                        fn = ssRecover[fn]
                                fn_line.append({s: (fn, str(sp[2]), str(sp[3]))})
                                break
                        if len(fn) == 0:
                            fontlist = fontlistAdd(fss, eventfont, fontlist)
                        else:
                            fontlist = fontlistAdd(fss, '?'.join([fn, str(sp[2]), str(sp[3])]), fontlist)
            # s = re.sub(effectDel, '', eventftext[:it[0]])
            # print('add', s)
            # print('fn', fn)
            # print('ef', eventftext)
            # it = re.search(r'\{\\fn.*?\}', eventftext)
            # os.system('pause')
            fn_lines.append(fn_line)
            del fn_line
        else:
            if eventfont is not None:
                # 去除行中非文本部分，包括特效标签{}，硬软换行符
                eventtext = re.sub(effectDel, '', eventftext)
                fontlist = fontlistAdd(eventtext, eventfont, fontlist)

    fl_popkey = []
    # 在字体列表中检查是否有没有在文本中使用的字体，如果有，添加到删去列表
    for s in fontlist.keys():
        if len(fontlist[s]) == 0:
            fl_popkey.append(s)
            # print('跳过没有字符的字体\"{0}\"'.format(s))
    # 删去 删去列表 中的字体
    if len(fl_popkey) > 0:
        for s in fl_popkey:
            fontlist.pop(s)
    # print(fontlist)
    # 如果含有字体恢复信息，则恢复字幕
    if len(ssRecover) > 0:
        try:
            nsdir = path.join(path.dirname(asspath), 'NoSubsetSub')
            # nsfile = path.join(nsdir, '.NoSubset'.join(path.splitext(path.basename(asspath))))
            nsfile = path.join(nsdir, path.basename(asspath))
            if not path.exists(nsdir): os.mkdir(nsdir)
            recoverass = open(nsfile, mode='w', encoding='utf-8')
            recoverass.writelines(fullass)
            recoverass.close()
            print('\033[1;33m已恢复字幕: \033[0m\033[1m\"{0}\"\033[0m'.format(nsfile))
        except:
            print('\033[1;31m[ERROR] 恢复的子集化字幕写入失败\n{0}\033[0m'.format(sys.exc_info()))
    # 如果 onlycheck 为 True，只返回字体列表
    if onlycheck:
        del fullass
        style_font.clear()
        return None, fontlist, None, None, None, None

    return fullass, fontlist, styleline, font_pos, fn_lines, infoline


def getFileList(customPath: list = [], font_name: dict = {}, noreg: bool = False):
    """
获取字体文件列表

接受输入
  customPath: 用户指定的字体文件夹
  font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）
  noreg: 只从用户提供的customPath获取输入

将会返回
  filelist: 字体文件清单 [[ 字体绝对路径, 读取位置（'': 注册表, '0': 自定义目录） ], ...]
  font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）
    """
    global ui 
    filelist = []

    if not noreg:
        # 从注册表读取
        fontkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
        fontkey_num = winreg.QueryInfoKey(fontkey)[1]
        # fkey = ''
        try:
            # 从用户字体注册表读取
            fontkey10 = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
            fontkey10_num = winreg.QueryInfoKey(fontkey10)[1]
            if fontkey10_num > 0:
                for i in range(fontkey10_num):
                    p = winreg.EnumValue(fontkey10, i)[1]
                    # n = winreg.EnumValue(fontkey10, i)[0]
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
            # 从 系统字体注册表 读取
            k = winreg.EnumValue(fontkey, i)[1]
            # n = winreg.EnumValue(fontkey, i)[0]
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
                    ui.statusbar.showMessage('请稍等，获取自定义文件夹中的字体 \"{0}\"'.format(path.basename(path.dirname(p))) , 5000)
                    QApplication.processEvents()
                    p = path.join(r, p)
                    if path.splitext(p)[1][1:].lower() not in ['ttf', 'ttc', 'otc', 'otf']: continue
                    filelist.append([path.join(s, p), 'xxx'])

    # print(font_name)
    # os.system('pause')
    return filelist, font_name


def fnReadCorrect(ttfont: ttLib.ttFont, index: int, fontpath: str) -> str:
    """修正字体内部名称解码错误"""
    name = ttfont['name'].names[index]
    namestr = name.toBytes().decode('utf-16-be', errors='ignore')
    try:
        # 尝试使用 去除 \x00 字符 解码
        print('')
        if len([i for i in name.toBytes() if i == 0]) > 0:
            nnames = ttfont['name']
            namebyteL = name.toBytes()
            for bi in range(0, len(namebyteL), 2):
                if namebyteL[bi] != 0:
                    print('\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
                    return namestr
            namebyte = b''.join([bytes.fromhex('{:0>2}'.format(hex(i)[2:])) for i in name.toBytes() if i > 0])
            nnames.setName(namebyte,
                           name.nameID, name.platformID, name.platEncID, name.langID)
            namestr = nnames.names[index].toStr()
            print('\033[1;33m已修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
            # os.system('pause')
        else:
            namestr = name.toBytes().decode('utf-16-be', errors='ignore')
        # 如果没有 \x00 字符，使用 utf-16-be 强行解码；如果有，尝试解码；如果解码失败，使用 utf-16-be 强行解码
    except:
        print('\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
    return namestr


def outputSameLength(s: str) -> str:
    length = 0
    output = ''
    for i in range(len(s) - 1, -1, -1):
        si = s[i]
        if 65281 <= ord(si) <= 65374 or ord(si) == 12288 or ord(si) not in range(33, 127):
            length += 2
        else:
            length += 1
        if length >= 56:
            output = '..' + output
            break
        else:
            output = si + output
    return output + ''.join([' ' for _ in range(0, 60 - length)])


def fontProgress(fl: list, font_info: list = None, overwrite: bool = False) -> list:
    """
字体处理部分

需要输入
  fl: 字体文件列表

可选输入
  f_n: 默认新建一个，可用于更新font_name

将会返回
  font_name: 字体内部名称与绝对路径的索引词典

会对以下全局变量进行变更
  dupfont: 重复字体的名称与其路径词典
  font_n_lower: 字体全小写名称与其标准名称对应词典
    """
    global dupfont, ui
    if font_info != None:
        warning_font = font_info[3]
        font_family = font_info[2]
        font_n_lower = font_info[1]
        f_n = font_info[0]
    else:
        warning_font = []
        font_family = {}
        font_n_lower = {}
        f_n = {}
    # print(fl)
    flL = len(fl)
    print('\033[1;32m正在读取字体信息...\033[0m')
    ui.pBar.setMaximum(flL)
    for si in range(0, flL):
        QApplication.processEvents()
        s = fl[si][0]
        fromCustom = False
        fl[si][1] = str(fl[si][1])
        if len(fl[si][1]) > 0: fromCustom = True
        # 如果有来自自定义文件夹标记，则将 fromCustom 设为 True
        ext = path.splitext(s)[1][1:]
        # 检查字体扩展名
        if ext.lower() not in ['ttf', 'ttc', 'otf', 'otc']: continue
        # 输出进度
        print('\r' + '\033[1;32m{0}/{1} {2:.2f}%\033[0m \033[1m{3}\033[0m'.format(si + 1, flL, ((si + 1) / flL) * 100,
                                                                                  outputSameLength(path.dirname(s))),
              end='', flush=True)
        ui.pBar.setValue(si + 1)
        ui.statusbar.showMessage('正在读取字体信息 第{0}个/共{1}个 \"{2}\"'.format(si + 1, flL, path.basename(s)))
        isTc = False
        if ext.lower() in ['ttf', 'otf']:
            # 如果是 TTF/OTF 单字体文件，则使用 TTFont 读取
            try:
                tc = [ttLib.TTFont(s, lazy=True)]
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTC/OTC模式读取\033[0m'
                      .format(s, sys.exc_info()))
                # 如果 TTFont 读取失败，可能是使用了错误扩展名的 TTC/OTC 文件，换成 TTCollection 尝试读取
                try:
                    tc = ttLib.TTCollection(s, lazy=True)
                    print('\033[1;34m[WARNING] 错误的字体扩展名\"{0}\" \033[0m'.format(s))
                    isTc = True
                except:
                    print('\033[1;31m\n[ERROR] \"{0}\": {1}\033[0m'.format(s, sys.exc_info()))
                    continue
        else:
            try:
                # 如果是 TTC/OTC 字体集合文件，则使用 TTCollection 读取
                tc = ttLib.TTCollection(s, lazy=True)
                isTc = True
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTF/OTF模式读取\033[0m'
                      .format(s, sys.exc_info()))
                try:
                    # 如果读取失败，可能是使用了错误扩展名的 TTF/OTF 文件，用 TTFont 尝试读取
                    tc = [ttLib.TTFont(s, lazy=True)]
                    print('\033[1;34m[WARNING] 错误的字体扩展名\"{0}\" \033[0m'.format(s))
                except:
                    print('\033[1;31m\n[ERROR] \"{0}\": {1}\033[0m'.format(s, sys.exc_info()))
                    continue
            # f_n[path.splitext(path.basename(s))[0]] = [s, 0]
        for ti in range(0, len(tc)):
            t = tc[ti]
            # 读取字体的 'OS/2' 表的 'fsSelection' 项查询字体的粗体斜体信息
            try:
                os_2 = bin(t['OS/2'].fsSelection)[2:].zfill(10)
            except:
                os_2 = '1'.zfill(10)
            isItalic = int(os_2[-1])
            isBold = int(os_2[-6])
            # isCustom = int(os_2[-7])
            # isRegular = int(os_2[-7])
            # 读取字体的 'name' 表
            fstyle = ''
            familyN = ''
            try:
                indexs = len(t['name'].names)
            except:
                break
            isWarningFont = False
            namestrl1 = []
            namestrl2 = []
            donotStyle = False
            for ii in range(0, indexs):
                name = t['name'].names[ii]
                # 若 nameID 为 1，读取 NameRecord 的字体家族名称
                if name.nameID == 1:
                    try:
                        familyN = name.toStr()
                    except:
                        familyN = fnReadCorrect(t, ii, s)
                    familyN = familyN.strip(' ')
                    font_family.setdefault(familyN, {})
                # 若 nameID 为 2，读取 NameRecord 的字体样式
                elif name.nameID == 2 and not donotStyle:
                    try:
                        fstyle = name.toStr()
                    except:
                        fstyle = fnReadCorrect(t, ii, s)
                # 若 nameID 为 4，读取 NameRecord 的字体完整名称           
                elif name.nameID == 4:
                    try:
                        # if name.isUnicode() or name.platformID == 0:
                        namestr = name.toStr()
                        # else:
                        #     errnum = 0
                        #     lcidcache = [lcc for lcc in lcidfil[name.platformID].values()
                        #     if re.search(r'(unicode)|(utf)|(roman)', lcc.lower()) is None]
                        #     for lcid in lcidcache:
                        #         try:
                        #             namestr = name.toBytes().decode(lcid)
                        #         except:
                        #             errnum += 1
                        #     if errnum >= len(lcidcache):
                        #         namestr = fnReadCorrect(t, ii, s)
                        #         isWarningFont = True
                    except:
                        # 如果 fontTools 解码失败，则尝试使用 utf-16-be 直接解码
                        namestr = fnReadCorrect(t, ii, s)
                        isWarningFont = True
                    if namestr is None: continue
                    if name.langID in lcidfil[name.platformID].keys():
                        if namestr.strip(' ') not in namestrl1: namestrl1.append(namestr.strip(' '))
                        if namestr.strip(' ').lower().find(familyN.lower()) > -1: donotStyle = True
                    else:
                        if namestr.strip(' ') not in namestrl2: namestrl2.append(namestr.strip(' '))
                    del namestr
            # print(namestrl)
            if len(namestrl1) > 0:
                namestrl = namestrl1
            elif len(namestrl2) > 0:
                namestrl = namestrl2
            else:
                continue
            del namestrl1, namestrl2
            addfmn = True
            for ns in namestrl:
                if ns.lower().find(familyN.lower()) > -1:
                    addfmn = False
            if addfmn: namestrl.append(familyN.strip(' '))
            # print(namestr, path.basename(s))
            for namestr in namestrl:
                # print(namestr)
                if isWarningFont:
                    if namestr not in warning_font:
                        warning_font.append(namestr.lower())

                if f_n.get(namestr) is not None and not overwrite:
                    # 如果发现列表中已有相同名称的字体，检测它的文件名、扩展名、父目录、样式是否相同
                    # 如果有一者不同且不来自自定义文件夹，添加到重复字体列表
                    dupp = f_n[namestr][0]
                    if (dupp != s and path.splitext(path.basename(dupp))[0] != path.splitext(path.basename(s))[0] and
                            not fromCustom and f_n[namestr][2].lower() == fstyle.lower()):
                        print('\n\033[1;35m[WARNING] 字体\"{0}\"与字体\"{1}\"的名称\"{2}\"重复！\033[0m'.format(
                            path.basename(f_n[namestr][0]), path.basename(s), namestr))
                        if dupfont.get(namestr) is not None:
                            if s not in dupfont[namestr]:
                                dupfont[namestr].append(s)
                        else:
                            dupfont[namestr] = [dupp, s]
                else:
                    f_n[namestr] = [s, ti, fstyle, isTc]
                font_n_lower[namestr.lower()] = namestr
                    # print(f_n[namestr], namestr)
                if len(familyN) > 0 and namestr != namestrl[-1]:
                    if font_family[familyN].get((isItalic, isBold)) is not None:
                        if namestr not in font_family[familyN][(isItalic, isBold)]:
                            font_family[familyN][(isItalic, isBold)].append(namestr)
                    else:
                        font_family[familyN].setdefault((isItalic, isBold), [namestr])
        tc[0].close()
    keys = list(font_family.keys())
    for k in keys:
        if not len(font_family[k]) > 1:
            font_family.pop(k)
    del keys
    dupfont_cache = dupfont
    dupfont = {}
    for k in dupfont_cache.keys():
        k2 = tuple(dupfont_cache[k])
        if dupfont.get(k2) is not None:
            dupfont[k2].append(k)
        else:
            dupfont[k2] = [k]
    del dupfont_cache
    ui.pBar.setValue(0)
    ui.statusbar.showMessage("字体信息读取完成", 5000)
    return [f_n, font_n_lower, font_family, warning_font]


# print(filelist)
# if path.exists(fontspath10): filelist = filelist.extend(os.listdir(fontspath10))

# for s in font_name.keys(): print('{0}: {1}'.format(s, font_name[s]))
def fnGetFromFamilyName(font_family: dict, fn: str, isitalic: int, isbold: int) -> str:
    if font_family.get(fn) is not None:
        if font_family[fn].get((isitalic, isbold)) is not None:
            return font_family[fn][(isitalic, isbold)][0]
        elif font_family[fn].get((0, isbold)) is not None:
            return font_family[fn][(0, isbold)][0]
        elif font_family[fn].get((isitalic, 0)) is not None:
            return font_family[fn][(isitalic, 0)][0]
        elif font_family[fn].get((0, 0)) is not None:
            return font_family[fn][(0, 0)][0]
        else:
            return fn
    else:
        return fn


def checkAssFont(fontlist: dict, font_info: list, fn_lines: list = [], onlycheck: bool = False):
    """
系统字体完整性检查，检查是否有ASS所需的全部字体，如果没有，则要求拖入

需要以下输入
  fontlist: 字体与其所需字符（只读字体部分） { ASS内的字体名称 : 字符串 }

  font_name: 字体名称与字体路径对应词典 { 字体名称 : [ 字体绝对路径, 字体索引 ] }

以下输入可选
  fn_lines: fn标签所指定的字体与其包含的文本
  onlycheck: 缺少字体时不要求输入

将会返回以下
  assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称, 修正名称 ]}
  font_name: 同上，用于有新字体拖入时对该词典的更新
    """
    # 从fontlist获取字体名称
    global ui
    assfont = {}
    font_name = font_info[0]
    font_n_lower = font_info[1]
    font_family = font_info[2]
    warning_font = font_info[3]
    keys = list(fontlist.keys())
    for s in keys:
        sp = s.split('?')
        isbold = int(sp[2])
        isitalic = int(sp[1])
        ns = sp[0]
        ns = fnGetFromFamilyName(font_family, ns, isitalic, isbold)
        cok = False
        # 在全局字体名称词典中寻找字体名称
        ss = ns
        if ns not in font_name:
            # 如果找不到，将字体名称统一为小写再次查找
            if font_n_lower.get(ns.lower()) is not None:
                ss = font_n_lower[ns.lower()]
                cok = True
        elif not path.exists(font_name[ns][0]):
            del font_name[ns]
            cok = False
        else:
            cok = True
        if not cok:
            # 如果 onlycheck 不为 True，向用户要求目标字体
            if not onlycheck:
                print('\033[1;31m[ERROR] 缺少字体\"{0}\"\n请输入追加的字体文件或其所在字体目录的绝对路径\033[0m'.format(ns))
                ui.statusbar.showMessage('缺少字体\"{0}\"\n请在\"自定义字体目录\"追加的字体文件或其所在字体目录的绝对路径')
                return None, [font_name, font_n_lower, font_family, warning_font]
            else:
                # 否则直接添加空
                assfont['?'.join([ns, ns])] = ['', sp[0], ns]
        if cok:
            if not onlycheck and ss.lower() in warning_font:
                print('\033[1;31m[WARNING] 字体\"{0}\"可能不能正常子集化\033[0m'.format(ss))
                if warningStop:
                    print('\033[1;31m[WARNING] 请修复\"{0}\"，工作已中断\033[0m'.format(ss))
                    ui.statusbar.showMessage('请修复\"{0}\"，工作已中断'.format(ss))
                    return None, [font_name, font_n_lower, font_family, warning_font]
            # 如果找到，添加到assfont列表
            font_path = font_name[ss][0]
            font_index = font_name[ss][1]
            # print(font_name[ss])
            dict_key = '?'.join([font_path, str(font_index)])
            # 如果 assfont 列表已有该字体，则将新字符添加到 assfont 中
            if assfont.get(dict_key) is None:
                assfont[dict_key] = [fontlist[s], sp[0], ns]
            else:
                tfname = assfont[dict_key][2]
                newfnamep = assfont[dict_key][1].split('|')
                oldstr = fontlist[s]
                for newfname in newfnamep:
                    if sp[0].lower() not in '|'.join(newfnamep).lower():
                        key1 = 0
                        key2 = 0
                        for ii in [0, 1]:
                            key1 = ii
                            key2 = 0
                            if fontlist.get('?'.join([newfname, str(key1), str(0)])) is not None:
                                break
                            elif fontlist.get('?'.join([newfname, str(key1), str(1)])) is not None:
                                key2 = 1
                                break
                        newfstr = fontlist['?'.join([newfname, str(key1), str(key2)])]
                        newfname = '|'.join([sp[0], newfname])
                        for i in range(0, len(newfstr)):
                            if newfstr[i] not in oldstr:
                                oldstr += newfstr[i]
                    else:
                        newstr = assfont[dict_key][0]
                        for i in range(0, len(newstr)):
                            if newstr[i] not in oldstr:
                                oldstr += newstr[i]
                    if ns.lower() not in tfname.lower():
                        tfname = '|'.join([ns, tfname])
                    assfont[dict_key] = [oldstr, newfname, tfname]
                    fontlist[s] = oldstr
                # else:
                #     assfont[dict_key] = [fontlist[s], s]
            # print(assfont[dict_key])
    if len(fn_lines) > 0 and not onlycheck:
        fn_lines_cache = fn_lines
        for i in range(0, len(fn_lines_cache)):
            s = fn_lines_cache[i]
            fi = s[1][s[1].keys()[0]]
            fn_lines[i] = [s[0], {s[1].keys()[0]: [fi[1], fnGetFromFamilyName(font_family, fi[1], fi[2], fi[3])]}]
    # print(assfont)
    return assfont, [font_name, font_n_lower, font_family, warning_font]


# print('正在输出字体子集字符集')
# for s in fontlist.keys():
#     logpath = '{0}_{1}.log'.format(path.join(os.getenv('TEMP'), path.splitext(path.basename(asspath))[0]), s)
#     log = open(logpath, mode='w', encoding='utf-8')
#     log.write(fontlist[s])
#     log.close()

def getNameStr(name, subfontcrc: str) -> str:
    """
字体内部名称变更

变更NameID为1, 3, 4, 6的NameRecord，它们分别对应

    ID  Meaning

    1   Font Family name

    3   Unique font identifier

    4   Full font name

    6   PostScript name for the font

注意本脚本并不更改 NameID 为 0 和 7 的版权信息
"""
    nameID = name.nameID

    if nameID in [1, 3, 4, 6]:
        namestr = subfontcrc
    else:
        try:
            namestr = name.toStr()
        except:
            namestr = name.toBytes().decode('utf-16-be', errors='ignore')
    return namestr


def assFontSubset(assfont: dict, fontdir: str):
    """
字体子集化

需要以下输入:
  assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称 ]}
  fontdir: 新字体存放目录

将会返回以下:
  newfont_name: { 原字体名 : [ 新字体绝对路径, 新字体名 ] }
    """
    global errorStop, ui, MW
    newfont_name = {}
    # print(fontdir)

    if path.exists(path.dirname(fontdir)):
        if not path.isdir(fontdir):
            try:
                os.mkdir(fontdir)
            except:
                print('\033[1;31m[ERROR] 创建文件夹\"{0}\"失败\033[0m'.format(fontdir))
                fontdir = os.getcwd()
        if not path.isdir(fontdir): fontdir = path.dirname(fontdir)
    else:
        fontdir = os.getcwd()
    print('\033[1;33m字体输出路径:\033[0m \033[1m\"{0}\"\033[0m'.format(fontdir))

    lk = len(assfont.keys())
    kip = 0
    for k in assfont.keys():
        QApplication.processEvents()
        kip += 1
        # 偷懒没有变更该函数中的assfont解析到新的词典格式
        # 在这里会将assfont词典转换为旧的assfont列表形式
        # assfont: [ 字体绝对路径, 字体索引, 字符串, ASS内的字体名称, 修正字体名称 ]
        s = k.split('?') + [assfont[k][0], assfont[k][1], assfont[k][2]]
        fontext = path.splitext(path.basename(s[0]))[1]
        if fontext[1:].lower() in ['otc', 'ttc']:
            subfontext = fontext[:3].lower() + 'f'
        else:
            subfontext = fontext
        # print(fontdir, path.exists(path.dirname(fontdir)), path.exists(fontdir))
        fontname = re.sub(cillegal, '_', s[4])
        subfontpath = path.join(fontdir, fontname + subfontext)
        subsetarg = [s[0], '--text={0}'.format(s[2]), '--output-file={0}'.format(subfontpath),
                     '--font-number={0}'.format(s[1]), '--passthrough-tables']
        print('\r\033[1;32m[{0}/{1}]\033[0m \033[1m正在子集化…… \033[0m'.format(kip, lk), end='')
        ui.statusbar.showMessage('正在子集化 {1}/{2} \"{0}\" '.format(fontname, kip, lk))
        try:
            subset.main(subsetarg)
        except PermissionError:
            print('\n\033[1;31m[ERROR] 文件\"{0}\"访问失败\033[0m'.format(path.basename(subfontpath)))
            ui.statusbar.showMessage('文件\"{0}\"访问失败'.format(path.basename(subfontpath)))
            continue
        except:
            # print('\033[1;31m[ERROR] 失败字符串: \"{0}\" \033[0m'.format(s[2]))
            print('\n\033[1;31m[ERROR] {0}\033[0m'.format(sys.exc_info()))
            if errorStop:
                print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，强制终止批量处理\033[0m'.format(path.basename(s[0])))
                QMessageBox.critical(MW, '错误', '字体\"{0}\"子集化失败\n将会强制终止批量处理'.format(path.basename(s[0])), QMessageBox.StandardButton.Ok)
                return None
            print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
            ui.statusbar.showMessage('字体\"{0}\"子集化失败，将会保留完整字体'.format(path.basename(s[0])))
            # crcnewf = ''.join([path.splitext(subfontpath)[0], fontext])
            # shutil.copy(s[0], crcnewf)
            ttLib.TTFont(s[0], lazy=False, fontNumber=int(s[1])).save(subfontpath, False)
            subfontcrc = None
            # newfont_name[s[3]] = [crcnewf, subfontcrc]
            if s[3].split('|') == 1:
                newfont_name[s[3]] = [subfontpath, subfontcrc]
            else:
                for sp in s[3].split('|'):
                    newfont_name[sp] = [subfontpath, subfontcrc]
            continue
        # os.system('pyftsubset {0}'.format(' '.join(subsetarg)))
        if path.exists(subfontpath):
            subfontbyte = open(subfontpath, mode='rb')
            subfontcrc = str(hex(zlib.crc32(subfontbyte.read())))[2:].upper()
            if len(subfontcrc) < 8: subfontcrc = '0' + subfontcrc
            # print('CRC32: {0} \"{1}\"'.format(subfontcrc, path.basename(s[0])))
            subfontbyte.close()
            rawf = ttLib.TTFont(s[0], lazy=True, fontNumber=int(s[1]))
            newf = ttLib.TTFont(subfontpath, lazy=False)
            if len(newf['name'].names) == 0:
                for i in range(0, 7):
                    if len(rawf['name'].names) - 1 >= i:
                        name = rawf['name'].names[i]
                        namestr = getNameStr(name, subfontcrc)
                        newf['name'].addName(namestr, minNameID=-1)
            else:
                for i in range(0, len(rawf['name'].names)):
                    name = rawf['name'].names[i]
                    namestr = getNameStr(name, subfontcrc)
                    newf['name'].setName(namestr, name.nameID, name.platformID, name.platEncID, name.langID)
            if len(newf.getGlyphOrder()) == 1 and '.notdef' in newf.getGlyphOrder():
                if errorStop:
                    print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，强制终止批量处理\033[0m'.format(path.basename(s[0])))
                    QMessageBox.critical(MW, '错误', '字体\"{0}\"子集化失败\n将会强制终止批量处理'.format(path.basename(s[0])), QMessageBox.StandardButton.Ok)
                    rawf.close()
                    newf.close()
                    if not subfontpath == s[0]: os.remove(subfontpath)
                    return None
                print('\n\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
                ui.statusbar.showMessage('字体\"{0}\"子集化失败，将会保留完整字体'.format(path.basename(s[0])))
                crcnewf = subfontpath
                newf.close()
                if not subfontpath == s[0]: os.remove(subfontpath)
                # shutil.copy(s[0], crcnewf)
                rawf.save(crcnewf, False)
                subfontcrc = None
            else:

                crcnewf = '.{0}'.format(subfontcrc).join(path.splitext(subfontpath))
                newf.save(crcnewf)
                newf.close()
            rawf.close()
            if path.exists(crcnewf):
                if not subfontpath == crcnewf: os.remove(subfontpath)
                if s[3].split('|') == 1:
                    newfont_name[s[3]] = [crcnewf, subfontcrc]
                else:
                    for sp in s[3].split('|'):
                        newfont_name[sp] = [crcnewf, subfontcrc]
    print('')
    # print(newfont_name)
    return newfont_name


def assFontChange(fullass: list, newfont_name: dict, asspath: str, styleline: int, infoline: int,
                  font_pos: int, outdir: str = '', ncover: bool = False, fn_lines: list = []) -> str:
    """
更改ASS样式对应的字体

需要以下输入
  fullass: 完整的ass文件内容，以行分割为列表
  newfont_name: { 原字体名 : [ 新字体路径, 新字体名 ] }
  asspath: 原ass文件的绝对路径
  styleline: [V4/V4+ Styles]标签在SSA/ASS中的行数，对应到fullass列表的索引
  font_pos: Font参数在 Styles 的 Format 中位于第几个逗号之后
  infoline: [Script Info]标签在SSA/ASS中的行数，对应到fullass列表的索引

以下输入可选
  outdir: 新字幕的输出目录，默认为源文件目录
  ncover: 为True时不覆盖原有文件，为False时覆盖
  fn_lines: 带有fn标签的行数，对应到fullass的索引

将会返回以下
  newasspath: 新ass文件的绝对路径
    """
    # 扫描Style各行，并替换掉字体名称
    # print('正在替换style对应字体......')
    for i in range(styleline + 2, len(fullass)):
        if len(fullass[i].split(':')) < 2:
            continue
        if fullass[i].split(':')[0].strip(' ').lower() != 'style':
            break
        styleStr2 = fullass[i].split(':')
        styleStr = ''.join(styleStr2[1:]).split(',')
        fontstr = styleStr[font_pos].strip(' ').lstrip('@')
        if not newfont_name.get(fontstr) is None:
            if not newfont_name[fontstr][1] is None:
                styleStr[font_pos] = styleStr[font_pos].replace(fontstr, newfont_name[fontstr][1])
                fullass[i] = styleStr2[0] + ':' + ','.join(styleStr)
    if len(fn_lines) > 0:
        # print('正在处理fn标签......')
        for fl in fn_lines:
            fn_line = fullass[fl[0]]
            for ti in range(1, len(fl)):
                for k in fl[ti].keys():
                    fname = fl[ti][k][0]
                    if len(fname) > 0:
                        fn_line = fn_line.replace(k, k.replace(fname, newfont_name[fname][1]))
            fullass[fl[0]] = fn_line
    for k in newfont_name.keys():
        fullass.insert(infoline + 1, '; Font Subset: {1} - {0}\n'.format(k, newfont_name[k][1]))
    fullass.insert(infoline + 1, '; Subset via ASFMKV_py 1.02-Pre9 (fontTools)\n')
    if path.exists(path.dirname(outdir)):
        if not path.isdir(outdir):
            try:
                os.mkdir(outdir)
            except:
                print('\033[1;31m[ERROR] 创建文件夹\"{0}\"失败\033[0m'.format(outdir))
                outdir = os.getcwd()
    print('\033[1;33m字幕输出路径:\033[0m \033[1m\"{0}\"\033[0m'.format(outdir))
    if path.isdir(outdir):
        newasspath = path.join(outdir, '.subset'.join(path.splitext(path.basename(asspath))))
    else:
        newasspath = '.subset'.join(path.splitext(asspath))
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
    # print('ASS样式转换完成: {0}'.format(path.basename(newasspath)))
    return newasspath


def ASFMKV(file: str, outfile: str = '', asslangs: list = [], asspaths: list = [], fontpaths: list = []) -> int:
    """
ASFMKV，将媒体文件、字幕、字体封装到一个MKV文件，需要mkvmerge命令行支持

需要以下输入

  file: 媒体文件绝对路径
  outfile: 输出文件的绝对路径，如果该选项空缺，默认为 输入媒体文件.muxed.mkv
  asslangs: 赋值给字幕轨道的语言，如果字幕轨道多于asslangs的项目数，超出部分将全部应用asslangs的末项
  asspaths: 字幕绝对路径列表
  fontpaths: 字体列表，格式为 [[字体1绝对路径], [字体1绝对路径], ...]，必须嵌套一层，因为主函数偷懒了

将会返回以下
  mkvmr: mkvmerge命令行的返回值    
    """
    # print(fontpaths)
    global rmAssIn, rmAttach, mkvout, notfont, ui
    if file is None:
        return 4
    elif file == '':
        return 4
    elif not path.exists(file) or not path.isfile(file):
        return 4
    if outfile is None: outfile = ''
    if outfile == '' or not path.exists(path.dirname(outfile)) or path.dirname(outfile) == path.dirname(file):
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
            # print(assfn, fn, assnote)
            if len(assnote) > 1:
                mkvargs.extend(['--track-name', '0:{0}'.format(assnote.lstrip('.'))])
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
    ui.statusbar.showMessage('请稍等，正在封装\"{0}\"'.format(path.basename(file)))
    QApplication.processEvents()
    mkvmr = os.system('mkvmerge @\"{0}\" -o \"{1}\"'.format(mkvjsonp, outfile))

    ui.tableWidget.insertRow(0)

    if mkvmr > 1:
        print('\n\033[1;31m[ERROR] 检测到不正常的mkvmerge返回值，重定向输出...\033[0m')
        ui.statusbar.showMessage('检测到不正常的mkvmerge返回值，正在重定向输出')
        QApplication.processEvents()
        os.system('mkvmerge -r \"{0}\" @\"{1}\" -o NUL'.format('{0}.{1}.log'
                                                               .format(path.splitext(file)[0],
                                                                       datetime.now().strftime('%Y-%m%d-%H%M-%S_%f')), mkvjsonp))
        ui.tableWidget.setItem(0, 0, QTWI_color(path.basename(outfile), 255, 255, 0))
        ui.tableWidget.setColumnWidth(0, 355)
    elif not notfont:
        for p in asspaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(p))
            if path.splitext(p)[1][1:].lower() in ['ass', 'ssa']:
                try:
                    os.remove(p)
                except:
                    print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(p))
        for f in fontpaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(f[0]))
            try:
                os.remove(f[0])
            except:
                print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(f[0]))
        try:
            os.remove(mkvjsonp)
        except:
            print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(mkvjsonp))
        ui.tableWidget.setItem(0, 0, QTWI_color(path.basename(outfile), 0, 255, 0))
        ui.tableWidget.setColumnWidth(0, 365)
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    else:
        ui.tableWidget.setItem(0, 0, QTWI_color(path.basename(outfile), 0, 255, 0))
        ui.tableWidget.setColumnWidth(0, 365)
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    return mkvmr


def getMediaFilelist(mdir: str) -> list:
    """
从输入的目录中获取媒体文件列表
需要以下输入
  dir: 要搜索的目录
返回以下结果
  medias: 多媒体文件列表
      结构: [[ 文件名（无扩展名）, 绝对路径 ], ...]
    """
    medias = []
    global v_subdir, extlist
    if path.isdir(mdir):
        if v_subdir:
            for r, ds, fs in os.walk(mdir):
                for f in fs:
                    if path.splitext(f)[1][1:].lower() in extlist:
                        medias.append([path.splitext(path.basename(f))[0], path.join(r, f)])
        else:
            for f in os.listdir(mdir):
                if path.isfile(path.join(mdir, f)):
                    if path.splitext(f)[1][1:].lower() in extlist:
                        medias.append([path.splitext(path.basename(f))[0], path.join(mdir, f)])
        return medias


def getSubtitles(cpath: str, medias: list = [], spcific: bool = False) -> dict:
    """
在目录中找到与媒体文件列表中的媒体文件对应的字幕

遵循以下原则
  媒体文件在上级目录，则匹配子目录中的字幕；媒体文件的字幕只能在媒体文件的同一目录或子目录中，不能在上级目录和其他同级目录

需要以下输入
  cpath: 开始搜索的顶级目录

以下输入可选
  medias: 媒体文件列表，结构见 getMediaFilelist
      在不输入 medias 时，将直接返回全部字幕文件于 media_ass['none']

将会返回以下
  media_ass: 媒体文件与字幕文件的对应词典
      结构: { 媒体文件绝对路径 : [ 字幕1绝对路径, 字幕2绝对路径, ...] }
    """
    media_ass = {}
    if len(medias) == 0: media_ass['none'] = []
    global s_subdir, matchStrict
    if s_subdir:
        for r, ds, fs in os.walk(cpath):
            for f in [path.join(r, s) for s in fs if path.splitext(s)[1][1:].lower() in subext]:
                if '.subset' in path.basename(f): continue
                if len(medias) > 0:
                    for li in medias:
                        vdir = path.dirname(li[1])
                        sdir = path.dirname(f)
                        sext = path.splitext(f)
                        if (li[0] in f and not matchStrict) or (li[0] == path.basename(f)[:len(li[0])] and matchStrict):
                            if (vdir in sdir and sdir not in vdir) or (vdir == sdir):
                                if sext[1][:1].lower() == 'idx':
                                    if not path.exists(sext[1] + '.sub'):
                                        continue
                                if media_ass.get(li[1]) is None:
                                    media_ass[li[1]] = [f]
                                else:
                                    media_ass[li[1]].append(f)
                else:
                    media_ass['none'].append(path.join(r, ds, fs))
    else:
        for f in [path.join(cpath, s) for s in os.listdir(cpath)
                  if not path.isdir(s) and path.splitext(s)[1][1:].lower() in subext]:
            # print(f, cpath)
            if '.subset' in path.basename(f): continue
            if len(medias) > 0:
                for li in medias:
                    # print(path.basename(f)[len(l[0]):], l)
                    sext = path.splitext(f)
                    if (li[0] in f and not matchStrict) or (li[0] == path.basename(f)[:len(li[0])] and matchStrict):
                        if path.dirname(li[1]) == path.dirname(f) or spcific:
                            if sext[1][:1].lower() == 'idx':
                                if not path.exists(sext[1] + '.sub'):
                                    continue
                            if media_ass.get(li[1]) is None:
                                media_ass[li[1]] = [f]
                            else:
                                media_ass[li[1]].append(f)
            else:
                media_ass['none'].append(f)
    return media_ass


def nameMatchingProgress(medias: list, mStart: int, mEnd: str, sStart: int, sEnd: str, subs: list) -> dict:
    """
视频-字幕匹配的流程部分，负责匹配视频和字幕

需要以下输入
  medias: 媒体文件列表，格式见 nameMatching
  mStart: 媒体文件剧集开始位置
  mEnd: 媒体文件剧集结束字符
  sStart: 字幕文件剧集开始位置
  sEnd: 字幕文件剧集结束字符
  subs: 字幕文件列表
    """
    media_ass = {}
    for m in medias:
        m = tuple(m)
        media_ass[m] = []
        mEndI = mStart + m[0][mStart:].find(mEnd)
        if mEndI < mStart: continue
        for s in subs:
            sName = path.basename(s)
            sEndI = sStart + sName[sStart:].find(sEnd)
            if sEndI < sStart: continue
            if m[0][mStart:mEndI].lower() == sName[sStart:sEndI].lower():
                media_ass[m].append(s)
        if len(media_ass[m]) == 0: del media_ass[m]
    return media_ass


def nameMatching(spath: str, medias: list):
    """
字幕匹配的控制部分

需要以下输入
  spath: 字幕路径
  medias: 媒体文件列表，遵循格式为 [[文件名(无扩展名), 完整路径], ...]
    """
    # 调用各函数获得剧集位置信息及字幕文件列表
    rule_v = ui.text.text().strip()
    rule_s = ui.text2.text().strip()
    annl = []
    for i in range(0, ui.list.count()):
        annl.append(ui.list.item(i).text().strip())
    if len(rule_v) > 0 and len(rule_s) > 0:
        if len(rule_v.split(':')) == 1 or len(rule_s.split(':')) == 1:
            ui.statusbar.showMessage('匹配规则里面没有发现半角冒号 \":\"，是不是忘记填了？', 5000)
            return None
        elif len(rule_v.split(':')) > 2 or len(rule_s.split(':')) > 2:
            ui.statusbar.showMessage('匹配规则里半角冒号 \":\" 太多了（超过一个）', 5000)
            return None
    elif len(rule_v) > 0 or len(rule_s) > 0:
        ui.statusbar.showMessage('做事善始善终，请把空着的匹配规则也填上，要不就全部空着', 5000)
        return None
    if len(medias) == 0:
        print('找不到媒体文件')
        ui.statusbar.showMessage('找不到媒体文件', 5000)
        return None
    subs = getSubtitles(spath)['none']
    if len(subs) == 0:
        print('找不到字幕文件')
        ui.statusbar.showMessage('找不到字幕文件', 5000)
        return None
    if len(rule_v) == 0 and len(rule_s) == 0:
        mStart, mEnd = namePosition([s[0] for s in medias])
        sStart, sEnd = namePosition([path.splitext(path.basename(s))[0] for s in subs])
        sStartA, sEndA = namePosition([path.splitext(path.basename(s))[0][::-1] for s in subs])
        # 生成第一次的 media_ass 和 renList，并询问
        media_ass = nameMatchingProgress(medias, mStart, mEnd, sStart, sEnd, subs)
        renList, isSkip = renamePreview(media_ass, sStartA, sEndA, annl=annl)
    else:
        vSplit = rule_v.split(':')
        sSplit = rule_s.split(':')
        media_ass = nameMatchingProgress(medias, len(vSplit[0]), vSplit[1][0], len(sSplit[0]), sSplit[1][0], subs)
        renList, isSkip = renamePreview(media_ass, annl=annl)
        del vSplit, sSplit
    if len(renList) > 0:
        if QMessageBox.question(MW, '重命名搜索完成', '''请在命令行界面确认重命名是否正确
如果后缀部分不正确您可以先选否在界面上手动输入字幕后缀
如果文件匹配不正确您可以先选否在界面上手动输入匹配规则''', QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == \
    QMessageBox.StandardButton.Yes:
            ui.result.setRowCount(0)
            ui.result.setColumnCount(2)
            ui.result.setRowCount(1)
            ui.result.setHorizontalHeaderLabels(['新名', '原名'])
            ui.result.verticalHeader().setVisible(False)
            try:
                logfile = open(path.join(spath, 'subsRename_{0}.log'.format(datetime.now().strftime('%Y%m%d-%H%M%S'))),
                            mode='w', encoding='utf-8')
                for r in renList:
                    ui.result.insertRow(0)
                    try:
                        os.rename(r[0], path.join(path.dirname(r[0]), r[1]))
                    except:
                        print('\033[31;1m[ERROR] \"{0}\"重命名失败\033[0m'.format(path.basename(r[0])))
                        ui.statusbar.showMessage('重命名失败 \"{0}\"'.format(path.basename(r[0])), 5000)
                        ui.result.setItem(0, 0, QTWI_color(path.basename(r[0]), 255, 0, 0))
                        ui.result.setItem(0, 1, QTWI_color(r[1], 255, 0, 0))
                    else:
                        print('\033[32;1m[SUCCESS]\033[33;0m \"{0}\"\n          \033[0m>>> \033[1m\"{1}\"\033[0m'.format(path.basename(r[0]), r[1]))
                        ui.statusbar.showMessage('重命名成功 \"{0}\"'.format(path.basename(r[0])), 5000)
                        ui.result.setItem(0, 0, QTWI_color(path.basename(r[0]), 0, 255, 0))
                        ui.result.setItem(0, 1, QTWI_color(r[1], 0, 255, 0))
                        try:
                            logfile.write('{0}|{1}\n'.format(r[0], r[1]))
                        except:
                            print('\033[31;1m[ERROR] 日志文件写入失败\033[0m')
                            ui.statusbar.showMessage('日志文件写入失败', 5000)
                    QApplication.processEvents()
                logfile.close()
                del logfile
            except:
                print('\033[31;1m[ERROR] 日志文件写入失败\033[0m')
                ui.statusbar.showMessage('日志文件写入失败', 5000)
        elif not isSkip:
            print('视频-字幕匹配失败，请手动输入匹配规则\n')
            ui.statusbar.showMessage('视频-字幕匹配失败，请手动输入匹配规则', 5000)
        else:
            print('视频-字幕匹配失败，请手动输入匹配规则\n')
            ui.statusbar.showMessage('视频-字幕匹配失败，请手动输入匹配规则', 5000)
#         elif rt == 3:
#             ann = input('请输入（可以为空，可以有多个，按预览顺序用\"\033[33;1m:\033[0m\"分隔）:')
#             if ':' in ann:
#                 annl = [re.sub(cillegal, '_', a) for a in ann.split(':') if len(a) > 0]
#                 ann = ''
#             elif len(ann) > 0:
#                 if ann[0] != '.': ann = '.' + ann
#         else:
#             break
#             rt = 2
    del mStart, mEnd, sStart, sEnd, sStartA, sEndA, media_ass, renList, annl


def renamePreview(media_ass: dict, sStartA: int = 0, sEndA: str = '', ann: str = '', annl: list = []) -> list:
    """
提供重命名预览及制作重命名对应表

需要以下输入
  media_ass: 媒体文件与字幕文件的对应词典，遵循 {(媒体文件名(无扩展名), 完整路径): [字幕完整路径1, 2, ..], ...}

以下输入可选
  sStartA: 字幕后缀开始位置
  sEndA: 字幕后缀结束字符
  ann: 单个字幕后缀
  annl: 多个字幕后缀列表

将会返回以下值
  renList: 字幕重命名前后对应表，格式为 [(重命名前完整路径, 重命名后文件名), ...]
    """
    global matchStrict, ui
    cls()
    print('[重命名预览]\n')
    renList = []
    isSkip = False
    ui.result.setRowCount(0)
    ui.result.setColumnCount(2)
    ui.result.setRowCount(1)
    ui.result.setColumnWidth(0, 260)
    ui.result.setColumnWidth(1, 260)
    ui.result.setHorizontalHeaderLabels(['新名', '原名'])
    ui.result.verticalHeader().setVisible(False)
    for k in media_ass.keys():
        sl = media_ass[k]
        annlUp = 0
        for s in sl:
            bname = path.basename(s)
            if matchStrict:
                if len(bname) >= len(k[0]):
                    if k[0] == bname[:len(k[0])]:
                        print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                        ui.result.insertRow(0)
                        ui.result.setItem(0, 0, QTWI_color('已匹配，跳过重命名', 0, 0, 255))
                        ui.result.setItem(0, 1, QTWI_color(bname, 0, 0, 255))
                        isSkip = True
                        continue
            else:
                if k[0] in bname:
                    print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                    ui.result.insertRow(0)
                    ui.result.setItem(0, 0, QTWI_color('已匹配，跳过重命名', 0, 0, 255))
                    ui.result.setItem(0, 1, QTWI_color(bname, 0, 0, 255))
                    isSkip = True
                    continue
            if len(annl) == 0:
                rsName = path.splitext(bname)[0][::-1]
                if sStartA - 2 > 1:
                    if rsName[sStartA - 2] == '.' and sEndA != '.':
                        ann = rsName[:sStartA - 1][::-1]
                elif rsName.find('.') > 0:
                    ann = '.' + rsName[:rsName.find('.')][::-1]
                    print('使用了暴力后缀抽取，请务必检查后缀是否有误')
                elif len(sl) > 1:
                    ann = '.sub{0}'.format(sl.index(s))
                    print('后缀自动检测失败，自动添加防重名后缀')
                del rsName
            else:
                if sl.index(s) <= len(annl) - 1:
                    ann = '.' + annl[sl.index(s)].lstrip('.')
                else:
                    print('输入的项数不足，已自动补足')
                    ann = '.' + annl[len(annl) - 1].lstrip('.') + '#{0}'.format(annlUp)
                    annlUp += 1
            print("\033[1m\"{0}\"\033[0m".format(bname))
            print('>>>> \033[33;1m\"{0}\033[31;1m{1}\033[33;1m{2}\"\033[0m\n'.format(k[0], ann, path.splitext(s)[1]))
            ui.result.insertRow(0)
            ui.result.setItem(0, 0, QTableWidgetItem('{0}{1}{2}'.format(k[0], ann, path.splitext(s)[1])))
            ui.result.setItem(0, 1, QTableWidgetItem(bname))
            QApplication.processEvents()
            renList.append((s, '{0}{1}{2}'.format(k[0], ann, path.splitext(s)[1])))
    if len(renList) > 0:
        print('字幕的后缀是指类似以下文件名的高亮部分\n\"\033[30;1m[VCB-Studio] Accel World [01][720p][x265_aac]'
              '\033[33;1m.Kamigami-sc\033[30;1m.ass\033[0m\"')
    return renList, isSkip


def namePosition(files: list):
    """
自动匹配文件名中的剧集位置

需要以下输入
  files: 文件列表 [文件名, ...]
    """
    diffList = {'mStart': [], 'mEnd': []}
    sOffset = 1
    step = 1
    for i1 in range(0, len(files)):
        m1 = files[i1]
        print('\r' + '正在计算\033[1;32m{0}/{1} {2:.2f}%\033[0m'.format(i1 + 1, len(files), ((i1 + 1) / len(files)) * 100),
              end='', flush=True)
        diffList[m1] = []
        for m2 in files:
            if m2 == m1:
                continue
            sStop = False
            dStart = 0
            dEndC = [0, 0, 0]
            for i in range(0, len(m1), step):
                if i >= len(m2): break
                if m1[i] != m2[i] or sStop:
                    if sStop:
                        if not m1[i].isnumeric():
                            dEndC[0] = i
                        if not m2[i].isnumeric():
                            dEndC[1] = i
                        if m1[i] == m2[i] and not m1[i].isnumeric() and not m2[i].isnumeric():
                            dEndC[2] = i
                        if dEndC[0] * dEndC[1] * dEndC[2] > 0:
                            break
                    else:
                        sStop = True
                        dStart = i
            dEndCmostC = 0
            dEndCmost = 0
            for i in set(dEndC):
                if dEndCmostC < dEndC.count(i):
                    dEndCmost = i
            if dEndCmostC == 1:
                dEndCmost = dEndC[0]
                for i in dEndC:
                    if i < dEndCmost:
                        dEndCmost = i
            dEnd = m1[dEndCmost]
            del dEndCmost, dEndCmostC
            if len(dEnd) > 0: diffList[m1].append((dStart, dEnd, m2))
            del sStop
        apStart = []
        alStart = []
        apEnd = []
        alEnd = []
        for li in diffList[m1]:
            if li[0] not in apStart: apStart.append(li[0])
            alStart.append(li[0])
            if li[1] not in apEnd: apEnd.append(li[1])
            alEnd.append(li[1])
        mStart = 0
        mStartC = 0
        mEnd = ''
        mEndC = 0
        for li in apStart:
            c = alStart.count(li)
            if c > mStartC:
                mStart, mStartC = li, c
        for li in apEnd:
            c = alEnd.count(li)
            if c > mEndC:
                mEnd, mEndC = li, c
        del diffList[m1]
        diffList['mStart'].append(mStart)
        diffList['mEnd'].append(mEnd)
        del apStart, alStart, apEnd, alEnd
    print('')
    del m1, m2
    mStart, mStartC, mEnd, mEndC = 0, 0, '', 0
    for li in set(diffList['mStart']):
        c = diffList['mStart'].count(li)
        if c > mStartC:
            mStart, mStartC = li, c
    for li in set(diffList['mEnd']):
        c = diffList['mEnd'].count(li)
        if c > mEndC:
            mEnd, mEndC = li, c
    mStartN = mStart
    for li in set(diffList['mStart']):
        if li in range(mStart - sOffset, mStart):
            mStartN = li
    if mStartN < mStart: mStart = mStartN
    del mStartN, mStartC, mEndC, diffList
    # print(diffList['mStart'])
    # print(diffList['mEnd'])
    return mStart, mEnd


def main(font_info: dict, asspath: list, outdir: list = ['', '', ''], mux: bool = False, vpath: str = '',
         asslangs: list = [], noProgress = False):
    """
主函数，负责调用各函数走完完整的处理流程

需要以下输入
  font_name: 字体名称与字体路径对应词典，结构见 fontProgress
  asspath: 字幕绝对路径列表

以下输入可选
  outdir: 输出目录，格式 [ 字幕输出目录, 字体输出目录, 视频输出目录 ]，如果项数不足，则取最后一项；默认为 asspaths 中每项所在目录
  mux: 不要封装，只运行到子集化完成
  vpath: 视频路径，只在 mux = True 时生效
  asslangs: 字幕语言列表，将会按照顺序赋给对应的字幕轨道，只在 mux = True 时生效

将会返回以下
  newasspath: 列表，新生成的字幕文件绝对路径
  newfont_name: 词典，{ 原字体名 : [ 新字体绝对路径, 新字体名 ] }
  ??? : 数值，mkvmerge的返回值；如果 mux = False，返回-1
    """
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
                else:
                    outdir[i] = path.join(os.getcwd(), s)
    # print(outdir)
    # os.system('pause')
    global notfont, ui
    # multiAss 多ASS文件输入记录词典
    # 结构: { ASS文件绝对路径 : [ 完整ASS文件内容(fullass), 样式位置(styleline), 字体在样式行中的位置(font_pos) ]}
    multiAss = {}
    assfont = {}
    fontlist = {}
    newasspath = []
    fo = ''
    if not notfont:
        # print('\n字体名称总数: {0}'.format(len(font_name.keys())))
        # noass = False
        if not noProgress: ui.pBar.setMaximum(len(asspath))
        ui.tabWidget.setCurrentIndex(0)
        for i in range(0, len(asspath)):
            s = asspath[i]
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                multiAss[s] = [[], 0, 0, 0]
            else:
                ui.statusbar.showMessage('正在分析字幕文件: \"{0}\"'.format(path.basename(s)))
                fullass, fontlist, styleline, font_pos, fn_lines, infoline = assAnalyze(s, fontlist)
                multiAss[s] = [fullass, styleline, font_pos, infoline]
                assfont2, font_info = checkAssFont(fontlist, font_info)
                assfont = updateAssFont(assfont, assfont2)
                del assfont2
                if assfont is None:
                    return None, None, -2
            if not noProgress: ui.pBar.setValue(i + 1)
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
        # print('\033[1m字幕所需字体\033[0m')
        # for af in assfont.values():
        #     print('\033[1m\"{0}\"\033[0m: 字符数[\033[1;33m{1}\033[0m]'.format(af[2], len(af[0])))
        newfont_name = assFontSubset(assfont, fo)
        if newfont_name is None: return [], {}, -2
        for s in asspath:
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                newasspath.append(s)
            elif len(multiAss[s][0]) == 0 or multiAss[s][1] == multiAss[s][2]:
                continue
            else:
                newasspath.append(
                    assFontChange(multiAss[s][0], newfont_name, s, multiAss[s][1], multiAss[s][3], multiAss[s][2],
                                  outdir[0], fn_lines=fn_lines))
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
                if path.exists(path.dirname(ap)) and path.splitext(ap)[1][1:].lower() not in ['ass', 'ssa']:
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


def templeFontLoad(_dir: str, font_info: list) -> list:
    """临时字体载入"""
    global s_fontload, ui
    font_list = []
    print('\033[1;32m正在尝试从字幕所在文件夹载入字体...\033[0m')
    ui.statusbar.showMessage('正在尝试从字幕所在文件夹载入字体', 5000)
    if s_fontload:
        for r, ds, fs in os.walk(_dir):
            for f in fs:
                fpath = path.join(r, f)
                if path.isfile(fpath):
                    if path.splitext(fpath)[1][1:].lower() in ['ttf', 'ttc', 'otf', 'otc']:
                        font_list.append([fpath, '1'])
    else:
        fdir = [_dir]
        for ds in os.listdir(_dir):
            if path.isdir(path.join(_dir, ds)):
                if 'font' in ds.lower():
                    fdir.append(path.join(_dir, ds))
        for ds in fdir:
            for fs in os.listdir(ds):
                fpath = path.join(ds, fs)
                if path.isfile(fpath):
                    if path.splitext(fpath)[1][1:].lower() in ['ttf', 'ttc', 'otf', 'otc']:
                        font_list.append([fpath, '1'])
    if len(font_list) > 0:
        font_info = fontProgress(font_list, font_info, True)
        print()
    else:
        print('\033[1;32m目录下没有字体\033[0m')
        ui.statusbar.showMessage('目录下没有字体', 5000)
    return font_info


def cls():
    os.system('cls')


def QTWI_color(text: str, br: int, bg: int, bb: int, ba: int = 60) -> QTableWidgetItem:
    newQTWI = QTableWidgetItem(text)
    newQTWI.setBackground(QColor.fromRgb(br, bg, bb, ba))
    return newQTWI


def cListAssFont(cpath: str):
    global resultw, s_subdir, copyfont, fontload, font_info
    fontlist = {}
    assfont = {}
    assfont2 = {}
    if path.isdir(cpath):
        # print(cpath)
        sublist = []
        if s_subdir:
            for r, ds, fs in os.walk(cpath):
                for f in fs:
                    if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                        sublist.append(path.join(r, ds, f))
        else:
            for f in os.listdir(cpath):
                if path.isfile(path.join(cpath, f)):
                    if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                        sublist.append(path.join(cpath, f))
        fd = path.join(cpath, 'Fonts')
    else:
        sublist = [cpath]
    if fontload:
        if path.isdir(cpath):
            font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
        else:
            font_info2 = templeFontLoad(path.dirname(cpath), copy.deepcopy(font_info))
    else:
        font_info2 = font_info
    ui.pBar.setMaximum(len(sublist))
    for i in range(0, len(sublist)):
        s = sublist[i]
        ui.pBar.setValue(i + 1)
        QApplication.processEvents()
        ui.statusbar.showMessage('正在分析字幕 \"{0}\"'.format(path.basename(s)))
        a, fontlist, b, c, d, g = assAnalyze(s, fontlist, onlycheck=True)
        del a, b, c, d, g
        assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True)
        assfont = updateAssFont(assfont, assfont2)
        del assfont2
    ui.pBar.setValue(0)
    if len(assfont.keys()) < 1:
        ui.statusbar.showMessage('目标路径没有ASS/SSA字幕文件', 5000)
    else:
        ui.pBar.setMaximum(len(assfont.keys()))
        ui.statusbar.showMessage('分析完成', 5000)
        ui.tableWidget.setRowCount(0)
        ui.tableWidget.setColumnCount(3)
        ui.tableWidget.setRowCount(len(assfont.keys()))
        if copyfont: ui.tableWidget.setHorizontalHeaderLabels(['名称', '字体文件', '字符数'])
        else: ui.tableWidget.setHorizontalHeaderLabels(['名称', '字体文件', '状态'])
        ui.tableWidget.verticalHeader().setVisible(False)
        wfile = None
        if copyfont or resultw:
            if not path.isdir(fd): os.mkdir(fd)
        if resultw:
            wfile = open(path.join(cpath, 'Fonts', 'fonts.txt'), mode='w', encoding='utf-8-sig')
        maxnum = 0
        for s in assfont.keys():
            ssp = s.split('?')
            if not ssp[0] == ssp[1]:
                lx = len(assfont[s][0])
                if lx > maxnum:
                    maxnum = lx
        maxnum = len(str(maxnum))
        for r in range(0, len(assfont.keys())):
            s = list(assfont.keys())[r]
            QApplication.processEvents()
            ssp = s.split('?')
            if not ssp[0] == ssp[1]:
                fp = ssp[0]
                fn = path.basename(fp)
                ann = ''
                errshow = False
                if copyfont:
                    try:
                        ui.statusbar.showMessage('正在拷贝字体 \"{0}\"'.format(fn))
                        if not path.exists(path.join(fd, fn)):
                            shutil.copy(fp, path.join(fd, fn))
                            ann = 'copied'
                        else:
                            ann = 'existed'
                    except:
                        ui.statusbar.show(sys.exc_info(), 5000)
                        ann = 'copy error'
                        errshow = True
                if resultw:
                    print('{0} <{1}> - {2}'.format(assfont[s][2], path.basename(fn), ann), file=wfile)
                if errshow:
                    ui.tableWidget.setItem(r, 2, QTWI_color(str(len(assfont[s][0])).rjust(maxnum), 255, 0, 0))
                    ui.tableWidget.setItem(r, 1, QTWI_color(path.basename(fn), 255, 0, 0))
                    ui.tableWidget.setItem(r, 0, QTWI_color(assfont[s][2], 255, 0, 0))
                else:
                    ui.tableWidget.setItem(r, 2, QTWI_color(str(len(assfont[s][0])).rjust(maxnum), 0, 255, 0))
                    ui.tableWidget.setItem(r, 1, QTWI_color(path.basename(fn), 0, 255, 0))
                    ui.tableWidget.setItem(r, 0, QTWI_color(assfont[s][2], 0, 255, 0))
                if copyfont:
                    ui.pBar.setValue(r + 1)
                    ui.tableWidget.setItem(r, 2, QTWI_color(ann, 0, 255, 0))
                # print(assfont[s][0])
            else:
                if resultw:
                    print('{0} - No Found'.format(ssp[0]), file=wfile)
                ui.tableWidget.setItem(r, 2, QTWI_color('N', 255, 0, 0))
                ui.tableWidget.setItem(r, 1, QTWI_color('No Found', 255, 0, 0))
                ui.tableWidget.setItem(r, 0, QTWI_color(ssp[0], 255, 0, 0))
            ui.tableWidget.setColumnWidth(2, 35)
        if resultw: wfile.close()
        if copyfont: ui.statusbar.showMessage('字体拷贝完成')
        print('')
    del font_info2


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


# def showMessageSubset(newasspaths: list, newfont_name: dict):
#     for ap in newasspaths:
#         if path.exists(ap):
#             print('\033[1;32m成功:\033[0m \033[1m\"{0}\"\033[0m'.format(path.basename(ap)))
#     for nf in newfont_name.keys():
#         if path.exists(newfont_name[nf][0]):
#             if newfont_name[nf][1] is None:
#                 print('\033[1;31m失败:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\"\033[0m'.format(path.basename(nf), path.basename(newfont_name[nf][0])))
#             else:
#                 print('\033[1;32m成功:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\" ({2})\033[0m'.format(path.basename(nf), path.basename(newfont_name[nf][0]), newfont_name[nf][1]))


def cFontSubset(work: int, cpath: str, spath: str = ''):
    global extlist, v_subdir, s_subdir, rmAssIn, rmAttach, fontload, ui, \
        mkvout, assout, fontout, matchStrict, no_mkvm, notfont, warningStop, errorStop, font_info
    if work in [1, 2]:
        subonly = False
        subonlyp = []
        while not path.exists(cpath):
            if path.isfile(cpath):
                testext = path.splitext(cpath)[1][1:].lower()
                if testext in extlist:
                    pass
                elif testext in ['ass', 'ssa'] and work == 1:
                    subonly = True
                else:
                    ui.statusbar.showMessage('扩展名不正确: \"{0}\"'.format(cpath), 5000)
                    return None
        # print(directout)
        medias = []
        if path.isfile(cpath):
            if not subonly:
                medias = [[path.splitext(path.basename(cpath))[0], cpath]]
            else:
                subonlyp = [cpath]
            cpath = path.dirname(cpath)
        else:
            medias = getMediaFilelist(cpath)
            if len(medias) == 0:
                if s_subdir:
                    for r, ds, fs in os.walk(spath):
                        if path.splitext(fs)[1][1:].lower() in ['ass', 'ssa']:
                            subonlyp.append(path.join(r, ds, fs))
                else:
                    for fs in os.listdir(spath):
                        if path.splitext(fs)[1][1:].lower() in ['ass', 'ssa']:
                            subonlyp.append(path.join(cpath, fs))
                if len(subonlyp) > 0 and work == 1: 
                    if QMessageBox.warning(MW, '', '''您输入的目录下只有字幕而无视频，每个字幕都将被当做单独对象处理
若是同一话有多个字幕，这话将会有多套子集化字体。
即便如此，您仍要继续吗？''', QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) == QMessageBox.StandardButton.Yes:
                        subonly = True
                    else:
                        return None
                elif work == 2:
                    ui.statusbar.showMessage('路径下找不到视频', 5000)
                else:
                    ui.statusbar.showMessage('路径下找不到字幕', 5000)
                    return None

        if assout == '':
            assout_cache = path.join(cpath, 'Subtitles')
        elif assout[0] == '?':
            assout_cache = path.join(cpath, assout[1:])
        else:
            assout_cache = assout
        if fontout == '':
            fontout_cache = path.join(cpath, 'Fonts')
        elif fontout[0] == '?':
            fontout_cache = path.join(cpath, fontout[1:])
        else:
            fontout_cache = fontout
        if mkvout == '':
            mkvout_cache = ''
        elif mkvout[0] == '?':
            mkvout_cache = path.join(cpath, mkvout[1:])
        else:
            mkvout_cache = mkvout

        domux = False
        if work == 2: domux = True

        if fontload:
            font_info2 = copy.deepcopy(font_info)
            font_info2 = templeFontLoad(cpath, font_info2)
            if path.isdir(spath):
                font_info2 = templeFontLoad(spath, font_info2)
        else:
            font_info2 = font_info
        if len(font_info2[0]) == 0:
            ui.statusbar.showMessage('程序未载入任何字体，可能需要按一下全局刷新？', 5000)
            return None
        if len(medias) > 0:
            if path.isdir(spath):
                media_ass = getSubtitles(spath, medias, True)
            elif path.splitext(spath)[1][1:].lower() in subext and path.isabs(spath):
                media_ass = {medias[0][0]: [spath]}
            else:
                media_ass = getSubtitles(cpath, medias)
            # print(media_ass)
            if len(media_ass) > 0:
                ui.pBar.setMaximum(len(media_ass.keys()))

                if domux:
                    ui.tableWidget.setRowCount(0)
                    ui.tableWidget.setColumnCount(1)
                    ui.tableWidget.setRowCount(1)
                    ui.tableWidget.insertRow(0)
                    ui.tableWidget.setHorizontalHeaderLabels(['文件名'])
                    ui.tableWidget.verticalHeader().setVisible(False)

                for r in range(0, len(media_ass.keys())):
                    QApplication.processEvents()
                    k = list(media_ass.keys())[r]
                    # print(k)
                    # print([assout_cache, fontout_cache, mkvout_cache])
                    _outdir = [assout_cache, fontout_cache, mkvout_cache]
                    newasspaths, newfont_name, mkvr = main(font_info2, media_ass[k], mux=domux, outdir=_outdir, vpath=k, asslangs=sublang, noProgress=True)
                    if mkvr != -2:
                        ui.pBar.setValue(r + 1)
                    else:
                        break
            else:
                ui.statusbar.showMessage('找不到视频对应字幕', 5000)
                return None
        elif subonly:
            ui.pBar.setMaximum(len(subonlyp))
            for i in range(0, len(subonlyp)):
                QApplication.processEvents()
                subp = subonlyp[i]
                _outdir = [assout_cache, fontout_cache, mkvout_cache]
                newasspaths, newfont_name, mkvr = main(font_info2, [subp], mux=False, outdir=_outdir, noProgress=True)
                if mkvr != -2:
                    ui.pBar.setValue(i + 1)
                else:
                    break
        del font_info2
        ui.statusbar.showMessage('子集化结束', 5000)
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
            # print('\033[1m mkvmerge 语言编码列表 \033[0m')
        else:
            print('没有检测到mkvmerge，无法输出语言编码列表')


def cLicense():
    cls()
    print('''AddSubFontMKV Python Remake 1.02 Preview9
Apache-2.0 License
Copyright(c) 2022 yyfll
依赖:
fontTools    |  MIT License
chardet      |  LGPL-2.1 License
colorama     |  BSD-3 License
mkvmerge     |  GPL-2 License
''')
    print('for more information:\nhttps://www.apache.org/licenses/')
    os.system('pause')


def cSVMatching(vPath: str, sPath: str):
    global ui, MW
    print('Subtitles-Videos Filename Matching (Experiment)')
    logfiles = []
    for f in os.listdir(sPath):
        if f[:11].lower() == 'subsrename_':
            logfiles.append(f)
    if len(logfiles) > 0:
        if QMessageBox.question(MW, '重命名撤销', '检测到重命名记录，您要撤销之前的重命名吗？',
        QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            ui.result.setRowCount(0)
            ui.result.setColumnCount(2)
            ui.result.setRowCount(1)
            ui.result.setColumnWidth(0, 260)
            ui.result.setColumnWidth(1, 260)
            ui.result.setHorizontalHeaderLabels(['新名', '原名'])
            ui.result.verticalHeader().setVisible(False)
            renCache = {}
            logfile = open(path.join(sPath, logfiles[len(logfiles) - 1]), mode='r', encoding='utf-8')
            for li in logfile.readlines():
                li = li.strip('\n').split('|')
                if len(li) < 2: continue
                renCache[path.basename(li[0])] = (path.join(path.dirname(li[0]), li[1]), li[0])
            logfile.close()
            print('')
            print('[重命名预览]\n')
            for r in renCache.keys():
                print("\033[1m\"{0}\"\033[0m".format(path.basename(renCache[r][0])))
                print('>>>> \033[33;1m\"{0}\"\033[0m\n'.format(r))
                ui.result.insertRow(0)
                ui.result.setItem(0, 0, QTableWidgetItem(r))
                ui.result.setItem(0, 1, QTableWidgetItem(path.basename(renCache[r][0])))
            print('')
            if QMessageBox.question(MW, '重命名确认', '请确认重命名后的文件名，要执行重命名吗？',
        QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                nodel = False
                ui.result.setRowCount(0)
                ui.result.setColumnCount(2)
                ui.result.setRowCount(1)
                ui.result.setColumnWidth(0, 260)
                ui.result.setColumnWidth(1, 260)
                ui.result.setHorizontalHeaderLabels(['新名', '原名'])
                ui.result.verticalHeader().setVisible(False)
                for r in renCache.values():
                    try:
                        if path.exists(r[0]):
                            os.rename(r[0], r[1])
                        else:
                            ui.result.insertRow(0)
                            ui.result.setItem(0, 0, QTWI_color('找不到文件', 0, 0, 255))
                            ui.result.setItem(0, 1, QTWI_color('No Found', 0, 0, 255))
                    except:
                        print('\033[31;1m[ERROR]\033[0m 文件\"\033[1m{0}\033[0m\"重命名失败'.format(
                            path.basename(r[0])))
                        ui.result.insertRow(0)
                        ui.result.setItem(0, 0, QTWI_color(path.basename(r[1]), 255, 0, 0))
                        ui.result.setItem(0, 1, QTWI_color(path.basename(r[0]), 255, 0, 0))
                        nodel = True
                    else:
                        print('''\033[32;1m[SUCCESS]\033[0m \"\033[1m{0}\033[0m\"
>>>> \"\033[33;1m{1}\033[0m\"'''.format(path.basename(r[0]), path.basename(r[1])))
                        ui.result.insertRow(0)
                        ui.result.setItem(0, 0, QTWI_color(path.basename(r[1]), 0, 255, 0))
                        ui.result.setItem(0, 1, QTWI_color(path.basename(r[0]), 0, 255, 0))
            if not nodel:
                for li in logfiles:
                    try:
                        os.remove(path.join(sPath, li))
                    except:
                        ui.statusbar.showMessage('删除撤销文件失败', 5000)
                        pass
        return None
    nameMatching(sPath, getMediaFilelist(vPath))


no_mkvm = False
mkvmv = ''
font_info = [{}, {}, {}, []]

# font_info 列表结构
#   [ font_name(dict), font_n_lower(dict), font_family(dict), warning_font(list) ]
def loadMain(reload: bool = False):
    global extlist, sublang, no_mkvm, dupfont, mkvmv, font_info, ui
    MW.setEnabled(False)
    del dupfont
    dupfont = {}
    del font_info
    font_info = [{}, {}, {}, []]
    # 初始化字体列表 和 mkvmerge 相关参数
    if not reload:
        if not o_fontload:
            font_list, font_name = getFileList(fontin + 
            [ui.customFontL.item(i).text() for i in range(0, ui.customFontL.count()) if path.isdir(ui.customFontL.item(i).text())])
            font_info = fontProgress(font_list +
            [[ui.customFontL.item(i).text(), '1'] for i in range(0, ui.customFontL.count()) if path.isfile(ui.customFontL.item(i).text())]
            , [font_name, {}, {}, []], f_priority)
            del font_name, font_list
        if not no_mkvm:
            print('\n\n\033[1;33m正在获取 mkvmerge 语言编码列表和支持格式列表，请稍等...\033[0m')
            langmkv = os.popen('mkvmerge --list-languages', mode='r')
            for s in langmkv.buffer.read().decode('utf-8').splitlines()[2:]:
                s = s.replace('\n', '').split('|')
                for si in range(1, len(s)):
                    ss = s[si]
                    if len(ss.strip(' ')) > 0:
                        langlist.append(ss.strip(' '))
            langmkv.close()
            if not no_extcheck:
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
                        sublang.append('und')
                        print('\033[1;31m[WARNING] 您设定的语言编码 {0} 无效，已替换为und\033[0m'.format(s))
                if len(sublang) != len(sublang_c):
                    print('\n\033[1;33m当前的语言编码列表: \"{0}\"\033[0m\n'.format(';'.join(sublang)))
                    os.system('pause')
                del sublang_c
    MW.setEnabled(True)
    # #     cls()
    # #     errorextl = []
    # #     for k in font_info[0].keys():
    # #         s = font_info[0][k]
    # #         if s[3] and (path.splitext(s[0])[1][1:].lower() in ['ttf', 'otf']):
    # #             errorextl.append((s[0], s[3]))
    # #         elif (not s[3]) and (path.splitext(s[0])[1][1:].lower() in ['ttc', 'otc']):
    # #             errorextl.append((s[0], s[3]))
    # #     if len(errorextl) > 0:
    # #         for s in errorextl:
    # #             rightext = ''
    # #             if s[1]:
    # #                 rightext = path.splitext(s[0])[1].upper()
    # #                 rightext = rightext[1:len(rightext) - 1] + 'C'
    # #             else:
    # #                 rightext = path.splitext(s[0])[1].upper()
    # #                 rightext = rightext[1:len(rightext) - 1] + 'F'
    # #             print('\033[1;31m[正确扩展名: {0}]\033[0m'.format(rightext))
    # #             print('\033[1m{0}\033[0m\n'.format(s[0]))
    # #     else:
    # #         print('没有扩展名错误的字体')
    # #     os.system('pause')

class eventW(object):
    def BlistAllClicked():
        global font_info, ui
        if not len(font_info[0]) > 0: 
            ui.statusbar.showMessage('您需要进行全局刷新才能显示字体列表', 5000)
            return None
        ui.tableWidget.setRowCount(0)
        ui.tableWidget.setColumnCount(3)
        ui.tableWidget.setColumnWidth(2, 30)
        ui.tableWidget.setRowCount(len(font_info[0]))
        ui.tableWidget.setHorizontalHeaderLabels(['文件名', '名称', '索引'])
        ui.tableWidget.verticalHeader().setVisible(False)
        fkeys = list(font_info[0].keys())
        for i in range(-1, len(fkeys)):
            finfo = font_info[0][fkeys[i]]
            if i > 0:
                ii = i
                while ui.tableWidget.item(ii, 0) is None:
                    ii -= 1
                if ui.tableWidget.item(ii, 0).text().lower() == path.basename(finfo[0]).lower():
                    ui.tableWidget.setSpan(ii, 0, i - ii + 2, 1)
                else:
                    ui.tableWidget.setItem(i + 1, 0, QtWidgets.QTableWidgetItem(path.basename(finfo[0])))
            else: 
                ui.tableWidget.setItem(i + 1, 0, QtWidgets.QTableWidgetItem(path.basename(finfo[0])))
            ui.tableWidget.setItem(i + 1, 1, QtWidgets.QTableWidgetItem(fkeys[i]))
            ui.tableWidget.setItem(i + 1, 2, QtWidgets.QTableWidgetItem(str(finfo[1])))


    def BdupFontClicked():
        global ui, dupfont
        if len(dupfont.keys()) > 0:
            ui.tableWidget.setRowCount(0)
            ui.tableWidget.setColumnCount(2)
            ui.tableWidget.setRowCount(len(dupfont.keys()))
            ui.tableWidget.setHorizontalHeaderLabels(['名称', '文件名'])
            ui.tableWidget.verticalHeader().setVisible(False)
            ui.tableWidget.setWordWrap(True)
            i = 0
            for s in dupfont.keys():
                si = i
                for ss in dupfont[s]:
                    if i > ui.tableWidget.rowCount() - 1: ui.tableWidget.insertRow(i)
                    ui.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(ss))
                    i += 1
                while si + len(dupfont[s]) >= ui.tableWidget.rowCount():
                    ui.tableWidget.insertRow(i)
                if len(dupfont[s]) > 1:
                    ui.tableWidget.setSpan(si, 1, len(dupfont[s]), 1)
                else:
                    if si + 1 > ui.tableWidget.rowCount():
                        ui.tableWidget.insertRow(i)
                    ui.tableWidget.setSpan(si, 0, 2, 1)
                    ui.tableWidget.setSpan(si, 1, 2, 1)
                    i += 1
                ui.tableWidget.setItem(si, 1, QtWidgets.QTableWidgetItem('\n'.join([path.basename(sn) for sn in s])))
        else:
            ui.statusbar.showMessage('没有重复字体')

    def BoldFontClicked():
        if len(font_info[3]) > 0:
            oldflist = {}
            for s in font_info[3]:
                oldfname = font_info[1].get(s)
                oldfpath = font_info[0].get(oldfname)
                if oldfpath is None: continue
                else: oldfpath = oldfpath[0]
                if oldflist.get(oldfpath) is None:
                    oldflist[oldfpath] = [oldfname]
                else:
                    oldflist[oldfpath].append(oldfname)
            ui.tableWidget.setRowCount(0)
            ui.tableWidget.setColumnCount(2)
            ui.tableWidget.setRowCount(len(oldflist.keys()))
            ui.tableWidget.setHorizontalHeaderLabels(['文件名', '名称'])
            ui.tableWidget.verticalHeader().setVisible(False)
            ui.tableWidget.setWordWrap(True)
            i = 0
            for s in oldflist.keys():
                si = i
                for ss in oldflist[s]:
                    if i > ui.tableWidget.rowCount() - 1: ui.tableWidget.insertRow(i)
                    ui.tableWidget.setItem(i, 1, QtWidgets.QTableWidgetItem(ss))
                    i += 1
                while si + len(oldflist[s]) >= ui.tableWidget.rowCount():
                    ui.tableWidget.insertRow(i)
                if len(oldflist[s]) > 1:
                    ui.tableWidget.setSpan(si, 0, len(oldflist[s]), 1)
                else:
                    if si + 1 > ui.tableWidget.rowCount():
                        ui.tableWidget.insertRow(i)
                    ui.tableWidget.setSpan(si, 0, 2, 1)
                    ui.tableWidget.setSpan(si, 1, 2, 1)
                    i += 1
                ui.tableWidget.setItem(si, 0, QtWidgets.QTableWidgetItem(path.basename(s)))
        else:
            ui.statusbar.showMessage('没有使用旧格式的字体', 5000)


    def checkFontSubClicked():
        global ui
        _path = ui.sPathL.text().strip()
        if path.exists(_path):
            if path.isabs(_path):
                MW.setEnabled(False)
                cListAssFont(_path)
                MW.setEnabled(True)
            else:
                ui.statusbar.showMessage('字幕路径必须是绝对路径', 5000)
        else:
            _path = ui.vPathL.text().strip()
            if path.isdir(_path):
                MW.setEnabled(False)
                cListAssFont(_path)
                MW.setEnabled(True)
            else:
                ui.statusbar.showMessage('字幕路径不存在', 5000)


    def checkDir(_path: str) -> str:
        _path = _path.strip()
        if len(_path) == 0:
            return _path
        elif _path[0] == '?':
            return _path[0] + re.sub(cillegal, '_', _path[1:])
        elif path.isdir(_path):
            return _path
        elif path.isdir(path.dirname(_path)):
            return _path
        else:
            return None


    def callMain(work: int):
        global ui, mkvout, assout, fontout
        _inpath = ui.vPathL.text().strip()
        if len(_inpath) == 0 and work != 2:
            _inpath = ui.sPathL.text().strip()
        if not path.exists(_inpath):
            ui.statusbar.showMessage('输入路径无效', 5000)
            return None
        _vpath = eventW.checkDir(ui.OvDir.text().strip(' '))
        _spath = eventW.checkDir(ui.OsDir.text().strip(' '))
        _fpath = eventW.checkDir(ui.OfDir.text().strip(' '))
        if _vpath is None or _spath is None or _fpath is None:
            return None
        else:
            mkvout, assout, fontout = _vpath, _spath, _fpath
            MW.setEnabled(False)
            if ui.sPathL.text().strip() != _inpath:
                cFontSubset(work, _inpath, ui.sPathL.text().strip())
            else:
                cFontSubset(work, _inpath)
            MW.setEnabled(True)

    def subsetClicked():
        eventW.callMain(1)


    def muxClicked():
        eventW.callMain(2)


    def tryMatchClicked():
        global ui
        _inpath = ui.vPathL.text().strip()
        _sipath = ui.sPathL.text().strip()
        if path.isdir(_inpath) and path.isdir(_sipath):
            cSVMatching(_inpath, _sipath)
        else:
            if not path.isdir(_inpath):
                ui.statusbar.showMessage('输入视频路径无效，它必须是一个包含视频的目录', 5000)
            else:
                ui.statusbar.showMessage('输入字幕路径无效，它必须是一个包含字幕的目录', 5000)


    def LaddClicked():
        global ui
        _text = re.sub(cillegal, '_', ui.addAnn.text().strip().lstrip('.'))
        if len(_text) > 0:
            for i in range(0, ui.list.count()):
                s = ui.list.item(i).text()
                if s.lower() == _text.lower():
                    ui.list.item(i).setText(_text)
                    return None
            ui.list.addItem(_text)


    def LdelClicked():
        global ui
        si = ui.list.currentIndex()
        ui.list.takeItem(si.row())


    def DaddClicked():
        global ui
        _text = ui.addCustomF.text().strip()
        if len(_text) > 0:
            if path.isfile(_text):
                if path.splitext(_text)[1][1:].lower() not in ['ttf', 'ttc', 'otc', 'otf']:
                    ui.statusbar.showMessage('添加的应是可接受的字体文件而非\"{0}\"文件'.format(path.splitext(_text)[1][1:].upper()))
                    return None
            for i in range(0, ui.customFontL.count()):
                s = ui.customFontL.item(i).text()
                if s.lower() == _text.lower():
                    ui.customFontL.item(i).setText(_text)
                    return None
            ui.customFontL.addItem(_text)


    def DdelClicked():
        global ui
        si = ui.customFontL.currentIndex()
        ui.customFontL.takeItem(si.row())


    def Bw2FileClicked():
        global resultw, ui
        resultw = ui.Bw2File.isChecked()


    def BsSubdirClicked():
        global s_subdir, ui
        if s_subdir:
            s_subdir = not ui.BsSubdir1.isChecked() and not ui.BsSubdir2.isChecked() and not ui.BsSubdir3.isChecked()
        else:
            s_subdir = not ui.BsSubdir1.isChecked() or not ui.BsSubdir2.isChecked() or not ui.BsSubdir3.isChecked()
        ui.BsSubdir1.setChecked(s_subdir)
        ui.BsSubdir2.setChecked(s_subdir)
        ui.BsSubdir3.setChecked(s_subdir)


    def BvSubdirClicked():
        global v_subdir, ui
        if v_subdir:
            v_subdir = not ui.BvSubdir2.isChecked() and not ui.BvSubdir3.isChecked()
        else:
            v_subdir = not ui.BvSubdir2.isChecked() or not ui.BvSubdir3.isChecked()
        ui.BvSubdir2.setChecked(v_subdir)
        ui.BvSubdir3.setChecked(v_subdir)


    def BstrictSClicked():
        global matchStrict, ui
        if matchStrict:
            matchStrict = not ui.BstrictS2.isChecked() and not ui.BstrictS3.isChecked()
        else:
            matchStrict = not ui.BstrictS2.isChecked() or not ui.BstrictS3.isChecked()
        ui.BstrictS2.setChecked(matchStrict)
        ui.BstrictS3.setChecked(matchStrict)


    def BcopyFontClicked():
        global copyfont, ui
        copyfont = ui.BcopyFont.isChecked()


    def BerrorStopClicked():
        global errorStop, ui
        errorStop = ui.BerrorStop.isChecked()


    def BwarningStopClicked():
        global warningStop, ui
        warningStop = ui.BwarningStop.isChecked()


    def BremoveFontClicked():
        global rmAttach, ui
        rmAttach = ui.BremoveFont.isChecked()


    def BremoveSubClicked():
        global rmAssIn, ui
        rmAssIn = ui.BremoveSub.isChecked()


    def BnotFontClicked():
        global notfont, ui
        notfont = ui.BnotFont.isChecked()


    def BworkDir1Clicked():
        global fontload, s_fontload, ui
        ui.BworkDir2.setCheckState(0)
        if ui.BworkDir1.checkState() == 2:
            s_fontload = True
            ui.BworkDir2.setCheckState(2)
        elif ui.BworkDir1.isChecked():
            s_fontload = False
            ui.BworkDir2.setCheckState(1)
        fontload = ui.BworkDir1.isChecked()
        
        
    def BworkDir2Clicked():
        global fontload, s_fontload, ui
        ui.BworkDir1.setCheckState(0)
        if ui.BworkDir2.checkState() == 2:
            s_fontload = True
            ui.BworkDir1.setCheckState(2)
        elif ui.BworkDir2.isChecked():
            s_fontload = False
            ui.BworkDir1.setCheckState(1)
        fontload = ui.BworkDir2.isChecked()


class dropQLineEdit(QtWidgets.QLineEdit):
    def __init__(self, *args, **kwargs):
        super(dropQLineEdit, self).__init__(*args, **kwargs)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e):
        global ui, extlist, subext
        if len(e.mimeData().urls()) > 0 and e.mimeData().urls()[0].isLocalFile():
            _path = e.mimeData().urls()[0].toString().replace('file:', '').replace('///', '')
            if path.exists(_path):
                if path.isdir(_path):
                    e.acceptProposedAction()
                elif self.objectName() in ['vPathL', 'sPathL', 'addCustomF']:
                    ext = path.splitext(_path)[1][1:].lower()
                    if self.objectName() == 'vPathL':
                        if ext in extlist:
                            e.acceptProposedAction()
                        else:
                            ui.statusbar.showMessage('您应该拖入一个可接受的视频文件而非\"{0}\"文件'.format(ext.upper()), 5000)
                    elif self.objectName() == 'sPathL':
                        if ext in subext:
                            e.acceptProposedAction()
                        else:
                            ui.statusbar.showMessage('您应该拖入一个可接受的字幕文件而非\"{0}\"文件'.format(ext.upper()), 5000)
                    elif self.objectName() == 'addCustomF':
                        if ext in ['ttf', 'otf', 'ttc', 'otc']:
                            e.acceptProposedAction()
                        else:
                            ui.statusbar.showMessage('您应该拖入一个可接受的字体文件而非\"{0}\"文件'.format(ext.upper()), 5000)
                else:
                    ui.statusbar.showMessage('您应该拖入一个目录而不是一个文件', 5000)
            else:
                ui.statusbar.showMessage('路径不存在', 5000)

    def dropEvent(self, e):
        _path = e.mimeData().urls()[0].toString().replace('file:', '').replace('///', '')
        self.setText(_path)
        e.acceptProposedAction()


class Ui_MW(object):
    def setupUi(self, MW):
        MW.setObjectName("MW")
        MW.resize(800, 360)
        MW.setDockOptions(QtWidgets.QMainWindow.AllowTabbedDocks|QtWidgets.QMainWindow.AnimatedDocks)
        MW.setFixedSize(MW.width(), MW.height())
        MW.setAcceptDrops(True)
        QToolTip.setFont(QtGui.QFont('黑体', 10, 700))
        self.centralwidget = QWidget(MW)
        self.centralwidget.setObjectName(u"centralwidget")
        self.tabWidget = QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tabWidget.setGeometry(QRect(6, 109, 571, 231))
        self.laf = QWidget()
        self.laf.setObjectName(u"laf")
        self.opt = QGroupBox(self.laf)
        self.opt.setObjectName(u"opt")
        self.opt.setGeometry(QRect(10, 10, 161, 101))
        self.Bw2File = QCheckBox(self.opt)
        self.Bw2File.setObjectName(u"Bw2File")
        self.Bw2File.setGeometry(QRect(10, 20, 131, 16))
        self.BcopyFont = QCheckBox(self.opt)
        self.BcopyFont.setObjectName(u"BcopyFont")
        self.BcopyFont.setGeometry(QRect(10, 40, 131, 16))
        self.BsSubdir1 = QCheckBox(self.opt)
        self.BsSubdir1.setObjectName(u"BsSubdir1")
        self.BsSubdir1.setGeometry(QRect(10, 60, 131, 16))
        self.BworkDir1 = QCheckBox(self.opt)
        self.BworkDir1.setObjectName(u"BworkDir1")
        self.BworkDir1.setGeometry(QRect(10, 80, 141, 16))
        self.BworkDir1.setTristate(True)
        self.listall = QPushButton(self.laf)
        self.listall.setObjectName(u"listall")
        self.listall.setGeometry(QRect(10, 150, 161, 21))
        self.checkFontSub = QPushButton(self.laf)
        self.checkFontSub.setObjectName(u"checkFontSub")
        self.checkFontSub.setGeometry(QRect(10, 180, 161, 21))
        self.tableWidget = QTableWidget(self.laf)
        self.tableWidget.setObjectName(u"tableWidget")
        self.tableWidget.setGeometry(QRect(180, 10, 381, 191))
        font = QtGui.QFont()
        font.setFamily(u"\u9ed1\u4f53")
        self.tableWidget.setFont(font)
        self.BdupFont = QPushButton(self.laf)
        self.BdupFont.setObjectName(u"BdupFont")
        self.BdupFont.setGeometry(QRect(10, 120, 71, 21))
        self.BoldFont = QPushButton(self.laf)
        self.BoldFont.setObjectName(u"BoldFont")
        self.BoldFont.setGeometry(QRect(100, 120, 71, 21))
        self.tabWidget.addTab(self.laf, "")
        self.ssm = QWidget()
        self.ssm.setObjectName(u"ssm")
        self.opt2 = QGroupBox(self.ssm)
        self.opt2.setObjectName(u"opt2")
        self.opt2.setGeometry(QRect(10, 10, 411, 101))
        self.optmux = QGroupBox(self.opt2)
        self.optmux.setObjectName(u"optmux")
        self.optmux.setGeometry(QRect(280, 10, 121, 81))
        self.BremoveSub = QCheckBox(self.optmux)
        self.BremoveSub.setObjectName(u"BremoveSub")
        self.BremoveSub.setGeometry(QRect(10, 20, 121, 16))
        self.BremoveFont = QCheckBox(self.optmux)
        self.BremoveFont.setObjectName(u"BremoveFont")
        self.BremoveFont.setGeometry(QRect(10, 40, 121, 16))
        self.BnotFont = QCheckBox(self.optmux)
        self.BnotFont.setObjectName(u"BnotFont")
        self.BnotFont.setGeometry(QRect(10, 60, 121, 16))
        self.optall = QGroupBox(self.opt2)
        self.optall.setObjectName(u"optall")
        self.optall.setGeometry(QRect(10, 10, 271, 81))
        self.BvSubdir2 = QCheckBox(self.optall)
        self.BvSubdir2.setObjectName(u"BvSubdir2")
        self.BvSubdir2.setGeometry(QRect(0, 20, 131, 16))
        self.BsSubdir2 = QCheckBox(self.optall)
        self.BsSubdir2.setObjectName(u"BsSubdir2")
        self.BsSubdir2.setGeometry(QRect(0, 40, 131, 16))
        self.BerrorStop = QCheckBox(self.optall)
        self.BerrorStop.setObjectName(u"BerrorStop")
        self.BerrorStop.setGeometry(QRect(140, 40, 131, 16))
        self.BworkDir2 = QCheckBox(self.optall)
        self.BworkDir2.setObjectName(u"BworkDir2")
        self.BworkDir2.setGeometry(QRect(140, 60, 131, 16))
        self.BworkDir2.setTristate(True)
        self.BstrictS2 = QCheckBox(self.optall)
        self.BstrictS2.setObjectName(u"BstrictS2")
        self.BstrictS2.setGeometry(QRect(0, 60, 131, 16))
        self.BwarningStop = QCheckBox(self.optall)
        self.BwarningStop.setObjectName(u"BwarningStop")
        self.BwarningStop.setGeometry(QRect(140, 20, 131, 16))
        self.outdir = QGroupBox(self.ssm)
        self.outdir.setObjectName(u"outdir")
        self.outdir.setGeometry(QRect(10, 110, 551, 91))
        self.OvDir = dropQLineEdit(self.outdir)
        self.OvDir.setObjectName(u"OvDir")
        self.OvDir.setGeometry(QRect(40, 20, 501, 20))
        self.OsDir = dropQLineEdit(self.outdir)
        self.OsDir.setObjectName(u"OsDir")
        self.OsDir.setGeometry(QRect(40, 40, 501, 20))
        self.OfDir = dropQLineEdit(self.outdir)
        self.OfDir.setObjectName(u"OfDir")
        self.OfDir.setGeometry(QRect(40, 60, 501, 20))
        self.label = QLabel(self.outdir)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 20, 31, 21))
        self.label_2 = QLabel(self.outdir)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(10, 40, 31, 21))
        self.label_2.setFrameShape(QFrame.NoFrame)
        self.label_3 = QLabel(self.outdir)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setGeometry(QRect(10, 60, 31, 21))
        self.label_3.setFrameShape(QFrame.NoFrame)
        self.subset = QPushButton(self.ssm)
        self.subset.setObjectName(u"subset")
        self.subset.setGeometry(QRect(430, 20, 131, 41))
        self.mux = QPushButton(self.ssm)
        self.mux.setObjectName(u"mux")
        self.mux.setGeometry(QRect(430, 70, 131, 41))
        self.tabWidget.addTab(self.ssm, "")
        self.vsm = QWidget()
        self.vsm.setObjectName(u"vsm")
        self.opt3 = QGroupBox(self.vsm)
        self.opt3.setObjectName(u"opt3")
        self.opt3.setGeometry(QRect(10, 10, 151, 91))
        self.BsSubdir3 = QCheckBox(self.opt3)
        self.BsSubdir3.setObjectName(u"BsSubdir3")
        self.BsSubdir3.setGeometry(QRect(10, 40, 131, 16))
        self.BvSubdir3 = QCheckBox(self.opt3)
        self.BvSubdir3.setObjectName(u"BvSubdir3")
        self.BvSubdir3.setGeometry(QRect(10, 20, 131, 16))
        self.BstrictS3 = QCheckBox(self.opt3)
        self.BstrictS3.setObjectName(u"BstrictS3")
        self.BstrictS3.setGeometry(QRect(10, 60, 131, 16))
        self.customRule = QGroupBox(self.vsm)
        self.customRule.setObjectName(u"customRule")
        self.customRule.setGeometry(QRect(310, 10, 251, 71))
        self.text = QLineEdit(self.customRule)
        self.text.setObjectName(u"text")
        self.text.setGeometry(QRect(40, 20, 201, 20))
        self.text.setAcceptDrops(False)
        self.text2 = QLineEdit(self.customRule)
        self.text2.setObjectName(u"text2")
        self.text2.setGeometry(QRect(40, 40, 201, 20))
        self.text2.setAcceptDrops(False)
        self.label_4 = QLabel(self.customRule)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setGeometry(QRect(10, 40, 31, 21))
        self.label_4.setFrameShape(QFrame.NoFrame)
        self.label_5 = QLabel(self.customRule)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(10, 20, 31, 21))
        self.tryMatch = QPushButton(self.vsm)
        self.tryMatch.setObjectName(u"tryMatch")
        self.tryMatch.setGeometry(QRect(310, 81, 251, 20))
        self.result = QTableWidget(self.vsm)
        self.result.setObjectName(u"result")
        self.result.setGeometry(QRect(10, 110, 551, 91))
        self.result.setFont(font)
        self.customAnn = QGroupBox(self.vsm)
        self.customAnn.setObjectName(u"customAnn")
        self.customAnn.setGeometry(QRect(170, 10, 131, 91))
        self.list = QListWidget(self.customAnn)
        self.list.setObjectName(u"list")
        self.list.setGeometry(QRect(5, 10, 121, 41))
        self.list.setFont(font)
        self.Ldel = QPushButton(self.customAnn)
        self.Ldel.setObjectName(u"Ldel")
        self.Ldel.setGeometry(QRect(70, 71, 41, 20))
        self.Ladd = QPushButton(self.customAnn)
        self.Ladd.setObjectName(u"Ladd")
        self.Ladd.setGeometry(QRect(20, 71, 41, 20))
        self.addAnn = QLineEdit(self.customAnn)
        self.addAnn.setObjectName(u"addAnn")
        self.addAnn.setGeometry(QRect(5, 50, 121, 20))
        self.tabWidget.addTab(self.vsm, "")
        self.sPathG = QGroupBox(self.centralwidget)
        self.sPathG.setObjectName(u"sPathG")
        self.sPathG.setGeometry(QRect(10, 50, 781, 51))
        self.sPathL = dropQLineEdit(self.sPathG)
        self.sPathL.setObjectName(u"sPathL")
        self.sPathL.setGeometry(QRect(10, 20, 761, 20))
        self.vPathG = QGroupBox(self.centralwidget)
        self.vPathG.setObjectName(u"vPathG")
        self.vPathG.setGeometry(QRect(10, 0, 781, 51))
        self.vPathL = dropQLineEdit(self.vPathG)
        self.vPathL.setObjectName(u"vPathL")
        self.vPathL.setGeometry(QRect(10, 20, 761, 20))
        self.customFontD = QGroupBox(self.centralwidget)
        self.customFontD.setObjectName(u"customFontD")
        self.customFontD.setGeometry(QRect(580, 110, 211, 191))
        self.customFontL = QListWidget(self.customFontD)
        self.customFontL.setObjectName(u"customFontL")
        self.customFontL.setGeometry(QRect(10, 20, 191, 111))
        self.customFontL.setFont(font)
        self.customFontL.setAcceptDrops(True)
        self.refresh = QPushButton(self.customFontD)
        self.refresh.setObjectName(u"refresh")
        self.refresh.setGeometry(QRect(10, 160, 91, 23))
        self.Ddel = QPushButton(self.customFontD)
        self.Ddel.setObjectName(u"Ddel")
        self.Ddel.setGeometry(QRect(160, 160, 41, 23))
        self.Dadd = QPushButton(self.customFontD)
        self.Dadd.setObjectName(u"Dadd")
        self.Dadd.setGeometry(QRect(110, 160, 41, 23))
        self.addCustomF = dropQLineEdit(self.customFontD)
        self.addCustomF.setObjectName(u"addCustomF")
        self.addCustomF.setGeometry(QRect(10, 130, 191, 20))
        self.pBar = QProgressBar(self.centralwidget)
        self.pBar.setObjectName(u"pBar")
        self.pBar.setGeometry(QRect(580, 310, 211, 23))
        self.pBar.setValue(0)
        self.pBar.setTextVisible(True)
        self.pBar.setOrientation(QtCore.Qt.Horizontal)
        self.pBar.setTextDirection(QProgressBar.TopToBottom)
        MW.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MW)
        self.statusbar.setObjectName(u"statusbar")
        MW.setStatusBar(self.statusbar)

        self.BdupFont.clicked.connect(eventW.BdupFontClicked)
        self.BoldFont.clicked.connect(eventW.BoldFontClicked)
        self.Bw2File.clicked.connect(eventW.Bw2FileClicked)
        self.BcopyFont.clicked.connect(eventW.BcopyFontClicked)
        self.BsSubdir1.clicked.connect(eventW.BsSubdirClicked)
        self.BworkDir1.clicked.connect(eventW.BworkDir1Clicked)
        self.listall.clicked.connect(eventW.BlistAllClicked)
        self.checkFontSub.clicked.connect(eventW.checkFontSubClicked)
        self.BremoveSub.clicked.connect(eventW.BremoveSubClicked)
        self.BremoveFont.clicked.connect(eventW.BremoveFontClicked)
        self.BnotFont.clicked.connect(eventW.BnotFontClicked)
        self.BvSubdir2.clicked.connect(eventW.BvSubdirClicked)
        self.BsSubdir2.clicked.connect(eventW.BsSubdirClicked)
        self.BerrorStop.clicked.connect(eventW.BerrorStopClicked)
        self.BworkDir2.clicked.connect(eventW.BworkDir2Clicked)
        self.BstrictS2.clicked.connect(eventW.BstrictSClicked)
        self.BwarningStop.clicked.connect(eventW.BwarningStopClicked)
        self.subset.clicked.connect(eventW.subsetClicked)
        self.mux.clicked.connect(eventW.muxClicked)
        self.BsSubdir3.clicked.connect(eventW.BsSubdirClicked)
        self.BvSubdir3.clicked.connect(eventW.BvSubdirClicked)
        self.BstrictS3.clicked.connect(eventW.BstrictSClicked)
        self.tryMatch.clicked.connect(eventW.tryMatchClicked)
        self.Ldel.clicked.connect(eventW.LdelClicked)
        self.Ladd.clicked.connect(eventW.LaddClicked)
        self.refresh.clicked.connect(loadMain)
        self.Ddel.clicked.connect(eventW.DdelClicked)
        self.Dadd.clicked.connect(eventW.DaddClicked)

        self.retranslateUi(MW)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MW)

    def retranslateUi(self, MW):
        _translate = QtCore.QCoreApplication.translate
        MW.setWindowTitle(_translate("MW", "ASFMKV Python Remake GUI-Beta | © 2022 yyfll | Apache-2.0"))
        self.opt.setTitle(_translate("MW", "选项"))
        self.Bw2File.setText(_translate("MW", "将结果写入文件"))
        self.BcopyFont.setText(_translate("MW", "复制字幕所需字体"))
        self.BsSubdir1.setText(_translate("MW", "搜索子目录中的字幕"))
        self.BworkDir1.setText(_translate("MW", "使用工作目录下的字体"))
        self.BdupFont.setText(_translate("MW", u"\u91cd\u590d\u5b57\u4f53", None))
        self.BoldFont.setText(_translate("MW", u"\u65e7\u683c\u5f0f\u5b57\u4f53", None))
        self.listall.setText(_translate("MW", "列出系统及自定义库全部字体"))
        self.checkFontSub.setText(_translate("MW", "检查字幕所需字体"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.laf), _translate("MW", "字幕字体检查与列出系统字体"))
        self.opt2.setTitle(_translate("MW", "选项"))
        self.optmux.setTitle(_translate("MW", "封装"))
        self.BremoveSub.setText(_translate("MW", "移除内挂字幕"))
        self.BremoveFont.setText(_translate("MW", "移除原有附件"))
        self.BnotFont.setText(_translate("MW", "不封装字体"))
        self.optall.setTitle(_translate("MW", "全局"))
        self.BvSubdir2.setText(_translate("MW", "搜索子目录中的视频"))
        self.BsSubdir2.setText(_translate("MW", "搜索子目录中的字幕"))
        self.BerrorStop.setText(_translate("MW", "字体子集化失败中断"))
        self.BworkDir2.setText(_translate("MW", "使用工作目录字体"))
        self.BstrictS2.setText(_translate("MW", "严格字幕匹配"))
        self.BwarningStop.setText(_translate("MW", "猜测旧格式字体中断"))
        self.outdir.setTitle(_translate("MW", "输出目录设置"))
        self.label.setText(_translate("MW", "视频"))
        self.label_2.setText(_translate("MW", "字幕"))
        self.label_3.setText(_translate("MW", "字体"))
        self.subset.setText(_translate("MW", "字体子集化"))
        self.mux.setText(_translate("MW", "子集化并封装"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.ssm), _translate("MW", "字体子集化与封装"))
        self.opt3.setTitle(_translate("MW", "选项"))
        self.BsSubdir3.setText(_translate("MW", "搜索子目录中的字幕"))
        self.BvSubdir3.setText(_translate("MW", "搜索子目录中的视频"))
        self.BstrictS3.setText(_translate("MW", "严格字幕搜索"))
        self.customRule.setTitle(_translate("MW", "自定义规则"))
        self.tryMatch.setText(_translate("MW", "尝试匹配"))
        self.customAnn.setTitle(_translate("MW", "自定义后缀"))
        self.Ldel.setText(_translate("MW", "删除"))
        self.Ladd.setText(_translate("MW", "添加"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.vsm), _translate("MW", "视频-字幕名称匹配"))
        self.sPathG.setTitle(_translate("MW", "字幕目录"))
        self.vPathG.setTitle(_translate("MW", "视频目录"))
        self.customFontD.setTitle(_translate("MW", "自定义字体目录"))
        self.refresh.setText(_translate("MW", "全局刷新"))
        self.Ddel.setText(_translate("MW", "删除"))
        self.Dadd.setText(_translate("MW", "添加"))
        self.label_4.setText(_translate("MW", "视频"))
        self.label_5.setText(_translate("MW", "字幕"))
        self.addCustomF.setToolTip(_translate("MW", '自定义字体，可以是包含字体的文件夹（强制搜索子目录），也可以是单个字体文件'))
        self.text2.setToolTip(_translate("MW", '''字幕文件搜索规则
请将字幕文件中的集数部分替换为半角冒号\":\"
如名为: [DMG&Mabors][Uchouten_Kazoku_S2][01][720P_Hi10P][AVC_AAC](1D26875F).sc.ass
替换成: [DMG&Mabors][Uchouten_Kazoku_S2][:][720P_Hi10P][AVC_AAC](1D26875F).sc.ass
要注意必须保留\":\"后一个字符，理论上再往后的不输都可以（未测试）'''))
        self.text.setToolTip(_translate("MW", '''视频文件搜索规则
请将视频文件中的集数部分替换为半角冒号\":\"
如名为: [Nekomoe kissaten] DECA-DENCE 01 [WebRip 1080p HEVC-10bit AAC].mkv
替换成: [Nekomoe kissaten] DECA-DENCE : [WebRip 1080p HEVC-10bit AAC].mkv
要注意必须保留\":\"后一个字符，理论上再往后的不输都可以（未测试）'''))
        self.BstrictS2.setToolTip(_translate("MW", '''是否启用相对严格的字幕匹配
√: 字幕文件名必须从头开始与视频文件名匹配
□: 只要字幕文件名中有视频文件名即可，不管在什么位置'''))
        self.BvSubdir3.setToolTip(_translate("MW", '''在搜索视频时不仅只搜索输入目录的顶层目录
同时搜索该目录下的所有子目录'''))
        self.BsSubdir3.setToolTip(_translate("MW", '''在搜索字幕时不仅只搜索输入目录的顶层目录
（已自动包含 \\subs 和 \\subtitles）
同时搜索该目录下的所有子目录'''))
        self.OfDir.setToolTip(_translate("MW", '输出字体的目录\n可以用写在最前的半角问号\"?\"来代替工作目录，之后的字符全部认为是文件夹名'))
        self.OsDir.setToolTip(_translate("MW", '输出字幕的目录\n可以用写在最前的半角问号\"?\"来代替工作目录，之后的字符全部认为是文件夹名'))
        self.OvDir.setToolTip(_translate("MW", '输出视频的目录\n可以用写在最前的半角问号\"?\"来代替工作目录，之后的字符全部认为是文件夹名'))
        self.BwarningStop.setToolTip(_translate("MW", '当试图子集化遇到读取时需要修正的旧格式字体（极有可能子集化失败）时中断任务\n您可以在前一个页面查看这些字体'))
        self.BstrictS2.setToolTip(_translate("MW", '''是否启用相对严格的字幕匹配
√: 字幕文件名必须从头开始与视频文件名匹配
□: 只要字幕文件名中有视频文件名即可，不管在什么位置'''))
        self.BworkDir2.setToolTip(_translate("MW", '''设置您是否使用工作目录下的字体
□: 不使用
■: 仅顶级目录和文件夹名包含\"Font\"的次一级目录
√: 搜索全部子目录中的字体'''))
        self.BerrorStop.setToolTip(_translate("MW", '在子集化错误的时候是否终止任务'))
        self.BsSubdir2.setToolTip(_translate("MW", '''在搜索字幕时不仅只搜索输入目录的顶层目录
（已自动包含 \\subs 和 \\subtitles）
同时搜索该目录下的所有子目录'''))
        self.BvSubdir2.setToolTip(_translate("MW", '''在搜索视频时不仅只搜索输入目录的顶层目录
同时搜索该目录下的所有子目录'''))
        self.BnotFont.setToolTip(_translate("MW", '是否只封装字幕，而不子集化和封装字体'))
        self.BremoveFont.setToolTip(_translate("MW", '是否移除MKV文件中的附件（主要是字体）'))
        self.BremoveSub.setToolTip(_translate("MW", '是否移除原本的文件中内挂的字幕\n（如MKV文件中的ASS、MP4文件中的SRT等）'))
        self.BworkDir1.setToolTip(_translate("MW", '''设置您是否使用工作目录下的字体
□: 不使用
■: 仅顶级目录和文件夹名包含\"Font\"的次一级目录
√: 搜索全部子目录中的字体'''))
        self.BsSubdir1.setToolTip(_translate("MW", '''在搜索字幕时不仅只搜索输入目录的顶层目录
（已自动包含 \\subs 和 \\subtitles）
同时搜索该目录下的所有子目录'''))
        self.BcopyFont.setToolTip(_translate("MW", '拷贝所需字体到\n\"工作目录\\Fonts\"'))
        self.Bw2File.setToolTip(_translate("MW", '将字体列表写入到\n\"工作目录\\Fonts\\fonts.txt\"'))
        self.BoldFont.setToolTip(_translate("MW", '显示您系统中格式太旧，可能导致 fontTools 子集化失败的字体'))
        self.BdupFont.setToolTip(_translate("MW", '''显示您系统中重复的字体，满足以下条件的会被判断为重复字体
1. 字体存在于系统字体文件夹中
2. 字体内部的名称相同
3. 字体文件名不同或存放于不同的系统字体文件夹中'''))

    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)


# try:
#     loadMain()
#     while 1:
#         loadMain(True)
# except Exception as e:
#     print('\n\033[31;1m')
#     traceback.print_exc()
#     print('\033[0m\n')
#     if os.system('choice /M 您要重新启动本程序吗？') == 1:
#         os.system('start cmd /c py \"{0}\"'.format(path.realpath(__file__)))
#     else:
#         os.system('pause')
if __name__ == '__main__':
    app = QApplication(sys.argv)
    MW = QMainWindow()
    ui = Ui_MW()
    ui.setupUi(MW)

    if os.system('mkvmerge -V 1>nul 2>nul') > 0:
        no_mkvm = True
        ui.mux.setEnabled(False)
    ui.Bw2File.setChecked(resultw)
    ui.BcopyFont.setChecked(copyfont)
    ui.BsSubdir1.setChecked(s_subdir)
    ui.BsSubdir2.setChecked(s_subdir)
    ui.BsSubdir3.setChecked(s_subdir)
    ui.BvSubdir2.setChecked(v_subdir)
    ui.BvSubdir3.setChecked(v_subdir)
    ui.BstrictS2.setChecked(matchStrict)
    ui.BstrictS3.setChecked(matchStrict)
    cs = 1
    if fontload:
        if s_fontload:
            cs = 2
    else:
        cs = 0
    ui.BworkDir1.setCheckState(cs)
    ui.BworkDir2.setCheckState(cs)
    del cs
    ui.BremoveFont.setChecked(rmAttach)
    ui.BremoveSub.setChecked(rmAssIn)
    ui.BnotFont.setChecked(notfont)
    ui.BerrorStop.setChecked(errorStop)
    ui.BwarningStop.setChecked(warningStop)
    ui.OvDir.setText(mkvout)
    ui.OsDir.setText(assout)
    ui.OfDir.setText(fontout)

    MW.show()
    QMessageBox.information(MW, '欢迎使用', '''本程序是在2天速成PyQt下由 CLI版 ASFMKV py1.02 Preview9_2f2 改造而来
由于本身逻辑就不是为图形化服务的，有很多东西会比较尴尬
因而本程序并不隐藏CLI，CLI上有许多提示比图形界面的提示详细
目前GUI的设计依然存在一些问题，比如表格异常难看之类的，还请见谅

如果您需要使用系统字体和自定义字体目录，请点击\"全局刷新\"
各功能的详细信息写在按钮、选择框和输入框的Tip里了，请悬停鼠标查看''')

    sys.exit(app.exec_())