# -*- coding: UTF-8 -*-
# *************************************************************************
#
# 请使用支持 UTF-8 NoBOM 并最好带有 Python 语法高亮的文本编辑器
# Windows 7 的用户请最好不要使用 记事本 打开本脚本
#
# *************************************************************************

# 调用库，请不要修改
import shutil
import traceback
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
lcidfil = { 
3: {
    2052 : 'gbk',
    1042 : 'euc-kr',
    1041 : 'shift-jis',
    1033 : 'utf-16be',
    1028 : 'big5',
    0 : 'utf-16be'
}, 
2: {
    0 : 'ascii',
    1 : 'utf-8',
    2 : 'iso-8859-1'
},
1: {
    33 : 'gbk',
    23 : 'euc-kr',
    19 : 'big5',
    11 : 'shift-jis',
    0 : 'mac-roman'
}, 
0 : {
    0 : 'utf-16be',
    1 : 'utf-16be',
    2 : 'utf-16be',
    3 : 'utf-16be',
    4 : 'utf-16be',
    5 : 'utf-16be'
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
if o_fontload and not fontload: fontload = True
langlist = []
extsupp = []
dupfont = {}
init()

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

def updateAssFont(assfont: dict = {}, assfont2: dict = {}) -> dict:
    if assfont2 is None: return None
    for key in assfont2.keys():
        if assfont.get(key) is not None:
            assfont[key] = [
                assfont[key][0] + ''.join([ks for ks in assfont2[key][0] if ks not in assfont[key][0]]), 
                assfont[key][1] + ''.join(['|' + ks for ks in assfont2[key][1].split('|') if ks.lower() not in assfont[key][1].lower()]), 
                assfont[key][2] + ''.join(['|' + ks for ks in assfont2[key][2].split('|') if ks.lower() not in assfont[key][2].lower()])
                ]
        else:
            assfont[key] = assfont2[key]
    del assfont2
    return assfont

def assAnalyze(asspath: str, fontlist: dict = {}, onlycheck: bool = False):
    '''
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
    '''
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
    # 如果在 [Script Info] 中找到字体子集化的注释信息，则将字体名称还原并移除注释
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
        else: continue
    
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

    #print(fontlist)

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
        #print(fullass[i])
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
        if re.search(r'\{.*?\\p[1-9][0-9]*.*?\}([\s\S]*?)\{.*?\\p0.*?\}', eventftext) is not None:
            vecpos = re.findall(r'\{.*?\\p[1-9][0-9]*.*?\}[\s\S]*?\{.*?\\p0.*?\}', eventftext)
            vecfind = 0
            nexts = 0
            for s in vecpos:
                vecfind = eventftext.find(s)
                textremain += eventftext[nexts:vecfind]
                nexts = vecfind
                s = re.sub(r'\\p[0-9]+', '', re.sub(r'}.*?{', '}{', s))
                textremain += s
        elif re.search(r'\{.*?\\p[1-9][0-9]*.*?\}', eventftext) is not None:
            eventftext = re.sub(r'\\p[0-9]+', '', eventftext[:re.search(r'\{.*?\\p[1-9][0-9]*.*?\}', eventftext).span()[0]])
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
                        if re.search(r'(\\b[7-9]00|\\b1)', st).span()[0] > max([st.find('\\b0'), st.find('\\b\\'), st.find('\\b}')]):
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
                else: continue
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
            fn = ''
            fn_line = [i]
            newstrl = []
            # it = Index Font
            it = (0, 0)
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
                        l = [sf.strip(' ') for sf in s.split('\\') if len(s.strip(' ')) > 0]
                        l.reverse()
                        for sf in l:
                            if 'fn' in sf.lower():
                                fn = sf[2:].strip(' ').lstrip('@')
                                if len(ssRecover) > 0:
                                    if ssRecover.get(fn) is not None:
                                        fullassSplit = fullass[i].split(',')
                                        fullass[i] = ','.join(fullassSplit[0:text_pos] + [','.join(fullassSplit[text_pos:]).replace(fn, ssRecover[fn])])
                                        s = s.replace(fn, ssRecover[fn])
                                        fn = ssRecover[fn]
                                fn_line.append({s : (fn, str(sp[2]), str(sp[3]))})
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
            if not eventfont is None:
                # 去除行中非文本部分，包括特效标签{}，硬软换行符
                eventtext = re.sub(effectDel, '', eventftext)
                fontlist = fontlistAdd(eventtext, eventfont, fontlist)

    fl_popkey = []
    # 在字体列表中检查是否有没有在文本中使用的字体，如果有，添加到删去列表
    for s in fontlist.keys():
        if len(fontlist[s]) == 0:
            fl_popkey.append(s)
            #print('跳过没有字符的字体\"{0}\"'.format(s))
    # 删去 删去列表 中的字体
    if len(fl_popkey) > 0:
        for s in fl_popkey:
            fontlist.pop(s)
    # print(fontlist)
    # 如果含有字体恢复信息，则恢复字幕
    if len(ssRecover) > 0:
        try:
            nsdir = path.join(path.dirname(asspath), 'NoSubsetSub')
            #nsfile = path.join(nsdir, '.NoSubset'.join(path.splitext(path.basename(asspath))))
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
    '''
获取字体文件列表

接受输入
  customPath: 用户指定的字体文件夹
  font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）
  noreg: 只从用户提供的customPath获取输入

将会返回
  filelist: 字体文件清单 [[ 字体绝对路径, 读取位置（'': 注册表, '0': 自定义目录） ], ...]
  font_name: 用于更新font_name（启用从注册表读取名称的功能时有效）    
    '''
    filelist = []

    if not noreg:
        # 从注册表读取
        fontkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts')
        fontkey_num = winreg.QueryInfoKey(fontkey)[1]
        #fkey = ''
        try:
            # 从用户字体注册表读取
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
            # 从 系统字体注册表 读取
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

def fnReadCorrect(ttfont: ttLib.ttFont, index: int, fontpath: str) -> str:
    '修正字体内部名称解码错误'
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
            #os.system('pause')
        else: namestr = name.toBytes().decode('utf-16-be', errors='ignore')
        # 如果没有 \x00 字符，使用 utf-16-be 强行解码；如果有，尝试解码；如果解码失败，使用 utf-16-be 强行解码
    except:
        print('\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
    return namestr

def outputSameLength(s: str) -> str:
    length = 0
    output = ''
    for i in range(len(s) -1, -1, -1):
        si = s[i]
        if 65281 <= ord(si) <= 65374 or ord(si) == 12288 or ord(si) not in range(33, 127):
            length += 2
        else: length += 1
        if length >= 56:
            output = '..' + output
            break
        else: output = si + output
    return output + ''.join([' ' for ss in range(0, 60 - length)])

# font_info 列表结构
#   [ font_name, font_n_lower, font_family, warning_font ]
# font_name 词典结构
#   { 字体名称 : [ 字体绝对路径 , 字体索引 (仅用于TTC/OTC; 如果是TTF/OTF，默认为0) , 字体样式 ] }
# dupfont 词典结构
#   { 重复字体名称 : [ 字体1绝对路径, 字体2绝对路径, ... ] }
def fontProgress(fl: list, font_info: list = [{}, {}, {}, []], overwrite: bool = False) -> list:
    '''
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
    '''
    global dupfont
    warning_font = font_info[3]
    font_family = font_info[2]
    font_n_lower = font_info[1]
    f_n = font_info[0]
    #print(fl)
    flL = len(fl)
    print('\033[1;32m正在读取字体信息...\033[0m')
    for si in range(0, flL):
        s = fl[si][0]
        fromCustom = False
        fl[si][1] = str(fl[si][1])
        if len(fl[si][1]) > 0: fromCustom = True
        # 如果有来自自定义文件夹标记，则将 fromCustom 设为 True
        ext = path.splitext(s)[1][1:]
        # 检查字体扩展名
        if ext.lower() not in ['ttf','ttc','otf','otc']: continue
        # 输出进度
        print('\r' + '\033[1;32m{0}/{1} {2:.2f}%\033[0m \033[1m{3}\033[0m'.format(si + 1, flL, ((si + 1)/flL)*100, outputSameLength(path.dirname(s))), end='', flush=True)
        isTc = False
        if ext.lower() in ['ttf', 'otf']:
            # 如果是 TTF/OTF 单字体文件，则使用 TTFont 读取
            try:
                tc = [ttLib.TTFont(s, lazy=True)]
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTC/OTC模式读取\033[0m'.format(s, sys.exc_info()))
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
                tc =  ttLib.TTCollection(s, lazy=True)
                isTc = True
            except:
                print('\033[1;31m\n[ERROR] \"{0}\": {1}\n\033[1;34m[TRY] 正在尝试使用TTF/OTF模式读取\033[0m'.format(s, sys.exc_info()))
                try:
                    # 如果读取失败，可能是使用了错误扩展名的 TTF/OTF 文件，用 TTFont 尝试读取
                    tc = [ttLib.TTFont(s, lazy=True)]
                    print('\033[1;34m[WARNING] 错误的字体扩展名\"{0}\" \033[0m'.format(s))
                except:
                    print('\033[1;31m\n[ERROR] \"{0}\": {1}\033[0m'.format(s, sys.exc_info()))
                    continue
            #f_n[path.splitext(path.basename(s))[0]] = [s, 0]
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
                    namestr = ''
                    try:
                        # if name.isUnicode() or name.platformID == 0:
                        namestr = name.toStr()
                        # else:
                        #     errnum = 0
                        #     lcidcache = [lcc for lcc in lcidfil[name.platformID].values() if re.search(r'(unicode)|(utf)|(roman)', lcc.lower()) is None]
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
            #print(namestrl)
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
            #print(namestr, path.basename(s))
            for namestr in namestrl:
                #print(namestr)
                if isWarningFont:
                    if namestr not in warning_font:
                        warning_font.append(namestr.lower())

                if f_n.get(namestr) is not None and not overwrite:
                    # 如果发现列表中已有相同名称的字体，检测它的文件名、扩展名、父目录、样式是否相同
                    # 如果有一者不同且不来自自定义文件夹，添加到重复字体列表
                    dupp = f_n[namestr][0]
                    if (dupp != s and path.splitext(path.basename(dupp))[0] != path.splitext(path.basename(s))[0] and 
                    not fromCustom and f_n[namestr][2].lower() == fstyle.lower()):
                        print('\n\033[1;35m[WARNING] 字体\"{0}\"与字体\"{1}\"的名称\"{2}\"重复！\033[0m'.format(path.basename(f_n[namestr][0]), path.basename(s), namestr))
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
                        if not namestr in font_family[familyN][(isItalic, isBold)]:
                            font_family[familyN][(isItalic, isBold)].append(namestr)
                    else: font_family[familyN].setdefault((isItalic, isBold), [namestr])
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
    return [f_n, font_n_lower, font_family, warning_font]

#print(filelist)
#if path.exists(fontspath10): filelist = filelist.extend(os.listdir(fontspath10))

#for s in font_name.keys(): print('{0}: {1}'.format(s, font_name[s]))
def fnGetFromFamilyName(font_family: str, fn: str, isitalic: int, isbold: int) -> str:
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
    '''
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
    '''
    # 从fontlist获取字体名称
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
        else: cok = True
        directout = 0
        if not cok:
            # 如果 onlycheck 不为 True，向用户要求目标字体
            if not onlycheck:
                print('\033[1;31m[ERROR] 缺少字体\"{0}\"\n请输入追加的字体文件或其所在字体目录的绝对路径\033[0m'.format(ns))
                addFont = {}
                inFont = ''
                while inFont == '' and directout < 3:
                    inFont = input().strip('\"').strip(' ')
                    if path.exists(inFont):
                        if path.isdir(inFont):
                            addFont = fontProgress(getFileList([inFont], noreg=True)[0])
                            print()
                        else:
                            addFont = fontProgress([[inFont, '0']])
                            print()
                        if ns.lower() not in addFont[1].keys():
                            if path.isdir(inFont):
                                print('\033[1;31m[ERROR] 输入路径中\"{0}\"没有所需字体\"{1}\"\033[0m'.format(inFont, ns))
                            else: print('\033[1;31m[ERROR] 输入字体\"{0}\"不是所需字体\"{1}\"\033[0m'.format('|'.join(addFont[0].keys()), ns))
                            inFont = ''
                        else:
                            if f_priority: 
                                font_name.update(addFont[0])
                                font_n_lower.update(addFont[1])
                                font_family.update(addFont[2])
                                warning_font.extend(addFont[3])
                            else:
                                for fd in addFont[0].keys():
                                    if fd not in font_name:
                                        font_name[fd] = addFont[0][fd]
                                for fd in addFont[1].keys():
                                    if fd not in font_n_lower:
                                        font_n_lower[fd] = addFont[1][fd]
                                for fd in addFont[2].keys():
                                    if fd not in font_family:
                                        font_family[fd] = addFont[2][fd]
                                for fd in addFont[3]:
                                    if fd not in warning_font:
                                        warning_font.append(fd)
                            cok = True
                    else:
                        print('\033[1;31m[ERROR] 您没有输入任何字符！再回车{0}次回到主菜单\033[0m'.format(3-directout))
                        directout += 1
                        inFont = ''
            else:
                # 否则直接添加空
                assfont['?'.join([ns, ns])] = ['', sp[0], ns]
        if cok:
            if not onlycheck and ss.lower() in warning_font: 
                print('\033[1;31m[WARNING] 字体\"{0}\"可能不能正常子集化\033[0m'.format(ss))
                if warningStop:
                    print('\033[1;31m[WARNING] 请修复\"{0}\"，工作已中断\033[0m'.format(ss))
                    return None, [font_name, font_n_lower, font_family, warning_font]
            if directout < 3:
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
                #print(assfont[dict_key])
        if directout >= 3:
            return None, [font_name, font_n_lower, font_family, warning_font]
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
    '''
字体内部名称变更

变更NameID为1, 3, 4, 6的NameRecord，它们分别对应

    ID  Meaning

    1   Font Family name

    3   Unique font identifier

    4   Full font name

    6   PostScript name for the font

注意本脚本并不更改 NameID 为 0 和 7 的版权信息
'''
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

def assFontSubset(assfont: dict, fontdir: str) -> dict:
    '''
字体子集化

需要以下输入:
  assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称 ]}
  fontdir: 新字体存放目录

将会返回以下:
  newfont_name: { 原字体名 : [ 新字体绝对路径, 新字体名 ] }
    '''
    global errorStop
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
    else: fontdir = os.getcwd()
    print('\033[1;33m字体输出路径:\033[0m \033[1m\"{0}\"\033[0m'.format(fontdir))

    lk = len(assfont.keys())
    kip = 0
    for k in assfont.keys():
        kip += 1
        # 偷懒没有变更该函数中的assfont解析到新的词典格式
        # 在这里会将assfont词典转换为旧的assfont列表形式
        # assfont: [ 字体绝对路径, 字体索引, 字符串, ASS内的字体名称, 修正字体名称 ]
        s = k.split('?') + [assfont[k][0], assfont[k][1], assfont[k][2]]
        subfontext = ''
        fontext = path.splitext(path.basename(s[0]))[1]
        if fontext[1:].lower() in ['otc', 'ttc']:
            subfontext = fontext[:3].lower() + 'f'
        else: subfontext = fontext
        #print(fontdir, path.exists(path.dirname(fontdir)), path.exists(fontdir))
        fontname = re.sub(cillegal, '_', s[4])
        subfontpath = path.join(fontdir, fontname + subfontext)
        subsetarg = [s[0], '--text={0}'.format(s[2]), '--output-file={0}'.format(subfontpath), '--font-number={0}'.format(s[1]), '--passthrough-tables']
        print('\r\033[1;32m[{0}/{1}]\033[0m \033[1m正在子集化…… \033[0m'.format(kip, lk), end='')
        try:
            subset.main(subsetarg)
        except PermissionError:
            print('\n\033[1;31m[ERROR] 文件\"{0}\"访问失败\033[0m'.format(path.basename(subfontpath)))
            continue
        except:
            # print('\033[1;31m[ERROR] 失败字符串: \"{0}\" \033[0m'.format(s[2]))
            print('\n\033[1;31m[ERROR] {0}\033[0m'.format(sys.exc_info()))
            if errorStop: 
                print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，强制终止批量处理\033[0m'.format(path.basename(s[0])))
                return None
            print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
            # crcnewf = ''.join([path.splitext(subfontpath)[0], fontext])
            # shutil.copy(s[0], crcnewf)
            ttLib.TTFont(s[0], lazy=False, fontNumber=int(s[1])).save(subfontpath, False)
            subfontcrc = None
            # newfont_name[s[3]] = [crcnewf, subfontcrc]
            newfont_name[s[3]] = [subfontpath, subfontcrc]
            continue
        #os.system('pyftsubset {0}'.format(' '.join(subsetarg)))
        if path.exists(subfontpath): 
            subfontbyte = open(subfontpath, mode='rb')
            subfontcrc = str(hex(zlib.crc32(subfontbyte.read())))[2:].upper()
            if len(subfontcrc) < 8: subfontcrc = '0' + subfontcrc
            # print('CRC32: {0} \"{1}\"'.format(subfontcrc, path.basename(s[0])))
            subfontbyte.close()
            rawf = ttLib.TTFont(s[0], lazy=True, fontNumber=int(s[1]))
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
                    namestr = getNameStr(name, subfontcrc)
                    newf['name'].setName(namestr ,name.nameID, name.platformID, name.platEncID, name.langID)
            if len(newf.getGlyphOrder()) == 1 and '.notdef' in newf.getGlyphOrder():
                if errorStop: 
                    print('\033[1;31m[WARNING] 字体\"{0}\"子集化失败，强制终止批量处理\033[0m'.format(path.basename(s[0])))
                    rawf.close()
                    newf.close()
                    if not subfontpath == s[0]: os.remove(subfontpath)
                    return None
                print('\n\033[1;31m[WARNING] 字体\"{0}\"子集化失败，将会保留完整字体\033[0m'.format(path.basename(s[0])))
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
                newfont_name[s[3]] = [crcnewf, subfontcrc]
    print('')
    #print(newfont_name)
    return newfont_name

def assFontChange(fullass: list, newfont_name: dict, asspath: str, styleline: int, infoline: int,
                        font_pos: int, outdir: str = '', ncover: bool = False, fn_lines: list = []) -> str:
    '''
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
    '''
    # 扫描Style各行，并替换掉字体名称
    #print('正在替换style对应字体......')
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
        #print('正在处理fn标签......')
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
    #print(fontpaths)
    global rmAssIn, rmAttach, mkvout, notfont
    if file is None: return 4
    elif file == '': return 4
    elif not path.exists(file) or not path.isfile(file): return 4
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
            #print(assfn, fn, assnote)
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
    mkvmr = os.system('mkvmerge @\"{0}\" -o \"{1}\"'.format(mkvjsonp, outfile))
    if mkvmr > 1:
        print('\n\033[1;31m[ERROR] 检测到不正常的mkvmerge返回值，重定向输出...\033[0m')
        os.system('mkvmerge -r \"{0}\" @\"{1}\" -o NUL'.format('{0}.{1}.log'
            .format(path.splitext(file)[0], datetime.now().strftime('%Y-%m%d-%H%M-%S_%f')), mkvjsonp))
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
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    else:
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    return mkvmr


def getMediaFilelist(dir: str) -> list:
    '''
从输入的目录中获取媒体文件列表
需要以下输入
  dir: 要搜索的目录
返回以下结果
  medias: 多媒体文件列表
      结构: [[ 文件名（无扩展名）, 绝对路径 ], ...]
    '''
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


def getSubtitles(cpath: str, medias: list = []) -> dict:
    '''
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
    '''
    media_ass = {}
    if len(medias) == 0: media_ass['none'] = []
    global s_subdir, matchStrict
    if s_subdir:
        for r,ds,fs in os.walk(cpath):
            for f in [path.join(r, s) for s in fs if path.splitext(s)[1][1:].lower() in subext]:
                if '.subset' in path.basename(f): continue
                if len(medias) > 0:
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
                    media_ass['none'].append(path.join(r, ds, fs))
    else:
        for f in [path.join(cpath, s) for s in os.listdir(cpath) if not path.isdir(s) and 
        path.splitext(s)[1][1:].lower() in subext]:
            # print(f, cpath)
            if '.subset' in path.basename(f): continue
            if len(medias) > 0:
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
            else:
                media_ass['none'].append(f)
    return media_ass

def nameMatchingProgress(medias: list, mStart: int, mEnd: str, sStart: int, sEnd: str, subs: list) -> dict:
    '''
视频-字幕匹配的流程部分，负责匹配视频和字幕

需要以下输入
  medias: 媒体文件列表，格式见 nameMatching
  mStart: 媒体文件剧集开始位置
  mEnd: 媒体文件剧集结束字符
  sStart: 字幕文件剧集开始位置
  sEnd: 字幕文件剧集结束字符
  subs: 字幕文件列表
    '''
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
    '''
字幕匹配的控制部分

需要以下输入
  spath: 字幕路径
  medias: 媒体文件列表，遵循格式为 [[文件名(无扩展名), 完整路径], ...]
    '''
    # 调用各函数获得剧集位置信息及字幕文件列表
    mStart, mEnd = namePosition([s[0] for s in medias])
    subs = getSubtitles(spath)['none']
    sStart, sEnd = namePosition([path.splitext(path.basename(s))[0] for s in subs])
    sStartA, sEndA = namePosition([path.splitext(path.basename(s))[0][::-1] for s in subs])
    # 生成第一次的 media_ass 和 renList，并询问
    media_ass = nameMatchingProgress(medias, mStart, mEnd, sStart, sEnd, subs)
    renList = renamePreview(media_ass, sStartA, sEndA)
    annl = []
    ann = ''
    if len(renList) > 0:
        print('''请确认重命名是否正确
如果注释部分不正确您可以选择\033[33;1m[A]\033[0m手动输入字幕注释
如果文件匹配不正确您可以选择\033[33;1m[F]\033[0m手动输入匹配规则''')
        rt = os.system('choice /M 请选择 /C YNAF')
    else:
        rt = 2
    while rt in [0,3,4]:
        if rt == 4:
            cls()
            print('''请将剧集所在位置用\"\033[33;1m:\033[0m\"代替
如 \"\033[30;1m[VCB-Studio] Accel World [\033[33;1m01\033[30;1m][720p][x265_aac].mp4\033[0m\"
替换为 \"\033[30;1m[VCB-Studio] Accel World [\033[33;1m:\033[30;1m][720p][x265_aac].mp4\033[0m\"
请务必保留剧集位置后的一个字符以以便程序正常工作''')
            vSearch = input('请输入视频匹配规则:').strip(' ').strip('\"').strip('\'')
            if len(vSearch) == 0:
                print('输入为空，退出')
                break
            sSearch = input('请输入字幕匹配规则:').strip(' ').strip('\"').strip('\'')
            if len(sSearch) == 0:
                print('输入为空，退出')
                break
            vSplit = vSearch.split(':')
            if len(vSplit) < 2:
                print('输入不正确，退出')
                break
            sSplit = sSearch.split(':')
            if len(sSplit) < 2:
                print('输入不正确，退出')
                break
            media_ass = nameMatchingProgress(medias, len(vSplit[0]), vSplit[1][0], len(sSplit[0]), sSplit[1][0], subs)
            del vSearch, sSearch, vSplit, sSplit
        elif rt == 3:
            ann = input('请输入（可以为空，可以有多个，按预览顺序用\"\033[33;1m:\033[0m\"分隔）:')
            if ':' in ann:
                annl = [re.sub(cillegal, '_', a) for a in ann.split(':') if len(a) > 0]
                ann = ''
            elif len(ann) > 0:
                if ann[0] != '.': ann = '.' + ann
        else:
            break
        renList = renamePreview(media_ass, ann=re.sub(cillegal, '_', ann), annl=annl)
        if len(renList) > 0:
            print('''请确认重命名是否正确
如果注释部分不正确您可以选择\033[33;1m[A]\033[0m手动输入字幕注释
如果文件匹配不正确您可以选择\033[33;1m[F]\033[0m手动输入匹配规则''')
            rt = os.system('choice /M 请选择 /C YNAF')
        else:
            rt = 2
    else:
        if rt == 1:
            logfile = open(path.join(spath, 'subsRename_{0}.log'.format(datetime.now().strftime('%Y%m%d-%H%M%S'))), mode='w', encoding='utf-8')
            for r in renList:
                try:
                    os.rename(r[0], path.join(path.dirname(r[0]), r[1]))
                except:
                    print('\033[31;1m[ERROR] \"{0}\"重命名失败\033[0m'.format(r[0]))
                else:
                    print('\033[32;1m[SUCCESS]\033[33;0m \"{0}\"\n          \033[0m>>> \033[1m\"{1}\"\033[0m'.format(path.basename(r[0]), r[1]))
                    try:
                        logfile.write('{0}|{1}\n'.format(r[0], r[1]))
                    except:
                        print('\033[31;1m[ERROR] 日志文件写入失败\033[0m')
            logfile.close()
            del logfile
    del rt, mStart, mEnd, sStart, sEnd, sStartA, sEndA, media_ass, renList, annl, ann


def renamePreview(media_ass: dict, sStartA: int = 0, sEndA: str = '', ann: str = '', annl: list = []) -> list:
    '''
提供重命名预览及制作重命名对应表

需要以下输入
  media_ass: 媒体文件与字幕文件的对应词典，遵循 {(媒体文件名(无扩展名), 完整路径): [字幕完整路径1, 2, ..], ...}

以下输入可选
  sStartA: 字幕注释开始位置
  sEndA: 字幕注释结束字符
  ann: 单个字幕注释
  annl: 多个字幕注释列表

将会返回以下值
  renList: 字幕重命名前后对应表，格式为 [(重命名前完整路径, 重命名后文件名), ...]
    '''
    global matchStrict
    cls()
    print('[重命名预览]\n')
    renList = []
    for k in media_ass.keys():
        sl = media_ass[k]
        annlUp = 0
        for s in sl:
            bname = path.basename(s)
            if matchStrict:
                if len(bname) >= len(k[0]):
                    if k[0] == bname[:len(k[0])]:
                        print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                        continue
            else:
                if k[0] in bname:
                    print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                    continue
            if len(annl) == 0:
                rsName = path.splitext(bname)[0][::-1]
                if sStartA - 2 > 1:
                    if rsName[sStartA - 2] == '.' and sEndA != '.': 
                        ann = rsName[:sStartA - 1][::-1]
                elif rsName.find('.') > 0:
                    ann = '.' + rsName[:rsName.find('.')][::-1]
                    print('使用了暴力注释抽取，请务必检查注释是否有误')
                elif len(sl) > 1:
                    ann = '.sub{0}'.format(sl.index(s))
                    print('注释自动检测失败，自动添加防重名注释')
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
            renList.append((s, '{0}{1}{2}'.format(k[0], ann, path.splitext(s)[1])))
    if len(renList) > 0:
        print('字幕的注释是指类似以下文件名的高亮部分\n\"\033[30;1m[VCB-Studio] Accel World [01][720p][x265_aac]\033[33;1m.Kamigami-sc\033[30;1m.ass\033[0m\"')
    return renList


def namePosition(files: list):
    '''
自动匹配文件名中的剧集位置

需要以下输入
  files: 文件列表 [文件名, ...]
    '''
    diffList = {}
    diffList['mStart'] = []
    diffList['mEnd'] = []
    sOffset = 1
    step = 1
    for m1 in files:
        diffList[m1] = []
        for m2 in files:
            if m2 == m1:
                continue
            sStop = False
            dStart = 0
            dEnd = ''
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
        for l in diffList[m1]:
            if l[0] not in apStart: apStart.append(l[0])
            alStart.append(l[0])
            if l[1] not in apEnd: apEnd.append(l[1])
            alEnd.append(l[1])
        mStart = 0
        mStartC = 0
        mEnd = ''
        mEndC = 0
        for l in apStart:
            c = alStart.count(l)
            if c > mStartC:
                mStart, mStartC = l, c
        for l in apEnd:
            c = alEnd.count(l)
            if c > mEndC:
                mEnd, mEndC = l, c
        del diffList[m1]
        diffList['mStart'].append(mStart)
        diffList['mEnd'].append(mEnd)
        del apStart, alStart, apEnd, alEnd
    del m1, m2
    mStart, mStartC, mEnd, mEndC = 0, 0, '', 0
    for l in set(diffList['mStart']):
        c = diffList['mStart'].count(l)
        if c > mStartC:
            mStart, mStartC = l, c
    for l in set(diffList['mEnd']):
        c = diffList['mEnd'].count(l)
        if c > mEndC:
            mEnd, mEndC = l, c
    mStartN = mStart
    for l in set(diffList['mStart']):
        if l in range(mStart - sOffset, mStart):
            mStartN = l
    if mStartN < mStart: mStart = mStartN
    del mStartN, mStartC, mEndC, diffList
    # print(diffList['mStart'])
    # print(diffList['mEnd'])
    return mStart, mEnd


def main(font_info: list, asspath: list, outdir: list = ['', '', ''], mux: bool = False, vpath: str = '', asslangs: list = []):
    '''
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
    '''
    print('')
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
    fullass = []
    newfont_name = {}
    fo = ''
    if not notfont:
        # print('\n字体名称总数: {0}'.format(len(font_name.keys())))
        # noass = False
        for i in range(0, len(asspath)):
            s = asspath[i]
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                multiAss[s] = [[], 0, 0, 0]
            else:
                # print('正在分析字幕文件: \"{0}\"'.format(path.basename(s)))
                fullass, fontlist, styleline, font_pos, fn_lines, infoline = assAnalyze(s, fontlist)
                multiAss[s] = [fullass, styleline, font_pos, infoline]
                assfont2, font_info = checkAssFont(fontlist, font_info)
                assfont = updateAssFont(assfont, assfont2)
                del assfont2
                if assfont is None:
                    return None, None, -2
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
        print('\033[1m字幕所需字体\033[0m')
        for af in assfont.values():
            print('\033[1m\"{0}\"\033[0m: 字符数[\033[1;33m{1}\033[0m]'.format(af[2], len(af[0])))
        newfont_name = assFontSubset(assfont, fo)
        if newfont_name is None: return [], {}, -2
        for s in asspath:
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                newasspath.append(s)
            elif len(multiAss[s][0]) == 0 or multiAss[s][1] == multiAss[s][2]: 
                continue
            else: newasspath.append(assFontChange(multiAss[s][0], newfont_name, s, multiAss[s][1], multiAss[s][3], multiAss[s][2], outdir[0], fn_lines=fn_lines))
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
    '临时字体载入'
    global s_fontload
    font_list = []
    print('\033[1;32m正在尝试从字幕所在文件夹载入字体...\033[0m')
    if s_fontload:
        for r,ds,fs in os.walk(_dir):
            fpath = path.join(r, ds, fs)
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
    if len(font_list) > 0 :
        font_info = fontProgress(font_list, font_info, True)
        print()
    else:
        print('\033[1;32m目录下没有字体\033[0m')
    return font_info

def cls():
    os.system('cls')

def cListAssFont(font_info):
    global resultw, s_subdir, copyfont, fontload
    font_name = font_info[0]
    leave = True
    while leave:
        fontlist = {}
        assfont = {}
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
[4] 使用工作目录字体：\033[1;33m{3}\033[0m
'''.format(resultw, copyfont, s_subdir, fontload))
        work = os.system('choice /M 请输入 /C AB1234L')
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
            if not directout:
                if path.isdir(cpath):
                    #print(cpath)
                    if fontload: font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                    else: font_info2 = font_info
                    if s_subdir:
                        for r,ds,fs in os.walk(cpath):
                            for f in fs:
                                if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                                    a, fontlist, b, c, d, e = assAnalyze(path.join(r, f), fontlist, onlycheck=True)
                                    del a,b,c,d,e
                                    assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True)
                                    assfont = updateAssFont(assfont, assfont2)
                                    del assfont2
                    else:
                        for f in os.listdir(cpath):
                            if path.isfile(path.join(cpath, f)):
                                #print(f, path.splitext(f)[1][1:].lower())
                                if path.splitext(f)[1][1:].lower() in ['ass', 'ssa']:
                                    #print(f, 'pass')
                                    a, fontlist, b, c, d, e = assAnalyze(path.join(cpath, f), fontlist, onlycheck=True)
                                    del a,b,c,d,e
                                    assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True)
                                    assfont = updateAssFont(assfont, assfont2)
                                    del assfont2
                    fd = path.join(cpath, 'Fonts')
                else:
                    if fontload: font_info2 = templeFontLoad(path.dirname(cpath), copy.deepcopy(font_info))
                    else: font_info2 = font_info
                    a, fontlist, b, c, d, e = assAnalyze(cpath, fontlist, onlycheck=True)
                    del a,b,c,d,e
                    assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True)
                    assfont = updateAssFont(assfont, assfont2)
                    del assfont2
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
                    maxnum = 0
                    for s in assfont.keys():
                        ssp = s.split('?')
                        if not ssp[0] == ssp[1]:
                            lx = len(assfont[s][0])
                            if lx > maxnum:
                                maxnum = lx
                    maxnum = len(str(maxnum))
                    for s in assfont.keys():
                        ssp = s.split('?')
                        if not ssp[0] == ssp[1]:
                            fp = ssp[0]
                            fn = path.basename(fp)
                            ann = ''
                            errshow = False
                            if copyfont: 
                                try:
                                    if not path.exists(path.join(fd, fn)):
                                        shutil.copy(fp, path.join(fd, fn))
                                        ann = ' - copied'
                                    else:
                                        ann = ' - existed'
                                except:
                                    print('[ERROR]', sys.exc_info())
                                    ann = ' - copy error'
                                    errshow = True
                            if resultw:
                                print('{0} <{1}>{2}'.format(assfont[s][2], path.basename(fn), ann), file=wfile)
                            if errshow:
                                print('\033[1;32m[{3}]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;31m{2}\033[0m'.format(assfont[s][2], 
                                path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum)))
                            else: print('\033[1;32m[{3}]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;32m{2}\033[0m'.format(assfont[s][2], 
                            path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum)))
                            # print(assfont[s][0])
                        else:
                            if resultw:
                                print('{0} - No Found'.format(ssp[0]), file=wfile)
                            print('\033[1;31m[{1}]\033[1;36m {0}\033[1;31m - No Found\033[0m'.format(ssp[0], 'N'.center(maxnum, 'N')))
                    if resultw: wfile.close()
                    print('')
                del font_info2
        elif work == 3:
            if resultw: resultw = False 
            else: resultw = True
        elif work == 4:
            if copyfont: copyfont = False
            else: copyfont = True
        elif work == 5:
            if s_subdir: s_subdir = False
            else: s_subdir = True
        elif work == 6:
            if o_fontload == False:
                if fontload: fontload = False
                else: fontload = True
            else:
                cls()
                print('在禁用系统字体源的情况下，fontload必须为True')
                os.system('pause')
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

def showMessageSubset(newasspaths: list, newfont_name: dict):
    for ap in newasspaths:
        if path.exists(ap): 
            print('\033[1;32m成功:\033[0m \033[1m\"{0}\"\033[0m'.format(path.basename(ap)))
    for nf in newfont_name.keys():
        if path.exists(newfont_name[nf][0]):
            if newfont_name[nf][1] is None:
                print('\033[1;31m失败:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\"\033[0m'.format(path.basename(nf), 
                path.basename(newfont_name[nf][0])))
            else:
                print('\033[1;32m成功:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\" ({2})\033[0m'.format(path.basename(nf), 
                path.basename(newfont_name[nf][0]), newfont_name[nf][1]))

def cFontSubset(font_info):
    global extlist, v_subdir, s_subdir, rmAssIn, rmAttach, fontload, \
    mkvout, assout, fontout, matchStrict, no_mkvm, notfont, warningStop, errorStop
    leave = True
    while leave:
        cls()
        print('''ASFMKV & ASFMKV-FontSubset
选择功能:
[A] 子集化字体
[B] 子集化并封装
[L] 回到上级菜单
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
[X] 旧格式字体中断: \033[1;33m{9}\033[0m
[Y] 子集化失败中断: \033[1;33m{10}\033[0m
[Z] 使用工作目录字体: \033[1;33m{11}\033[0m
'''.format(v_subdir, s_subdir, rmAssIn, rmAttach, mkvout, assout, 
fontout, matchStrict, notfont, warningStop, errorStop, fontload))
        work = 0
        work = os.system('choice /M 请输入 /C AB1234567890XYZL')
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
旧格式字体中断: \033[1;33m{5}\033[0m
子集化失败中断: \033[1;33m{6}\033[0m
使用工作目录字体：\033[1;33m{7}\033[0m
'''.format(v_subdir, s_subdir, matchStrict, assout, fontout, warningStop, warningStop, fontload))
            else: print('''子集化字体并封装
搜索子目录(视频): \033[1;33m{4}\033[0m
搜索子目录(字幕): \033[1;33m{5}\033[0m
移除内挂字幕: \033[1;33m{0}\033[0m
移除原有附件: \033[1;33m{1}\033[0m
不封装字体: \033[1;33m{2}\033[0m
严格字幕匹配: \033[1;33m{6}\033[0m
媒体文件输出文件夹: \033[1;33m{3}\033[0m
旧格式字体中断: \033[1;33m{8}\033[0m
子集化失败中断: \033[1;33m{7}\033[0m
使用工作目录字体：\033[1;33m{9}\033[0m
'''.format(rmAssIn, rmAttach, notfont, mkvout, v_subdir, s_subdir, matchStrict, errorStop, warningStop, fontload))
            cpath = ''
            directout = False
            subonly = False
            subonlyp = []
            while not path.exists(cpath) and not directout:
                directout = True
                cpath = input('不输入任何值 直接回车回到上一页面\n请输入文件或目录路径: ').strip('\"')
                if cpath == '':  print('没有输入，回到上级菜单')
                elif not path.isabs(cpath): print('\033[1;31m[ERROR] 输入的必须是绝对路径！: \"{0}\"\033[0m'.format(cpath))
                elif not path.exists(cpath): print('\033[1;31m[ERROR] 找不到路径: \"{0}\"\033[0m'.format(cpath))
                elif path.isfile(cpath): 
                    testext = path.splitext(cpath)[1][1:].lower()
                    if testext in extlist:
                        directout = False
                    elif testext in ['ass', 'ssa'] and work == 1:
                        directout = False
                        subonly = True
                    else: print('\033[1;31m[ERROR] 扩展名不正确: \"{0}\"\033[0m'.format(cpath))
                elif not path.isdir(cpath):
                    print('\033[1;31m[ERROR] 输入的应该是目录或媒体文件！: \"{0}\"\033[0m'.format(cpath))
                else: directout = False
            # print(directout)
            medias = []
            if not directout:
                if path.isfile(cpath): 
                    if not subonly: medias = [[path.splitext(path.basename(cpath))[0], cpath]]
                    else: subonlyp = [cpath]
                    cpath = path.dirname(cpath)
                else: 
                    medias = getMediaFilelist(cpath)
                    if len(medias) == 0:
                        if s_subdir:
                            for r, ds, fs in os.walk(cpath):
                                if path.splitext(fs)[1][1:].lower() in ['ass', 'ssa']:
                                    subonlyp.append(path.join(r, ds, fs))
                        else:
                            for fs in os.listdir(cpath):
                                if path.splitext(fs)[1][1:].lower() in ['ass', 'ssa']:
                                    subonlyp.append(path.join(cpath, fs))
                        if len(subonlyp) > 0:
                            cls()
                            print('\033[1;33m[WARNING]\033[0m')
                            print('您输入的目录下只有字幕而无视频，每个字幕都将被当做单独对象处理\n若是同一话有多个字幕，这话将会有多套子集化字体。')
                            if os.system('choice /M \"即便如此，您仍要继续吗？\"') == 1:
                                subonly = True
                            else:
                                print('\n已终止运行')
                                directout = True
                        else: 
                            print('\033[1;31m[ERROR] 路径下找不到字幕\033[0m')
                            directout = True

                # print(medias)
            if not directout:
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

                domux = False
                if work == 2: domux = True

                if len(medias) > 0:
                    media_ass = getSubtitles(cpath, medias)
                    # print(media_ass)
                    if len(media_ass) > 0:
                        for k in media_ass.keys():
                            #print(k)
                            #print([assout_cache, fontout_cache, mkvout_cache])
                            if fontload: font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                            else: font_info2 = font_info
                            newasspaths, newfont_name, mkvr = main(font_info2, media_ass[k], 
                            mux = domux, outdir=[assout_cache, fontout_cache, mkvout_cache], vpath=k, asslangs=sublang)
                            if mkvr != -2:
                                showMessageSubset(newasspaths, newfont_name)
                            else:
                                break
                    else: print('\033[1;31m[ERROR] 找不到视频对应字幕\033[0m')
                elif subonly:
                    if fontload: font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                    else: font_info2 = font_info
                    for subp in subonlyp:
                        newasspaths, newfont_name, mkvr = main(font_info2, [subp], 
                            mux = False, outdir=[assout_cache, fontout_cache, mkvout_cache])
                    if mkvr != -2:
                        showMessageSubset(newasspaths, newfont_name)
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
        elif work == 13:
            if warningStop: warningStop = False
            else: warningStop = True
        elif work == 14:
            if errorStop: errorStop = False
            else: errorStop = True
        elif work == 15:
            if o_fontload == False:
                if fontload: fontload = False
                else: fontload = True
            else:
                cls()
                print('在禁用系统字体源的情况下，fontload必须为True')
                os.system('pause')
        else:
            leave = False
        if work < 4:
            os.system('pause')

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

def cFontSearch(font_info: list):
    leave = True
    while leave == True:
        cls()
        print('''ASFMKV FontSearch
请选择功能:
[1] 完全匹配（不区分大小写）
[2] 部分匹配（不区分大小写）
[0] 回到主菜单
''')
        work = os.system('choice /M 请选择: /C 012')
        if work == 1:
            leave = False
        elif work in [2,3]:
            cls()
            if work == 3: print('AND = \" \"（半角空格）\n仅支持AND')
            f_n = input('请输入字体名称:\n').strip(' ').strip('\"').lower()
            print('')
            if len(f_n) > 0:
                f_nget = None
                f_nt = None
                hadget = False
                if work == 2:
                    f_nt = font_info[1].get(f_n)
                    if f_nt is not None:
                        f_nget = font_info[0].get(f_nt)
                        if f_nget is not None:
                            print('\033[1;31m[{0}]\033[0m {2}\033[1m\n\"{1}\"\033[0m\n'.format(f_nt, f_nget[0], f_nget[2]))
                            hadget = True
                else:
                    for k in font_info[0].keys():
                        f_nget = font_info[0][k]
                        ksame = 0
                        for kand in f_n.split(' '):
                            if kand in k.lower():
                                ksame += 1
                        if ksame == len(f_n.split(' ')):
                            print('\033[1;31m[{0}]\033[0m {2}\033[1m\n\"{1}\"\033[0m\n'.format(k, f_nget[0], f_nget[2]))
                            hadget = True
                if not hadget:
                    print('\033[1;31m没有找到所搜索的字体\033[0m')
                os.system('pause')

no_mkvm = False
no_cmdc = False
mkvmv = ''
font_info = [{},{},{},[]]

# font_info 列表结构
#   [ font_name(dict), font_n_lower(dict), font_family(dict), warning_font(list) ]
def loadMain(reload: bool = False):
    global extlist, sublang, no_mkvm, no_cmdc, dupfont, mkvmv, font_info
    # 初始化字体列表 和 mkvmerge 相关参数
    os.system('title ASFMKV Python Remake 1.02-Pre9 ^| (c) 2022 yyfll ^| Apache-2.0')
    if not reload:
        if not o_fontload:
            font_list, font_name = getFileList(fontin)
            font_info = fontProgress(font_list, [font_name, {}, {}, []], f_priority)
            del font_name, font_list
        mkvmv = '\n\033[1;33m没有检测到 mkvmerge\033[0m'
        if os.system('mkvmerge -V 1>nul 2>nul') > 0:
            no_mkvm = True
        else: 
            print('\n\n\033[1;33m正在获取 mkvmerge 语言编码列表和支持格式列表，请稍等...\033[0m')
            mkvmv = '\n' + os.popen('mkvmerge -V --ui-language en', mode='r').read().replace('\n', '')
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

        if os.system('choice /? 1>nul 2>nul') > 0:
            no_cmdc = True

    cls()
    if o_fontload: fltip = '\033[1;35m已禁用系统字体读取\033[0m'
    else: fltip = '包含乱码的名称'
    print('''ASFMKV Python Remake 1.02-Pre9_2 | (c) 2022 yyfll{0}
字体名称数: [\033[1;33m{2}\033[0m]（{4}）
请选择功能:
[A] ListAssFont
[B] 字体子集化 & MKV封装
[M] 字幕-视频名称匹配
[S] 搜索字体
[C] 检视重复字体: 重复名称[\033[1;33m{1}\033[0m]
[W] 检视旧格式字体（可能无法子集化）: [\033[1;33m{3}\033[0m]
其他:
[R] 重载字体列表
[D] 依赖与许可证
[L] 直接退出'''.format(mkvmv, len(dupfont.keys()), len(font_info[0].keys()), len(font_info[3]), fltip))
    print('')
    work = os.system('choice /M 请选择: /C ABCSWDLRM')
    if work == 1:
        cListAssFont(font_info)
    elif work == 2:
        cFontSubset(font_info)
    elif work == 3:
        cls()
        if len(dupfont.keys()) > 0:
            for s in dupfont.keys():
                for ss in dupfont[s]:
                    print('\033[1;31m[{0}]\033[0m'.format(ss))
                print('\033[1m{0}\033[0m'.format('\n'.join(s)))
                print('')
        else:
            print('没有名称重复的字体')
        os.system('pause')
    elif work == 4:
        if not o_fontload: cFontSearch(font_info)
        else: 
            cls()
            print('您已禁用系统字体读取，该功能不可用')
            os.system('pause')
    elif work == 5:
        cls()
        if len(font_info[3]) > 0:
            oldflist = {}
            for s in font_info[3]:
                oldfname = font_info[1].get(s)
                oldfpath = font_info[0].get(oldfname)[0]
                if oldflist.get(oldfpath) is None:
                    oldflist[oldfpath] = [oldfname]
                else:
                    oldflist[oldfpath].append(oldfname)
            for k in oldflist:
                for kp in oldflist[k]:
                    print('\033[1;31m[{0}]\033[0m'.format(kp))
                print('\033[1m{0}\033[0m\n'.format(k))
        else:
            print('没有使用旧格式的字体')
        os.system('pause') 
    elif work == 8:
        cls()
        if not o_fontload:
            dupfont = {}
            font_list, font_name = getFileList(fontin)
            font_info = fontProgress(font_list, [font_name, {}, {}, []])
            del font_name
    # elif work == 6:
    #     cls()
    #     errorextl = []
    #     for k in font_info[0].keys(): 
    #         s = font_info[0][k]
    #         if s[3] and (path.splitext(s[0])[1][1:].lower() in ['ttf', 'otf']):
    #             errorextl.append((s[0], s[3]))
    #         elif (not s[3]) and (path.splitext(s[0])[1][1:].lower() in ['ttc', 'otc']):
    #             errorextl.append((s[0], s[3]))
    #     if len(errorextl) > 0:
    #         for s in errorextl:
    #             rightext = ''
    #             if s[1]:
    #                 rightext = path.splitext(s[0])[1].upper()
    #                 rightext = rightext[1:len(rightext) - 1] + 'C'
    #             else:
    #                 rightext = path.splitext(s[0])[1].upper()
    #                 rightext = rightext[1:len(rightext) - 1] + 'F'
    #             print('\033[1;31m[正确扩展名: {0}]\033[0m'.format(rightext))
    #             print('\033[1m{0}\033[0m\n'.format(s[0]))
    #     else:
    #         print('没有扩展名错误的字体')
    #     os.system('pause')
    elif work == 6:
        cLicense()
    elif work == 7:
        work2 = os.system('choice /M 您真的要退出吗 /C YN')
        if work2 != 2:
            exit()
    elif work == 9:
        cls()
        print('Subtitles-Videos Filename Matching (Experiment)')
        nomatching = 0
        sPath = input('字幕目录:').strip(' ').strip('\"').strip('\'')
        while not (path.exists(sPath) and path.isdir(sPath)):
            nomatching += 1
            if nomatching > 2: 
                print('回到主界面')
                break
            sPath = input('字幕目录:').strip(' ').strip('\"').strip('\'')
        else:
            if not nomatching > 2:
                logfiles = []
                for f in os.listdir(sPath):
                    if f[:11].lower() == 'subsrename_':
                        logfiles.append(f)
                if len(logfiles) > 0:
                    if os.system('choice /M 检测到重命名记录，您要撤销之前的重命名吗？') == 1:
                        logCount = 0
                        nomatching = 0
                        if len(logfiles) > 1:
                            print('请选择您要撤销到的时间点')
                            for i in range(0, len(logfiles)):
                                print('[{0}] {1}'.format(i, logfiles[i][11:]))
                            print('')
                            logCountC = input('请输入方括号内的数字:')
                            while not logCountC.isnumeric():
                                nomatching += 1
                                if nomatching > 2:
                                    logCount = -1
                                    break
                                print('输入有误，请重新输入；或者继续有误输入{0}次退出'.format(3 - nomatching))
                                logCountC = input('请输入方括号内的数字:')
                            else:
                                logCount = logCountC
                            if logCount < 0: 
                                logCount = -logCount
                                print('输入数字不能为负，已自动取相反数')
                            if logCount > len(logfiles) - 1: 
                                logCount = len(logfiles) - 1
                                print('输入数字过大，已自动调整为最后一项')
                            del logCountC
                        if logCount >= 0:
                            renCache = {}
                            for i in range(0, logCount + 1):
                                logfile = open(path.join(sPath, logfiles[i]), mode='r', encoding='utf-8')
                                for l in logfile.readlines():
                                    l = l.strip('\n').split('|')
                                    if len(l) < 2: continue
                                    if l[1] in renCache:
                                        renCache[path.basename(l[0])] = (renCache[l[1]][0], l[0])
                                        del renCache[l[1]]
                                    else:
                                        renCache[path.basename(l[0])] = (path.join(path.dirname(l[0]), l[1]), l[0])
                                logfile.close()
                            cls()
                            print('[重命名预览]\n')
                            for r in renCache.keys():
                                print("\033[1m\"{0}\"\033[0m".format(path.basename(renCache[r][0])))
                                print('>>>> \033[33;1m\"{0}\"\033[0m\n'.format(r))
                            print('')
                            if os.system('choice /M 是否要执行？') == 1:
                                nodel = False
                                for r in renCache.values():
                                    try:
                                        os.rename(r[0], r[1])
                                    except:
                                        print('\033[31;1m[ERROR]\033[0m 文件\"\033[1m{0}\033[0m\"重命名失败'.format(path.basename(r[0])))
                                        nodel = True
                                    else:
                                        print('''\033[32;1m[SUCCESS]\033[0m \"\033[1m{0}\033[0m\"
          >>>> \"\033[33;1m{1}\033[0m\"'''.format(path.basename(r[0]), path.basename(r[1])))
                            nomatching = 3
                            if not nodel:
                                for l in logfiles:
                                    try:
                                        os.remove(path.join(sPath, l))
                                    except:
                                        pass
                if nomatching > 2: 
                    print('回到主界面')
                else:
                    nomatching = 0
                    vPath = input('视频目录:').strip('\"').strip('\'')
                    while not (path.exists(vPath) and path.isdir(vPath)):
                        nomatching += 1
                        if nomatching > 2: 
                            print('回到主界面')
                            break
                        vPath = input('视频目录:').strip('\"').strip('\'')
                    if nomatching <= 2: nameMatching(sPath, getMediaFilelist(vPath))
                    else: print('回到主界面')
            else:
                print('回到主界面')
        os.system('pause')
    else: exit()
    cls()

try:
    loadMain()
    while 1: 
        loadMain(True)
except Exception as e:
    print('\n\033[31;1m')
    traceback.print_exc()
    print('\033[0m\n')
    if os.system('choice /M 您要重新启动本程序吗？') == 1:
        os.system('start cmd /c py \"{0}\"'.format(path.realpath(__file__)))
    else:
        os.system('pause')