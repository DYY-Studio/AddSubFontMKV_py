# -*- coding: UTF-8 -*-
# *************************************************************************
#
# 请使用支持 UTF-8 NoBOM 并最好带有 Python 语法高亮的文本编辑器
# Windows 7 的用户请最好不要使用 记事本 打开本脚本
#
# *************************************************************************

# 调用库，请不要修改
from typing import Optional
import fontTools.misc.encodingTools
from fontTools import ttLib, subset
from chardet.universaldetector import UniversalDetector
from chardet import detect_all
import os, sys, re, winreg, zlib, json, copy, traceback, shutil, configparser, time, locale, codecs
from os import path
from colorama import init
from datetime import datetime

#fontTools.misc.encodingTools.getEncoding()

# 初始化环境变量
# *************************************************************************
# 自定义变量
# 修改注意: Python的布尔类型首字母要大写 True 或 False；在名称中有单引号的，需要输入反斜杠转义如 \'
# *************************************************************************
# usingFF 同时提供FFmpeg封装
# 注意：如果要使用FFmpeg封装，必须有FFprobe
# 注意：如果设定为False，在mkvmerge不可用时，会检测FFmpeg的可用性
usingFF = False
# *************************************************************************
# extlist 可输入的视频媒体文件扩展名
extlist = 'mkv;mp4;mts;mpg;flv;mpeg;m2ts;avi;webm;rm;rmvb;mov;mk3d;vob'
# *************************************************************************
# no_extcheck 关闭对extlist的扩展名检查来添加一些可能支持的格式
no_extcheck = False
# *************************************************************************
# subl_local 设置sublangs默认语言（如果有对应语言包则显示对应语言） 不区分大小写
subl_local = 'zh_CN'
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
# exact_lost 在列出字幕所需字体时，是否显示缺字的具体字幕文件
#   True    显示
#   False   不显示
exact_lost = False
# *************************************************************************
# char_lost 在列出字幕所需字体时，是否检查系统中已有字体是否缺字
#   True    检查
#   False   不检查
char_lost = False
# *************************************************************************
# char_compatible 是否采用更高兼容性的子集化方式
# 由于目前各种奇妙的子集化问题层出不穷，根据建议添加了本选项，默认启用
# 在该选项下，以下范围的字符会被全部保留：0-9(仅当文本中含有任意一个数字的时候)，a-z A-Z(任何时候)
# 如果您希望追求极致小的字体，可以选择关闭，测试时在VSFilterMod上未出现渲染问题
#   True    启用
#   False   不启用
char_compatible = True
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
# matchStrict 严格匹配
#   True   媒体文件名必须在字幕文件名的最前方，如'test.mkv'的字幕可以是'test.ass'或是'test.sub.ass'，但不能是'sub.test.ass'
#   False  只要字幕文件名中有媒体文件名就行了，不管它在哪
matchStrict = True
# *************************************************************************
# startQuiet 静默启动
#   True   不显示字体重复、字体名称修正信息
#   False  反之
startQuiet = True
# *************************************************************************
# warningStop 对有可能子集化失败的字体的处理方式
#   True   不子集化
#   False  子集化（如果失败，会保留原始字体；如果原始字体是TTC/OTC，则会保留其中的TTF/OTF）
warningStop = False
# ignoreLost 遇到字体内没有的字符时的处理方式
#   True   删掉没有的字符，继续子集化
#   False  中断（=errorStop）
ignoreLost = True
# noRequestFont 遇到字幕缺少字体时的处理方式
#   True   不询问缺少的字体
#   False  询问
noRequestFont = False
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
# embeddedFontExtract 在分析ASS/SSA文件时顺便抽出内嵌的字体文件
# 不建议开启该功能，应当保持常关
embeddedFontExtract = False
# *************************************************************************

# 以下变量谨慎更改
# subext 可输入的字幕扩展名，按照python列表语法
subext = ['ass', 'ssa', 'srt', 'sup', 'idx']
# lcidfil
# LCID过滤器，用于选择字体名称所用语言
# 结构为：{ platformID(Decimal) : { LCID : textcoding } }
# Textcoding将用于进行平台优先语言筛选
# 详情可在 https://docs.microsoft.com/en-us/typography/opentype/spec/name#platform-encoding-and-language-ids 查询
# 或见 https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
lcidfil = {
3: {
    2052: 'gbk',
    1042: 'euc-kr',
    1041: 'shift_jis',
    1033: 'utf-16-be',
    1028: 'big5',
    0: 'utf-16-be'
}, 2: {
    0: 'ascii',
    1: 'utf_16_be',
    2: 'latin1'
}, 1: {
    33: 'gbk',
    23: 'euc_kr',
    19: 'big5',
    11: 'shift_jis',
    0: 'mac-roman'
}, 0: {
    0: 'utf-16-be',
    1: 'utf-16-be',
    2: 'utf-16-be',
    3: 'utf-16-be',
    4: 'utf-16-be',
    5: 'utf-16-be',
    6: 'utf-16-be'
}}

# 将fontTools文本编解码模块中gb2312等效参数改为gbk
fontTools.misc.encodingTools._encodingMap[3][3] = 'gbk'
fontTools.misc.encodingTools._encodingMap[1][25] = 'gbk'

# lcidfil = fontTools.misc.encodingTools._encodingMap

# 以下环境变量不应更改
# 编译 style行 搜索用正则表达式
style_read = re.compile('.*\nStyle:.*')
cillegal = re.compile(r'[\\/:\*"><\|]')
# 切分extlist列表
extlist = [s.strip(' ').lstrip('.').lower() for s in extlist.split(';') if len(s) > 0]
# 切分fontin列表
fontin = [s.strip(' ') for s in fontin.split('?') if len(s) > 0]
fontin = [s for s in fontin if path.isdir(s)]
if o_fontload and not fontload:
    fontload = True
langlist = {}
extsupp = []
dupfont = {}
insteadFF = False
init()

iso639_all = {}
preferLang = subl_local
preferEncoding = codecs.lookup(locale.getlocale()[1]).name

# 读取外部INI文件内的环境变量值以取代程序内预设
if path.exists(path.join(path.dirname(__file__), 'ASFMKVpy.ini')):
    conf = configparser.ConfigParser()
    conf.read('ASFMKVpy.ini', encoding='utf-8')
    settingList = ["usingFF", "extlist", "no_extcheck", "mkvout", "assout", "fontout", "fontin", "exact_lost", "char_lost", "char_compatible", "fontload", "o_fontload", "s_fontload", "f_priority", "notfont", "matchStrict", "startQuiet", "warningStop", "ignoreLost", "errorStop", "rmAssIn", "rmAttach", "v_subdir", "s_subdir", "copyfont", "resultw", "subl_local"]
    strSettings = ['mkvout', 'assout', 'fontout', 'subl_local']
    for k in conf['settings']:
        if k in settingList:
            try:
                if k in strSettings:
                    exec('{} = \"{}\"'.format(k, conf['settings'][k].strip("\"\' ")))
                elif k == 'extlist':
                    if len(conf['settings'][k]) > 0:
                        extlist.clear()
                        for i in conf['settings'][k].split(','):
                            extlist.append(i.strip(' ').lower())
                elif k == 'fontin':
                    if len(conf['settings'][k]) > 0:
                        fontin.clear()
                        for i in conf['settings'][k].split('?'):
                            if path.isdir(i.strip(' ')):
                                fontin.append(i.strip(' '))
                else:
                    if conf['settings'][k].strip(' ') == '1':
                        exec(f'{k} = True')
                    else:
                        exec(f'{k} = False')
            except:
                traceback.print_exc()
                time.sleep(5)
    conf.clear()
    del conf


def showColorBool(text: str, tf: bool, tColor: int = 31, fColor: int = 33) -> str:
    if tf: return f'\033[1;{tColor}m{text}\033[0m'
    else: return f'\033[1;{fColor}m{text}\033[0m'


def getISOLangs():
    '''读取APPDATA中的ISO639语言代码到系统语言呈现的语言名称对照表'''
    global iso639_all, preferLang
    langDir = path.join(os.getenv('APPDATA'), 'ASFMKVpy', 'isoLang')
    if path.isdir(langDir):
        langDirList = []
        for d in os.listdir(langDir):
            if path.isdir(path.join(langDir, d)):
                if path.exists(path.join(langDir, d, 'language.json')):
                    langDirList.append(d)
        langSubDir = ''
        if preferLang in langDirList:
            langSubDir = preferLang
        else:
            for d in langDirList:
                if preferLang.split('_')[0] == d.split('_')[0]:
                    langSubDir = d
        if len(langSubDir) == 0: return None

        print('读取ISO-639语言翻译列表...{}'.format(str.ljust(d, 10)), end='\r', flush=True)
        isoJSON = open(path.join(langDir, langSubDir, 'language.json'), mode='r', encoding='utf-8')
        allISO = json.load(isoJSON)
        isoJSON.close()

        for lc in allISO.keys():
            if not len(lc) > 3:
                if iso639_all.get(langSubDir, False):
                    iso639_all[langSubDir][lc] = allISO[lc]
                else:
                    iso639_all[langSubDir] = {lc: allISO[lc]}
        
        print('读取ISO-639语言翻译列表...{}'.format(str.ljust('完成', 10)))
    pass

# getISOLangs()

def fontlistAdd(s: str, fn: tuple, fontlist: dict) -> dict:
    if len(s) > 0:
        si = 0
        fn = (fn[0].lstrip('@'), fn[1], fn[2])
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
        if assfont.get(key, False):
            # if assfont[key][4] != assfont2[key][4] or assfont[key][5] != assfont2[key][5]:
            assfont[key] = [
                assfont[key][0] + ''.join([ks for ks in assfont2[key][0] if ks not in assfont[key][0]]),
                assfont[key][1] + ''.join(
                    ['|' + ks for ks in assfont2[key][1].split('|') if ks.lower() not in assfont[key][1].lower()]),
                assfont[key][2] + ''.join(
                    ['|' + ks for ks in assfont2[key][2].split('|') if ks.lower() not in assfont[key][2].lower()]),
                list(set(assfont[key][3] + assfont2[key][3])),
                list(set(assfont[key][4] + assfont2[key][4]))
                ]
        else:
            assfont[key] = assfont2[key]
    del assfont2
    return assfont


def eventTagsSplit(assInfo: dict, changeOnly: bool = True) -> dict:
    '''
对ASS/SSA的Events逐行进行分析
判断每行每个部分使用的字体和粗体斜体情况

目前暂时只支持assInfo形式的输入
changeOnly: 是否只输出字体在行中发生变化的行（带有 \\fn \\r \\b \\i 特效）

返回格式如下
{ 
    assInfo['Events']索引 (int): 
        [
            {
            'Text': <文本 (str)>, 
            'Fontname': <字体名 (str)>, 
            'Bold': <粗体 (bool)>, 
            'Italic': <斜体 (bool)>
            },
            ...
        ]
}
    '''

    events = assInfo['Events']
    lines = {}

    for i in range(0, len(events)):
        if events[i]['Event'] != 'Dialogue':
            continue
        eventftext = events[i]['Text']
        effectDel = r'\{.*?\}|\\[nNhs]|\n'
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
        eventfont = assInfo['Styles'].get(events[i]['Style'])
        if eventfont is None:
            eventfont = assInfo['Styles'].get(events[i]['Style'].lstrip('*'))
        if eventfont is None:
            if len(re.sub(effectDel, '', eventftext)) > 0:
                raise Exception('未定义 行{0} 所用样式"{1}\"'.format(i, events[i]['Style'].lstrip('*')))
            continue

        # 粗体、斜体标签处理
        splittext = []
        lfn = eventfont['Fontname']
        lfi = eventfont['Italic']
        lfb = eventfont['Bold']
        # 首先查找蕴含有启用粗体/斜体标记的特效标签
        if re.search(r'\{.*?(?:\\b|\\i|\\fn|\\r).*?\}', eventftext) is not None:
            lastfind = 0
            allfind = re.findall(r'\{.*?\}', eventftext)
            eventftext2 = eventftext
            # 在所有特效标签中寻找
            # 然后分别确认该特效标签的适用范围，以准确将字体子集化
            for sti in range(0, len(allfind)):
                st = allfind[sti]
                ibopen = re.search(r'\\b|\\i|\\fn|\\r', st)
                if ibopen is not None:
                    stfind = eventftext2.find(st) + lastfind
                    addbold = lfb
                    additalic = lfi
                    # 不管有没有 \r 标签，先获取了再说
                    rstylel = re.findall(r'\\r.*?[\\\}]', st)
                    rfn = lfn
                    rfb = lfb
                    rfi = lfi
                    pos_r = -1
                    # 如果有就处理，没有就算了
                    if len(rstylel) > 0: 
                        pos_r = st.rfind(rstylel[-1])
                        rstyle = rstylel[-1].split('\\')[1].lstrip('r').rstrip('}')
                        if rstyle in assInfo['Styles']:
                            rfn = assInfo['Styles'][rstyle]['Fontname']
                            rfi = assInfo['Styles'][rstyle]['Italic']
                            rfb = assInfo['Styles'][rstyle]['Bold']
                        else:
                            rfn = eventfont['Fontname']
                            rfi = eventfont['Italic']
                            rfb = eventfont['Bold']
                    else: rstyle = []
                    del rstylel

                    # 统一获取标签位置
                    pos_b1 = -1
                    fbopen = re.findall(r'\\b[7-9]00|\\b1', st)
                    if len(fbopen) > 0:
                        pos_b1 = st.rfind(fbopen[-1])
                    pos_b0 = max([st.rfind('\\b0'), st.rfind('\\b\\'), st.rfind('\\b}')])
                    pos_i1 = st.rfind('\\i1')
                    pos_i0 = max([st.rfind('\\i0'), st.rfind('\\i\\'), st.rfind('\\i}')])
                    pos_fn = st.rfind('\\fn')

                    # 处理 \b 标签
                    if max([pos_b1, pos_b0]) > -1:
                        if pos_b1 > max([pos_b0, pos_r]):
                            addbold = '1'
                        elif pos_b0 > max([pos_b1, pos_r]):
                            addbold = '0'
                        elif pos_r > max([pos_b1, pos_b0]):
                            addbold = rfb

                    # 处理 \i 标签
                    if max([pos_i1, pos_i0]) > -1:
                        if pos_r > max([pos_i1, pos_i0]):
                            additalic = rfi
                        elif pos_i1 > max([pos_r, pos_i0]):
                            additalic = '1'
                        elif pos_i0 > max([pos_i1, pos_r]):
                            additalic = '0'

                    # 处理 \fn 标签 和 \r 标签
                    if pos_fn > -1:
                        fnf = re.findall(r'\\fn.*?[\\\}]', st)
                        if len(fnf) > 0 and pos_fn > pos_r:
                            fnn = fnf[-1].split('\\')[1].rstrip('}')[2:].lstrip('@')
                            if len(fnn) == 0:
                                fnn = eventfont['Fontname']
                            elif len(assInfo['Subset']) > 0:
                                if assInfo['Subset'].get(fnn) is not None:
                                    fnn = assInfo['Subset'][fnn]
                                    del lsplit
                            if stfind > 0:
                                if len(splittext) == 0:
                                    splittext.append([eventftext[:stfind], lfi, lfb, lfn])
                                else:
                                    splittext[-1][0] = eventftext[lastfind:stfind]
                            splittext.append([eventftext[stfind:], additalic, addbold, fnn])
                            lfn = fnn
                            del fnn
                        elif pos_fn < pos_r:
                            if stfind > 0:
                                if len(splittext) == 0:
                                    splittext.append([eventftext[:stfind], lfi, lfb, lfn])
                                else:
                                    splittext[-1][0] = eventftext[lastfind:stfind]
                            lfn = rfn
                            splittext.append([eventftext[stfind:], additalic, addbold, lfn])
                    # 在处理 \fn 标签时已经处理了 \r 标签了，但是如果没有 \fn 标签，则进入这里的 \r 标签处理
                    elif pos_r > -1:
                        if stfind > 0:
                            if len(splittext) == 0:
                                splittext.append([eventftext[:stfind], lfi, lfb, lfn])
                            else:
                                splittext[-1][0] = eventftext[lastfind:stfind]
                        lfn = rfn
                        splittext.append([eventftext[stfind:], additalic, addbold, lfn])
                    # 如果 \fn 标签 和 \r 标签 都没有，那就直接写
                    else:
                        if len(splittext) == 0:
                            if stfind > 0:
                                splittext.append([eventftext[:stfind], lfi, lfb, lfn])
                            splittext.append([eventftext[stfind:], additalic, addbold, lfn])
                        else:
                            if stfind > 0:
                                splittext[-1][0] = eventftext[lastfind:stfind]
                            splittext.append([eventftext[stfind:], additalic, addbold, lfn])
                    lastfind = stfind
                    eventftext2 = eventftext[stfind:]
                    lfb = addbold
                    lfi = additalic
                else:
                    continue
                    # del rfn, rfi, rfb, pos_b0, pos_b1, pos_r, pos_fn, pos_i1, pos_i0
        if len(splittext) == 0:
            splittext.append([eventftext, lfi, lfb, lfn])

        eventSplit = []
        for l in splittext:
            eventSplit.append({'Text': l[0], 'Fontname': l[3], 'Bold': int(l[2]), 'Italic': int(l[1])})
        del splittext

        lines[i] = eventSplit

    if changeOnly:
        chLines = {}
        for k in lines.keys():
            if len(lines[k]) > 1:
                chLines[k] = lines[k]
        return chLines
    else:
        return lines


def assAnalyseV2(assPath: str, keepEmbedded: bool = False) -> dict:
    '''
assAnalyse V2 function

assPath: ASS/SSA文件路径 (str)
keepEmbedded: 是否保留ASS/SSA内嵌文件（字体/图像） (bool)

返回的变量结构
assInfo:
{
    Lines: ASS/SSA标签位于文件的哪一行，如果没有该标签，则行数为-1
        {
            <标签名(str)>: <行数(int)>,
            ...
        }
    
    Styles: ASS/SSA样式表
        {
            <样式名(str)>: 
            {
                <样式信息(str)>: <内容(str)>,
                ...
            },
            <样式名(str)>: 
            {
                <样式信息(str)>: <内容(str)>,
                ...
            }
            ...
        }
    
    Events: ASS/SSA事件表
        [
            {<事件信息(str)>: <内容(str)>}, 
            ...
        ]
    
    Fonts: ASS/SSA字体附件
        {
            <附件文件名(str)>: <内容(str)>, 
            ...
        }
    Graphics: ASS/SSA图像附件（同）

    *Fonts和Graphics的内容均已经过解码，可以直接以二进制输出为对应的文件*

    Garbage: Aegisub留在ASS/SSA文件内的一些信息
        {
            '#GarbageWholeTag#' : <完整标签名>,
            <名称(str)>: <内容(str)>,
            ...
        }
}
    '''
    assTagLines = {
        'Info': -1, 
        'Styles': -1, 
        'Events': -1, 
        'Fonts': -1, 
        'Graphics': -1, 
        'Garbage': -1
        }
    
    assInfo = {
        'Lines': assTagLines, 
        'Info': {}, 
        'Styles': {}, 
        'Events': [], 
        'Fonts': {}, 
        'Graphics': {}, 
        'Garbage': {},
        'Subset': {}
        }

    ass = open(assPath, mode='rb')
    detector = UniversalDetector()
    for b in ass:
        detector.feed(b)
        if detector.done:
            break
    ass.close()

    ass = open(assPath, mode='r', encoding=detector.result['encoding'].replace('gb2312', 'gbk'))
    print(f'\"{path.basename(assPath)}\" - {ass.encoding.upper()}')
    del detector

    assLines = ass.readlines()
    ass.close()

    def assDecodeUUE(embedded: list[str]) -> bytes:
        '''
    解码UUEncoding
        '''
        if 'fontname' in embedded[0] or 'filename' in embedded[0]:
            embedded.pop(0)
        content = [i-33 for i in ''.join(embedded).strip(' \n').encode('ascii') if i != '\n']
        size = len(content)
        
        f6bytes = []

        for i in range(0, size, 4):
            if size - i < 4: break
            bs = (content[i]<<18)|(content[i+1]<<12)|(content[i+2]<<6)|(content[i+3])
            f6bytes.extend([(bs>>16)&255, (bs>>8)&255, bs&255])
        
        if size % 4 == 2:
            bs = ((content[-2]<<6)|(content[-1]))>>4
            f6bytes.extend([bs&255])
        elif size % 4 == 3:
            bs = ((content[-2]<<12)|(content[-1]<<6)|(content[-1]))>>2
            f6bytes.extend([(bs>>8)&255, bs&255])
        
        return bytes(f6bytes)


    def stripListItem(l: list[str], strStrip: str = ' \n') -> list[str]:
        return [s.strip(strStrip) for s in l]


    def assStylesEventsRead(tagLine: int) -> list[dict]:
        if tagLine == -1: return []
        # result: [{<Field 1>: <Content>, <Field 2>: <Context>}(Line 1), {...}(Line ...)]
        lines = []
        isEvent = False
        if assLines[tagLine].strip('[]\n') == 'Events':
            isEvent = True
        
        offset = 1
        lineSp = assLines[tagLine+offset].split(':')
        while(len(lineSp) == 1):
            offset += 1
            lineSp = assLines[tagLine+offset].split(':')
        lineFormat = stripListItem(':'.join(lineSp[1:]).strip(' ').split(','))

        if isEvent and 'Text' not in lineFormat:
            raise Exception(f'ASS/SSA格式错误 - 缺少Format行 - \"{path.basename(assPath)}\"')

        for i in range(tagLine + offset + 1, len(assLines)):

            line = assLines[i].split(':')
            if len(line) == 1: 
                if assLines[i].strip(' ')[0] == '[':
                    break
                else:
                    continue

            line = stripListItem(':'.join(line[1:]).strip(' ').split(','))

            singleLine = {}
            if isEvent:
                singleLine['Event'] = assLines[i].split(':')[0]

            for j in range(len(line)):
                if isEvent and lineFormat[j] == 'Text':
                    singleLine[lineFormat[j]] = ','.join(line[j:])
                    break
                else:
                    singleLine[lineFormat[j]] = line[j]

            if (isEvent and len(singleLine) != len(lineFormat) + 1) or (not isEvent and len(singleLine) != len(lineFormat)):
                raise Exception(f'ASS/SSA格式错误 - 缺少逗号分隔 - 位于行{i+1} 文件\"{path.basename(assPath)}\"')
            
            lines.append(singleLine)
            del singleLine
        del lineFormat
        
        return lines


    def assStyles2Dict(lines: list[dict]) -> dict:
        styleDict = {}
        for i in lines:
            styleName = i['Name']
            i.pop('Name')
            styleDict[styleName] = i
        del lines
        return styleDict
    

    def assInfoRead(tagLine: int) -> dict:
        if tagLine == -1: return {}
        lines = {}

        for i in range(tagLine + 1, len(assLines)):

            line = assLines[i].split(':')
            if assLines[i].lstrip(' ')[0] == ';': 
                if 'font subset' in assLines[i].lower():
                    fnRe = stripListItem(assLines[i].split(':')[1].split('-'))
                    assInfo['Subset'][fnRe[0]] = '-'.join(fnRe[1:])
                continue

            if len(line) == 1: 
                if i >= len(assLines) or assLines[i].strip(' ')[0] == '[':
                    break
                else:
                    continue
            lines[line[0]] = ':'.join(line[1:]).strip(' \n')
    
        return lines
    

    def assEmbeddedRead(tagLine: int) -> dict:
        if tagLine == -1: return {}
        lines = {}

        i = tagLine
        while i < len(assLines):
            i+=1

            if assLines[i].lstrip(' ')[:8].lower() in ['fontname', 'filename']: 
                filename = assLines[i].lstrip(' ')[8:].strip(': \n')
                content = []
                i+=1
                while i < len(assLines) and assLines[i].lstrip(' ')[:8] not in ['fontname', 'filename'] and re.match(r'^\[.*\]$', assLines[i].strip(' ')) == None:
                    if assLines[i].strip(' \n') != '':
                        content.append(assLines[i].strip(' \n'))
                    i+=1
                lines[filename] = assDecodeUUE(content)
                del content
            
            if i >= len(assLines) or assLines[i].strip(' ')[0] == '[':
                break
            elif assLines[i].lstrip(' ')[:8] in ['fontname', 'filename']:
                i-=1
            continue
    
        return lines

    for i in range(len(assLines)):
        l = assLines[i].strip(' \n')
        if len(l) > 0 and l[0] == '[' and l[-1] == ']':
            l = l.strip('[]')
            lSpace = l.split(' ')
            if len(lSpace) > 1:
                assTagLines[lSpace[-1]] = i
            else:
                assTagLines[l] = i
    
    assInfo['Lines'] = assTagLines
    assInfo['Info'] = assInfoRead(assTagLines['Info'])
    assInfo['Garbage'] = assInfoRead(assTagLines['Garbage'])
    if assTagLines['Garbage'] > -1:
        assInfo['Garbage']['#GarbageWholeTag#'] = assLines[assTagLines['Garbage']].strip(' []')
    assInfo['Styles'] = assStyles2Dict(assStylesEventsRead(assTagLines['Styles']))
    assInfo['Events'] = assStylesEventsRead(assTagLines['Events'])
    if keepEmbedded:
        assInfo['Fonts'] = assEmbeddedRead(assTagLines['Fonts'])
        assInfo['Graphics'] = assEmbeddedRead(assTagLines['Graphics'])

    del assLines

    return assInfo


def assFontList(assInfo: dict, eventSplit: dict) -> dict:
    '''
fontList: {(字体名<str>, 斜体<int>, 粗体<int>): '字符<str>'}
    '''
    fontList = {}

    for i in eventSplit.keys():
        for l in eventSplit[i]:
            fn = l['Fontname'].lstrip('@')
            l['Text'] = re.sub(r'\{.*?\\.*?\}', '', l['Text'])
            if len(assInfo['Subset']) > 0 and fn in assInfo['Subset']:
                fn = assInfo['Subset'][fn]
            flIndex = (fn, abs(int(l['Italic'])), abs(int(l['Bold'])))
            if flIndex in fontList:
                for char in set(l['Text']):
                    if char not in fontList[flIndex]:
                        fontList[flIndex] += char
            else:
                fontList[flIndex] = ''.join(set(l['Text']))
    
    return fontList

def assFileWrite(filePath: str, assInfo: dict, subsetRecover: bool = True, keepGarbage: bool = True, keepEmbedded: bool = True, ignoreExists: bool = False, Textencoding: str = 'utf-8-sig'):
    if not ignoreExists and path.isfile(filePath):
        raise FileExistsError()
    file = open(filePath, mode='w', encoding=Textencoding)

    def assEncodeUUE(content: bytes, basename: str = '', isFont: bool = True) -> list[str]:
        '''
    对ASS/SSA内嵌字体进行UUEncoding以嵌入
        '''
        # 获取字节大小
        size = len(content)
        # 定义返回变量，同时记录第一行文件名
        if len(basename) != 0:
            if isFont:
                result = [f'fontname: {basename}\n', '']
            else:
                result = [f'filename: {basename}\n', '']
        else:
            result = ['']

        # 定义保证单行80字符的函数
        def addResult(s):
            if len(result[len(result) - 1]) < 80:
                result[len(result) - 1] += s
            else:
                if len(result) > 0:
                    result[-1] += '\n'
                result.append(s)

        # 对文件按照规范进行UUEncoding
        for i in range(0, size, 3):
            if size - i < 3: break
            # 生成一个24位二进制数
            bs = (content[i]<<16)|(content[i+1]<<8)|(content[i+2])
            # 将24位二进制分为4组6位二进制
            f6bits = [(bs>>18), (bs>>12), (bs>>6), bs]

            # 将每个6位二进制 + 33(10进制)后转换为字符，交给函数写入
            addResult(''.join([chr((s&63)+33) for s in f6bits]))

        # 对于字节数不能被三整除的情况，按标准对末尾几个字符进行处理
        if size % 3 != 0:
            if size % 2 != 0:
                # 字节数为奇数，则将剩下的最后一位*256(0x100)，然后只保留前12位，分为两个二进制组
                bs = content[-1]<<4
                f6bits = [(bs>>6), bs]
            else:
                # 字节数为偶数，则将剩下的最后两位*65536(0x10000)，然后只保留前18位，分为三个二进制组
                bs = (content[-2]<<8)|content[-1]
                bs = bs<<2
                f6bits = [(bs>>12), (bs>>6), bs]

            # 转换为字符
            addResult(''.join([chr((s&63)+33) for s in f6bits]))
        
        if result[-1][-1] != '\n':
            result[-1] += '\n'
        
        return result

    assLines = []
    
    def subsetWrite():
        for k in assInfo['Subset'].keys():
            assLines.append('; Font Subset: {0} - {1}\n'.format(k ,assInfo['Subset'][k]))


    if len(assInfo['Info']) > 0:
        assLines.append('[Script Info]\n')
        if not subsetRecover and len(assInfo['Subset']) > 0:
            subsetWrite()
        for k in assInfo['Info'].keys():
            assLines.append('{0}: {1}\n'.format(k, assInfo['Info'][k]))
        assLines.append('\n')
    elif not subsetRecover and len(assInfo['Subset']) > 0:
        assLines.append('[Script Info]\n')
        subsetWrite()
        assLines.append('\n')


    if keepGarbage and len(assInfo['Garbage']) > 0:
        assLines.append('[{}]\n'.format(assInfo['Garbage']['#GarbageWholeTag#']))
        for k in assInfo['Garbage'].keys():
            assLines.append('{0}: {1}\n'.format(k, assInfo['Garbage'][k]))
        assLines.append('\n')
    

    if len(assInfo['Styles']) > 0:
        if path.splitext(filePath)[1][1:].lower() == 'ssa':
            assLines.extend(['[V4 Styles]\n', 'Format: Name'])
        else:
            assLines.extend(['[V4+ Styles]\n', 'Format: Name'])
        for f in list(assInfo['Styles'].values())[0].keys():
            assLines[-1] += (f', {f}')
        assLines[-1] += ('\n')
        for k in assInfo['Styles'].keys():
            assLines.append(f'Style: {k}')
            for k2 in assInfo['Styles'][k].keys():
                v = assInfo['Styles'][k][k2]
                if subsetRecover and k2 == 'Fontname':
                    v = assInfo['Subset'].get(v, v)
                assLines[-1] += f',{v}'
            assLines[-1] += '\n'
        assLines.append('\n')
    

    if len(assInfo['Events']) > 0:
        assLines.extend(['[Events]\n', 'Format: '])
        for f in list(assInfo['Events'][0].keys())[1:]:
            assLines[-1] += f'{f}, '
        assLines[-1] = assLines[-1].rstrip(', ')
        assLines[-1] += '\n'
        for i in assInfo['Events']:
            assLines.append('{}: '.format(i['Event']))
            for k in i.keys():
                if k != 'Event':
                    assLines[-1] += (i[k] + ',')
            assLines[-1] = assLines[-1][:-1] + '\n'
        assLines.append('\n')
    

    if keepEmbedded and len(assInfo['Fonts']) > 0:
        assLines.append('[Fonts]\n')
        for k in assInfo['Fonts'].keys():
            assLines.extend(assEncodeUUE(assInfo['Fonts'][k], k) + ['\n'])
        assLines.append('\n')


    if keepEmbedded and len(assInfo['Graphics']) > 0:
        assLines.append('[Graphics]\n')
        for k in assInfo['Graphics'].keys():
            assLines.extend(assEncodeUUE(assInfo['Graphics'][k], k) + ['\n'])
        assLines.append('\n')
    
    file.writelines(assLines)
    file.close()

def getFontFileList(customPath: list = [], noreg: bool = False):
    """
获取字体文件列表

接受输入
  customPath: 用户指定的字体文件夹
  noreg: 只从用户提供的customPath获取输入

将会返回
  filelist: 字体文件清单 [[ 字体绝对路径, 读取位置（'': 注册表, '0': 自定义目录） ], ...]
    """
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
                    if path.exists(p):
                        filelist.append([p, ''])
        except:
            pass
        for i in range(fontkey_num):
            # 从 系统字体注册表 读取
            k = winreg.EnumValue(fontkey, i)[1]
            pk = path.join(r'C:\Windows\Fonts', k)
            if path.exists(pk):
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

    # print(font_name)
    # os.system('pause')
    return filelist


charEx = re.compile(r'[\U00020000-\U0003134F\u3400-\u4DBF\u2070-\u2BFF\u02B0-\u02FF\uE000-\uF8FF\u0080-\u00B1\u00B4-\u00BF]')
def fnReadCheck(ttFont: ttLib.TTFont, index: int, fontpath: str):
    """
修正字体内部名称解码错误

现在包含两种方式:
    快速乱码检查 >> 从字体Name表中寻找可能的其它文本编码 >> 找不到就用chardet尝试修复
    解码错误 >> fnReadCorrect尝试修复
    """
    global charEx, startQuiet
    indexs = len(ttFont['name'].names)
    name = ttFont['name'].names[index]
    isWarningFont = False
    c = name.getEncoding()
    try:
        try:
            s = name.toStr()
        except:
            s, c = fnReadCorrect(ttFont, index, fontpath)
            isWarningFont = True
        s2 = s
        
        if re.search(charEx, s):
            c = ''
            cs = {}
            for ci in range(0, indexs):
                if not ttFont['name'].names[ci].isUnicode():
                    cs[ttFont['name'].names[ci].getEncoding().lower()] = True
            for ec in cs.keys():
                if ec not in ['mac_roman', 'utf_16_be']:
                    try:
                        s2 = name.toBytes().decode(ec)
                        c = ec
                        break
                    except:
                        pass
            sc = 1
            if cs.get('mac_roman') and len(c) == 0:
                todetectc = [''.join(h) for h in re.findall(r'(\\x..\\)\\|(\\x..[\u0041-\u005B\u005D-\u007A]?)', str(name.toBytes())[2:-1]) if len(h) > 0]
                todetect = []
                for td in todetectc:
                    if len(td) > 4:
                        todetect.append(int(td[2:][:2], 16))
                        todetect.append(ord(td[4:][:1]))
                    elif len(td) > 2:
                        todetect.append(int(td[2:], 16))
                todetect = todetect + todetect
                da = [d for d in detect_all(bytes(todetect)) if d['encoding'] is not None]
                if len(da) > 0: da = [d for d in da if d['encoding'].lower() not in ['mac_roman', 'windows-1252', 'iso-8859-1', 'ascii']]
                if len(da) > 0: 
                    try:
                        s2 = name.toBytes().decode(da[0]['encoding'])
                        c = da[0]['encoding']
                        sc = da[0]['confidence']
                    except:
                        pass
            if not re.search(charEx, s2):
                if len(c) > 0:
                    if not startQuiet: print('\n\033[1;33m尝试修正字体\"{0}\"文本编码 \"{1}\" >> \"{2}\"({4:.1f}%) = \"{3}\"\033[0m'
                        .format(path.basename(fontpath), name.getEncoding(), c.lower(), s2, sc * 100))
                    return s2, c, isWarningFont
                else:
                    return s, c, isWarningFont
            else:
                if not startQuiet: print('\n\033[1;33m字体\"{0}\"文本编码修复失败 \"{1}\" = \"{2}\"\033[0m'.format(path.basename(fontpath), name.getEncoding(), s))
                return s, c, isWarningFont
            # for dc in detect_all(name.toBytes()):
            #     if dc['encoding'] is not None:
            #         print(dc)
            #         if not dc['encoding'].lower() in ['windows-1252', 'ascii', 'iso-8859-1']:
            #             c = dc['encoding']
            #             break
        else: return s, c, isWarningFont
    except Exception as e:
        print(e)
        s, c = fnReadCorrect(ttFont, index, fontpath)
        return s, c, True


def fnReadCorrect(ttfont: ttLib.ttFont, index: int, fontpath: str):
    """修正字体内部名称解码错误"""
    global startQuiet
    name = ttfont['name'].names[index]
    c = 'utf-16-be'
    namestr = name.toBytes().decode(c, errors='ignore')
    try:
        if not startQuiet: print('')
        if len([i for i in name.toBytes() if i == 0]) > 0:
            namebyteL = name.toBytes()
            for bi in range(0, len(namebyteL), 2):
                if namebyteL[bi] != 0:
                    if not startQuiet: print('\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
                    return namestr, '\u0000'
            namebyte = b''.join([bytes.fromhex('{:0>2}'.format(hex(i)[2:])) for i in name.toBytes() if i > 0])
            namestr = namebyte.decode(name.getEncoding())
            c = name.getEncoding()
            if not startQuiet: print('\033[1;33m已修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
            # os.system('pause')
        else:
            namestr = name.toBytes().decode('utf-16-be', errors='ignore')
        # 如果没有 \x00 字符，使用 utf-16-be 强行解码；如果有，尝试解码；如果解码失败，使用 utf-16-be 强行解码
    except:
        if not startQuiet: print('\033[1;33m尝试修正字体\"{0}\"名称读取 >> \"{1}\"\033[0m'.format(path.basename(fontpath), namestr))
    return namestr, c



def outputSameLength(s: str) -> str:
    length = 0
    output = ''
    for i in range(len(s) - 1, -1, -1):
        si = s[i]
        if 65281 <= ord(si) <= 65374 or ord(si) == 12288 or ord(si) not in range(33, 127):
            length += 2
        else:
            length += 1
        if length >= 50:
            output = '..' + output
            break
        else:
            output = si + output
    return output + ''.join([' ' for _ in range(0, 60 - length)])


# font_info 列表结构
#   [ font_name, font_n_lower, font_family, warning_font, font_all ]
# font_name 词典结构
#   { 字体名称 : [ 字体绝对路径 , 字体索引 (仅用于TTC/OTC; 如果是TTF/OTF，默认为0) ] }
# font_all 字典结构
#   { 字体绝对路径: { 字体索引: (字体名称(dict), 斜体, 粗体, 字体样式(dict), 字体家族名称(dict)) } }
# dupfont 词典结构
#   { 重复字体名称 : [ 字体1绝对路径, 字体2绝对路径, ... ] }
def fontProgress(fl: list, font_info: list = [{}, {}, {}, [], {}], overwrite: bool = False, usingCache: bool = True) -> list:
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
    global dupfont, startQuiet
    warning_font = font_info[3]
    font_family = font_info[2]
    font_n_lower = font_info[1]
    f_n = font_info[0]
    f_all = font_info[4]

    # 字体缓存所用词典
    # 格式为
    # {dirCRCKey: {
    #   fontCRCKey: {
    #       字体索引: (字体名称(dict), 斜体, 粗体, 字体样式(dict), 字体家族名称(dict))
    #   }
    # }}
    fontCache = {}

    fontCacheN = {}
    dirKeyList = {}
    appdata = os.getenv('APPDATA')
    fontCacheDir = path.join(appdata, 'ASFMKVpy')
    if not path.isdir(fontCacheDir):
        os.mkdir(fontCacheDir)

    # print(fl)
    flL = len(fl)
    print('\033[1;32m正在读取字体信息...\033[0m')
    for si in range(0, flL):
        s = fl[si][0]
        fromCustom = False
        fl[si][1] = str(fl[si][1])

        if len(fl[si][1]) > 0: fromCustom = True
        # 如果有来自自定义文件夹标记，则将 fromCustom 设为 True
        
        # 检查字体扩展名
        ext = path.splitext(s)[1][1:]
        if ext.lower() not in ['ttf', 'ttc', 'otf', 'otc']: continue

        fontCRCKey = ''
        readFontCache = None
        dirCRCKey = ''
        isTc = False

        if not path.dirname(s).lower() in dirKeyList:
            dirCRCKey = hex(zlib.crc32(path.dirname(s).lower().encode('utf-8')))[2:].upper().rjust(8, '0')
            dirKeyList[path.dirname(s).lower()] = dirCRCKey
        else:
            dirCRCKey = dirKeyList[path.dirname(s).lower()]

        if usingCache:
            if not dirCRCKey in fontCache:
                fontCacheFile = path.join(fontCacheDir, '{}.json'.format(dirCRCKey))
                if path.exists(fontCacheFile):
                    try:
                        jsonFile = open(fontCacheFile, mode='r', encoding='utf-8')
                        fontCache[dirCRCKey] = json.load(jsonFile)
                        jsonFile.close()
                    except:
                        print('\n\033[1;31m[ERROR] 缓存文件读取失败 \"{}\"\033[0m'.format('{}.json'.format(dirCRCKey)))
                        os.remove(fontCacheFile)
                    else:
                        if not isinstance(list(fontCache[dirCRCKey].values())[0]['0'][3], dict):
                            print('\n\033[1;32m[FIX] 已删除旧版缓存 \"{}\"\033[0m'.format('{}.json'.format(dirCRCKey)))
                            os.remove(fontCacheFile)
                            del fontCache[dirCRCKey]
            
            if dirCRCKey in fontCache:
                fontCRCKey = hex(zlib.crc32((s.lower() + str(path.getsize(s)) + str(path.getctime(s))).encode('utf-8')))[2:].upper().rjust(8, '0')
                if fontCRCKey in fontCache[dirCRCKey]:
                    readFontCache = fontCache[dirCRCKey][fontCRCKey]
                    if len(readFontCache) > 1: isTc = True
                else:
                    readFontCache = None
                
        # 输出进度
        if readFontCache is None:
            print('\r' + '\033[1;32m{0}/{1} {2:.2f}%\033[0m \033[1m{3}\033[0m'.format(si + 1, flL, ((si + 1) / flL) * 100, outputSameLength(path.dirname(s))), end='', flush=True)
        else:
            print('\r' + '\033[1;32m{0}/{1} {2:.2f}%\033[0m \033[1m【缓存】{3}\033[0m'.format(si + 1, flL, ((si + 1) / flL) * 100, outputSameLength(path.dirname(s))), end='', flush=True)
        
        if readFontCache is None:
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
        else:
            tc = readFontCache
        
        # 计算字体CRC
        if len(fontCRCKey) == 0:
            fontCRCKey = hex(zlib.crc32((s.lower() + str(path.getsize(s)) + str(path.getctime(s))).encode('utf-8')))[2:].upper().rjust(8, '0')
        if not fontCacheN.get(dirCRCKey, False):
            fontCacheN[dirCRCKey] = {fontCRCKey: {}}
        else:
            fontCacheN[dirCRCKey].update({fontCRCKey: {}})

        # 处理字体信息
        for ti, t in enumerate(tc):

            tencodings = []

            if readFontCache is None:

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
                dictFstyle = {'other': []}
                familyN = ''
                dictFamilyN = {'other': []}
                try:
                    len(t['name'].names)
                except:
                    break

                isWarningFont = False
                dictNameStr = {'other': []}

                for ii, name in enumerate(t['name'].names):

                    platfID = name.platformID
                    langID = name.langID
                    encoding = lcidfil[platfID].get(langID, 'other')

                    c = ''
                    # 若 nameID 为 1，读取 NameRecord 的字体家族名称
                    if name.nameID == 1:
                        familyN, c, isWarningFont = fnReadCheck(t, ii, s)
                        familyN = familyN.strip(' ').strip('\u0000')

                        if encoding not in dictFamilyN or isinstance(dictFamilyN[encoding], str):
                            dictFamilyN[encoding] = familyN
                        else:
                            dictFamilyN['other'].append(familyN)
                        
                        font_family.setdefault(familyN, {})

                    # 若 nameID 为 2，读取 NameRecord 的字体样式
                    elif name.nameID == 2:
                        fstyle, c, isWarningFont = fnReadCheck(t, ii, s)
                        if not fstyle is None:
                            fstyle = fstyle.strip(' ').strip('\u0000')

                            if encoding not in dictFstyle or isinstance(dictFstyle[encoding], str):
                                dictFstyle[encoding] = fstyle
                            else:
                                dictFstyle['other'].append(fstyle)
            
                    # 若 nameID 为 4，读取 NameRecord 的字体完整名称           
                    elif name.nameID == 4:
                        # 如果 fontTools 解码失败，则尝试使用 utf-16-be 直接解码
                        namestr, c, isWarningFont = fnReadCheck(t, ii, s)
                        if namestr is None: continue
                        namestr = namestr.strip(' \u0000')

                        if encoding not in dictNameStr:
                            dictNameStr[encoding] = [namestr]
                        elif namestr not in dictNameStr[encoding]:
                            dictNameStr[encoding].append(namestr)
                    
                    else:
                        continue

            else:
                # 对于从字体缓存读取的字体，则直接赋值
                c = tc[t][6]
                isWarningFont = tc[t][5]
                dictNameStr = tc[t][0]
                isItalic = int(tc[t][1])
                isBold = int(tc[t][2])
                dictFstyle = tc[t][3]
                dictFamilyN = tc[t][4]
                for k, v in dictFamilyN.items():
                    if k == 'other':
                        for n in v:
                            font_family.setdefault(n, {})
                    else:
                        font_family.setdefault(v, {})
            
            if isWarningFont:
                if c not in tencodings:
                    tencodings.append(c)

            # 写入到字体缓存
            fontCacheN[dirCRCKey][fontCRCKey][ti] = (dictNameStr, isItalic, isBold, dictFstyle, dictFamilyN, isWarningFont, c)

            for ec, nsList in dictNameStr.items():
                for namestr in nsList:
                    # print(namestr)
                    if isWarningFont:
                        if namestr not in warning_font:
                            warning_font[namestr.lower()] = [c for c in tencodings if len(c) > 0]
                    
                    if f_n.get(namestr, False) and not overwrite:
                        # 如果发现列表中已有相同名称的字体，检测它的文件名、扩展名、父目录、样式是否相同
                        # 如果有一者不同且不来自自定义文件夹，添加到重复字体列表
                        dupp = f_n[namestr][0]
                        if (dupp != s and path.splitext(path.basename(dupp))[0] != path.splitext(path.basename(s))[0] and not fromCustom and f_n[namestr][2] == dictFstyle):
                            if not startQuiet: print('\n\033[1;35m[WARNING] 字体\"{0}\"与字体\"{1}\"的名称\"{2}\"重复！\033[0m'.format(
                                path.basename(f_n[namestr][0]), path.basename(s), namestr))
                            if dupfont.get(namestr) is not None:
                                if s not in dupfont[namestr]:
                                    dupfont[namestr].append(s)
                            else:
                                dupfont[namestr] = [dupp, s]
                    else:
                        f_n[namestr] = [s, ti, dictFstyle, isTc]
                    
                    if not font_n_lower.get(namestr.lower(), False):
                        font_n_lower[namestr.lower()] = namestr
                        
                if len(dictFamilyN) >= 1:
                    listFamilyN = []
                    if ec in dictFamilyN and isinstance(dictFamilyN[ec], str):
                        listFamilyN.append(dictFamilyN[ec])
                    else:
                        listFamilyN.extend(dictFamilyN['other'])
                    for fN in listFamilyN:
                        if font_family[fN].get((isItalic, isBold)) is not None:
                            if namestr not in font_family[fN][(isItalic, isBold)]:
                                font_family[fN][(isItalic, isBold)].append(namestr)
                        else:
                            font_family[fN].setdefault((isItalic, isBold), [namestr])
            
            if readFontCache is None: t.close()

            f_all_item = (dictNameStr, isItalic, isBold, dictFstyle, dictFamilyN)
            if s in f_all:
                f_all[s].setdefault(ti, f_all_item)
            else:
                f_all.setdefault(s, {ti : f_all_item})

    for dirK in fontCacheN.keys():
        fontCacheFile = path.join(fontCacheDir, '{0}.json'.format(dirK))
        jsonFile = open(fontCacheFile, mode='w', encoding='utf-8')
        json.dump(fontCacheN[dirK], fp=jsonFile, indent=2, ensure_ascii=False)
        jsonFile.close()

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

    del dupfont_cache, readFontCache, fontCache, fontCacheN

    f_all.clear()

    return [f_n, font_n_lower, font_family, warning_font, f_all]


# print(filelist)
# if path.exists(fontspath10): filelist = filelist.extend(os.listdir(fontspath10))

# for s in font_name.keys(): print('{0}: {1}'.format(s, font_name[s]))
def fnGetFromFamilyName(font_family: dict, fn: str, isitalic: int, isbold: int) -> str:
    if font_family.get(fn, False):
        if font_family[fn].get((isitalic, isbold), False):
            return font_family[fn][(isitalic, isbold)][0]
        elif font_family[fn].get((0, isbold), False):
            return font_family[fn][(0, isbold)][0]
        elif font_family[fn].get((isitalic, 0), False):
            return font_family[fn][(isitalic, 0)][0]
        elif font_family[fn].get((0, 0), False):
            return font_family[fn][(0, 0)][0]
        else:
            return fn
    else:
        return fn


def checkAssFont(fontlist: dict, font_info: list, fn_lines: list = [], onlycheck: bool = False, checkf: str = ''):
    """
系统字体完整性检查，检查是否有ASS所需的全部字体，如果没有，则要求拖入

需要以下输入
  fontlist: 字体与其所需字符（只读字体部分） { ASS内的字体名称 : 字符串 }
  font_name: 字体名称与字体路径对应词典 { 字体名称 : [ 字体绝对路径, 字体索引 ] }

以下输入可选
  fn_lines: fn标签所指定的字体与其包含的文本
  onlycheck: 缺少字体时不要求输入

将会返回以下
  assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称, 修正名称, [(斜体, 粗体)] ]}
  font_name: 同上，用于有新字体拖入时对该词典的更新
    """
    global noRequestFont

    # 从fontlist获取字体名称
    assfont = {}
    font_name = font_info[0]
    font_n_lower = font_info[1]
    font_family = font_info[2]
    warning_font = font_info[3]
    f_all = font_info[4]

    ignoreLostFonts = noRequestFont

    keys = list(fontlist.keys())
    for s in keys:
        isbold = s[2]
        isitalic = s[1]
        ns = s[0]
        ns = fnGetFromFamilyName(font_family, ns, isitalic, isbold)
        cok = False
        # 在全局字体名称词典中寻找字体名称
        ss = ns
        if ns not in font_name:
            # 如果找不到，将字体名称统一为小写再次查找
            if font_n_lower.get(ns.lower(), False):
                ss = font_n_lower[ns.lower()]
                cok = True
        elif not path.exists(font_name[ns][0]):
            del font_name[ns]
            cok = False
        else:
            cok = True
        directout = 0
        recheck = False
        if not cok:
            # 如果 onlycheck 不为 True，向用户要求目标字体
            if not onlycheck:
                if ignoreLostFonts: 
                    del fontlist[s]
                    continue
                print('\033[1;31m[ERROR] 缺少字体\"{0}\"\n请输入追加的字体文件或其所在字体目录的绝对路径\n或输入\"SKIPFONT\"（全大写）以跳过剩余的缺少字体\033[0m'.format(ns))
                inFont = ''
                while inFont == '' and directout < 3:
                    inFont = input().strip('\'\" ')

                    if inFont == 'SKIPFONT':
                        cok = False
                        ignoreLostFonts = True
                        del fontlist[s]
                        break

                    if path.exists(inFont):
                        if path.isdir(inFont):
                            addFont = fontProgress(getFontFileList([inFont], noreg=True)[0])
                            print()
                        else:
                            addFont = fontProgress([[inFont, '0']])
                            print()
                        if ns.lower() not in addFont[1].keys():
                            if path.isdir(inFont):
                                print('\033[1;31m[ERROR] 输入路径中\"{0}\"没有所需字体\"{1}\"\033[0m'.format(inFont, ns))
                            else:
                                print('\033[1;31m[ERROR] 输入字体\"{0}\"不是所需字体\"{1}\"\033[0m'.format(
                                    '|'.join(addFont[0].keys()), ns))
                            inFont = ''
                        else:
                            if f_priority:
                                font_name.update(addFont[0])
                                font_n_lower.update(addFont[1])
                                font_family.update(addFont[2])
                                warning_font.extend(addFont[3])
                                f_all.update(addFont[4])
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
                                        warning_font[fd] = addFont[3][fd]
                                for fd in addFont[4].keys():
                                    if fd not in f_all:
                                        f_all[fd] = addFont[4][fd]
                            cok = True
                            recheck = True
                    else:
                        print('\033[1;31m[ERROR] 您没有输入任何字符！再回车{0}次回到主菜单\033[0m'.format(3 - directout))
                        directout += 1
                        inFont = ''
            else:
                # 否则直接添加空
                assfont[(ns, ns)] = ['', s[0], ns, [checkf], [(isitalic, isbold)]]
        
        # 如果找不到，将字体名称统一为小写再次查找
        if recheck:
            if ns not in font_name:
                if font_n_lower.get(ns.lower()) is not None:
                    ss = font_n_lower[ns.lower()]
                    cok = True
            elif not path.exists(font_name[ns][0]):
                del font_name[ns]
                cok = False
            else:
                cok = True

        if cok:
            if not onlycheck and ss.lower() in warning_font:
                print('\033[1;31m[WARNING] 字体\"{0}\"可能不能正常子集化\033[0m'.format(ss))
                if warningStop:
                    print('\033[1;31m[WARNING] 请修复\"{0}\"，工作已中断\033[0m'.format(ss))
                    os.system('pause')
                    return None, [font_name, font_n_lower, font_family, warning_font, f_all]
            if directout < 3:
                # 如果找到，添加到assfont列表
                font_path = font_name[ss][0]
                font_index = font_name[ss][1]
                # print(font_name[ss])
                dict_key = (font_path, str(font_index))
                # 如果 assfont 列表已有该字体，则将新字符添加到 assfont 中
                if assfont.get(dict_key) is None:
                    assfont[dict_key] = [fontlist[s], s[0], ns, [checkf], [(isitalic, isbold)]]
                else:
                    tfname = assfont[dict_key][2]
                    newfnamep = assfont[dict_key][1].split('|')
                    oldstr = fontlist[s]
                    for newfname in newfnamep:
                        if s[0].lower() not in '|'.join(newfnamep).lower():
                            key1 = 0
                            key2 = 0
                            for ii in [0, 1]:
                                key1 = ii
                                key2 = 0
                                if fontlist.get((newfname, key1, 0)) is not None:
                                    break
                                elif fontlist.get((newfname, key1, 1)) is not None:
                                    key2 = 1
                                    break
                            newfstr = fontlist[(newfname, key1, key2)]
                            newfname = '|'.join([s[0], newfname])
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
                        assfont[dict_key] = [oldstr, newfname, tfname, list(set(assfont[dict_key][3] + [checkf])), list(set(assfont[dict_key][4] + [(isitalic, isbold)]))]
                        fontlist[s] = oldstr
                    # else:
                    #     assfont[dict_key] = [fontlist[s], s]
                # print(assfont[dict_key])
        if directout >= 3:
            return None, [font_name, font_n_lower, font_family, warning_font, f_all]
    if len(fn_lines) > 0 and not onlycheck:
        fn_lines_cache = fn_lines
        for i in range(0, len(fn_lines_cache)):
            s = fn_lines_cache[i]
            fi = s[1][s[1].keys()[0]]
            fn_lines[i] = [s[0], {s[1].keys()[0]: [fi[1], fnGetFromFamilyName(font_family, fi[1], fi[2], fi[3])]}]
    # print(assfont)
    return assfont, [font_name, font_n_lower, font_family, warning_font, f_all]


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


def charExistCheck(font_path: str, font_number: int, toCheck: str):
    '''
字符存在检查

通过字体内的CMAP，即文本编码(Textcoding)到字形(Glyph)映射表确认字体中是否有所需字形
分为 Unicode CMAP 和 ANSI CMAP 两部分

需要以下输入:
    font_path       字体绝对路径
    font_number     字体索引（针对TTC/OTC文件）
    toCheck         需要检查的字符串

将会返回以下:
    gfs             字符对应的Glyph-name（ANSI CMAP）
                    字符对应的Unicode（Unicode CMAP）
    gfs_uni         ANSI到Unicode对应词典（ANSI CMAP）
                    特殊键:
                        cmapFault CMAP读取失败
                        textInput 表示该字体为Unicode CMAP
    out_of_range    字体中缺少的字符
    '''
    gfs = ''
    gfs_uni = {}
    tf = ttLib.TTFont(font_path, fontNumber=int(font_number), lazy=True)
    cmap = {}
    toCheck = re.sub(r'[\u007F]', '', toCheck)
    out_of_range = ''
    # 有部分字体的Format12表格fontTools读不出来，所以先Try，如果不行就直接退出
    try:
        for c in tf['cmap'].tables:
            if c.isUnicode() and len(c.cmap) > len(cmap):
                cmap = c.cmap
    except:
        tf.close()
        return 'cmapFault', {'cmapFault': True}, ''
    # 如果第一步的Try中获取到了Unicode CMAP，则直接处理 Unicode CMAP
    # 处理非常简单: 将字符转换为 Unicode 编码对应的十进制数 (ord)，然后在词典里查找即可
    # 如果词典里没有，就是没有
    if len(cmap) > 0:
        for c in toCheck:
            dec = ord(c)
            if cmap.get(dec) is None:
                out_of_range = out_of_range + c
            else:
                addChar = hex(dec)[2:]
                if len(addChar) > 4:
                    addChar = 'u' + addChar.rjust(8, '0')
                elif len(addChar) > 2:
                    addChar = 'u' + addChar.rjust(4, '0')
                gfs = ','.join([gfs, addChar])
        tf.close()
        return gfs, {'textInput': True}, out_of_range
    else:
        # 如果没找到 Unicode CMAP，那就找 ANSI CMAP
        if len(tf['cmap'].tables) > 0:
            tcoding = ''
            for t in tf['cmap'].tables:
                if len(t.cmap) > len(cmap):
                    cmap = t.cmap
                    tcoding = t.getEncoding()
            # 如果这也找不到，那就退
            if len(cmap) == 0:
                tf.close()
                return '', {}, ''
            # fontTools默认返回GB2312，但是GB2312范围太小，如"·"是无法解码的，这里替换为GBK来解决这一问题
            if tcoding.lower() in ['gb2312', 'x_mac_simp_chinese_ttx']: tcoding = 'gbk'
            print('\n\033[1;33m[Warning] 正在使用地区编码\033[0m\"{0}\"\033[1;33m读取CMAP映射表\033[0m'.format(tcoding))
            for c in toCheck:
                dec = 0
                # 无论是哪个文本编码，0-127都是与ASCII相同的，所以直接ord就行了
                if re.match(r'[\u0000-\u007F]', c):
                    dec = ord(c)
                else:
                    # 如果是超出 ASCII 字符范围的，则将其用对应编码encode，然后将得到的 bytes 转换为十六进制，再转为十进制
                    try:
                        dec = int(''.join([hex(h).lstrip('0x') for h in c.encode(tcoding)]), 16)
                    except UnicodeError as e:
                        showUTF = 'U+' + ''.join([hex(h).lstrip('0x').upper() for h in c.encode('utf-16-be')])
                        reUTF = '[\\u{}]'.format(showUTF[2:])
                        raise UnicodeError(f'使用文本编码\"{e.encoding}\"处理字符\"{showUTF}\"时出错，请确认该字符是否正确。字符正则表达式：\"{reUTF}\"')
                if dec >= 0:
                    # 如果找到，比Unicode会多创建一个词典 gfs_uni，里面是 本地文本编码十进制 到 Unicode十进制 的对应关系
                    # 以便之后修复CMAP时使用
                    if cmap.get(dec) is None:
                        out_of_range = out_of_range + c
                    else:
                        gfs = ','.join([gfs, cmap.get(dec)])
                        gfs_uni[dec] = ord(c)
                else:
                    out_of_range = out_of_range + c
    gfs = gfs.lstrip(',')
    tf.close()
    return gfs, gfs_uni, out_of_range


def assFontSubset(assfont: dict, fontdir: str, allTTF: bool = False):
    """
字体子集化

需要以下输入:
  assfont: { 字体绝对路径?字体索引 : [ 字符串, ASS内的字体名称 ]}
  fontdir: 新字体存放目录

将会返回以下:
  newfont_name: { 原字体名 : [ 新字体绝对路径, 新字体名 ] }
    """
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
    else:
        fontdir = os.getcwd()
    print('\033[1;33m字体输出路径:\033[0m \033[1m\"{0}\"\033[0m'.format(fontdir))

    lk = len(assfont.keys())
    kip = 0
    for k in assfont.keys():
        kip += 1
        # 偷懒没有变更该函数中的assfont解析到新的词典格式
        # 在这里会将assfont词典转换为旧的assfont列表形式
        # assfont: [ 0字体绝对路径, 1字体索引, 2字符串, 3 ASS内的字体名称, 4修正字体名称, 5[(斜体, 粗体)] ]
        s = list(k) + [assfont[k][0], assfont[k][1], assfont[k][2], assfont[k][4]]
        fontext = path.splitext(path.basename(s[0]))[1]
        if allTTF:
            subfontext = '.ttf'
        elif fontext[1:].lower() in ['otc', 'ttc']:
            subfontext = fontext[:3].lower() + 'f'
        else:
            subfontext = fontext
        # print(fontdir, path.exists(path.dirname(fontdir)), path.exists(fontdir))
        fontname = re.sub(cillegal, '_', s[4])
        subfontpath = path.join(fontdir, fontname + subfontext)
        print('\r\033[1;32m[{0}/{1}]\033[0m \033[1m正在子集化…… \033[0m'.format(kip, lk), end='')
        if char_compatible:
            if re.search(r'[0-9]', s[2]):
                s[2] = '{0}{1}'.format(re.sub(r'[0-9]', '', s[2].replace('\n', '')), '0123456789')
            s[2] = '{0}{1}'.format(re.sub(r'[a-zA-Z]', '', s[2].replace('\n', '')), 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ')

        # 通过CMAP映射表确认字体中是否有所需字符
        gfs, gfs_uni, out_of_range = charExistCheck(s[0], s[1], s[2])
        gfs = gfs.lstrip(',')
        if len(out_of_range) > 0:
            if ignoreLost:
                print('\n\033[1;31m[WARNING] 已忽略不在字体中的字符\033[0m')
            else:
                print('\n\033[1;31m[ERROR] 以下字符不在字体\"{1}\"内\033[0m\n\"{0}\"\n\033[1;31m[ERROR] 以上字符不在字体\"{1}\"内\033[0m'.format(out_of_range, s[3]))
                print('\033[1;31m[ERROR] 已停止子集化，如果您想要强行子集化，请启用ignoreLost\033[0m')
                return None
        if len(gfs) == 0:
            print('\n\033[1;31m[WARNING] 没有要子集化的字符\033[0m')
            continue


        subsetarg = None
        if gfs_uni.get('cmapFault', False):
            print('\n\033[1;31m[WARNING] 字符缺失检查失败，该字体可能无法正常子集化\033[0m')
            subsetarg = [s[0], '--text={0}'.format(s[2]), '--output-file={0}'.format(subfontpath),
                     '--font-number={0}'.format(s[1]), '--passthrough-tables', '--glyph-names', '--recommended-glyphs', '--ignore-missing-unicodes']
        elif gfs_uni.get('textInput', False):
            subsetarg = [s[0], '--unicodes={0}'.format(gfs), '--output-file={0}'.format(subfontpath),
                     '--font-number={0}'.format(s[1]), '--passthrough-tables', '--glyph-names', '--recommended-glyphs', '--ignore-missing-unicodes']
        else:
            subsetarg = [s[0], '--glyphs={0}'.format(gfs), '--output-file={0}'.format(subfontpath), '--font-number={0}'.format(s[1]),
            '--passthrough-tables', '--name-legacy', '--legacy-cmap', '--glyph-names', '--recommended-glyphs', '--ignore-missing-glyphs']
        # else:
        #     print('\n\033[0;32m[CHECK] \"{0}\"已通过字符完整性检查\033[0m'.format(s[3]))

        try:
            subset.main(subsetarg)
        # except PermissionError:
        #     print('\n\033[1;31m[ERROR] 文件\"{0}\"访问失败\033[0m'.format(path.basename(subfontpath)))
        #     continue
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
            for si in s[5]:
                for si3 in s[3].split("|"):
                    newfont_name[si3, si[0], si[1]] = [crcnewf, subfontcrc]
            continue
        # os.system('pyftsubset {0}'.format(' '.join(subsetarg)))
        if path.exists(subfontpath):
            subfontcrc = hex(zlib.crc32((gfs + s[4]).encode('utf-8', 'replace')))[2:].upper()
            if len(subfontcrc) < 8: subfontcrc = '0' + subfontcrc
            # print('CRC32: {0} \"{1}\"'.format(subfontcrc, path.basename(s[0])))

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


            # 将ANSI编码的CMAP转换为Unicode编码以修复兼容性
            cmap4 = None
            if len(gfs_uni) > 0:
                cmap = newf['cmap'].tables[0].cmap
                use_table = newf['cmap'].tables[0]
                for t in newf['cmap'].tables:
                    if len(cmap) < len(t.cmap):
                        cmap = t.cmap
                        use_table = t
                cmap4 = use_table.newSubtable(4)
                cmap4.platformID = 0
                cmap4.platEncID = 3
                cmap4.language = 0
                for d in gfs_uni.keys():
                    if cmap.get(d) is not None:
                        cmap[gfs_uni[d]] = cmap[d]
                        if not d == gfs_uni[d]: cmap.pop(d)
                cmap4.cmap = cmap
                newf['cmap'].tables = [cmap4]
                cmap4 = use_table.newSubtable(4)
                cmap4.platformID = 1
                cmap4.platEncID = 0
                cmap4.language = 0
                for d in gfs_uni.keys():
                    if cmap.get(d) is not None:
                        if d in range(0, 128):
                            cmap[gfs_uni[d]] = cmap[d]
                            if not d == gfs_uni[d]: cmap.pop(d)
                        else:
                            cmap.pop(d)
                cmap4.cmap = cmap
                newf['cmap'].tables.append(cmap4)
                cmap4 = use_table.newSubtable(4)
                cmap4.platformID = 3
                cmap4.platEncID = 1
                cmap4.language = 0
                for d in gfs_uni.keys():
                    if cmap.get(d) is not None:
                        cmap[gfs_uni[d]] = cmap[d]
                        if not d == gfs_uni[d]: cmap.pop(d)
                cmap4.cmap = cmap
                newf['cmap'].tables.append(cmap4)
                newf['cmap'].tables = newf['cmap'].tables[:3]


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
                newf.save(crcnewf, False)
                newf.close()
            rawf.close()
            if path.exists(crcnewf):
                if not subfontpath == crcnewf: os.remove(subfontpath)
                for si in s[5]:
                    for si3 in s[3].split("|"):
                        newfont_name[si3, si[0], si[1]] = [crcnewf, subfontcrc]
    print('')
    # print(newfont_name)
    return newfont_name


def dExistsPath(filePath: str, isFile: bool = True):

    def isExists(fPath: str):
        return (isFile and path.isfile(fPath)) or (not isFile and path.isdir(fPath))

    testpath = filePath

    if isExists(filePath):
        testpathl = path.splitext(filePath)
        testc = 1
        testpath = '{0}#{1}{2}'.format(testpathl[0], testc, testpathl[1])
        while isExists(testpath):
            testc += 1
            testpath = '{0}#{1}{2}'.format(testpathl[0], testc, testpathl[1])
    
    return testpath


def assFontChange(newfont_name: dict, asspath: str, assInfo: dict, splitEvents: dict[int, list[dict]], outdir: str = '', 
                  cover: bool = False, embedded: bool = False) -> str:
    """
更改ASS样式对应的字体

需要以下输入
  fullass: 完整的ass文件内容，以行分割为列表
  newfont_name: { 修正字体名 : [ 新字体路径, 新字体名 ] }
  asspath: 原ass文件的绝对路径
  assInfo
  splitEvents

以下输入可选
  outdir: 新字幕的输出目录，默认为源文件目录
  cover: 为True时覆盖原有文件，为False时不覆盖
  fn_lines: 带有fn标签的行数，对应到fullass的索引
  embedded: 是否使用ASS/SSA内嵌字体
  fontline: 字体标签所在行

将会返回以下
  newasspath: 新ass文件的绝对路径
    """

    # 扫描Style各行，并替换掉字体名称
    if len(newfont_name) > 0: 
        for k in assInfo['Styles'].keys():
            styleItalic = abs(int(assInfo['Styles'][k]['Italic']))
            styleBold = abs(int(assInfo['Styles'][k]['Bold']))
            fontstr = assInfo['Styles'][k]['Fontname']
            nfnKey = (fontstr.lstrip('@'), styleItalic, styleBold)
            if newfont_name.get(nfnKey) is None:
                continue
            else:
                if not newfont_name[nfnKey][1] is None:
                    assInfo['Styles'][k]['Fontname'] = newfont_name[nfnKey][1].split('?')[0]
                    if fontstr[0] == '@': assInfo['Styles'][k]['Fontname'] = '@' + assInfo['Styles'][k]['Fontname']
                    
        used_nf_name2 = {}
        for k2 in [k for k in newfont_name.keys() if not newfont_name[k][1] is None]:
            used_nf_name2[k2] = newfont_name[k2]
        # used_nf_name = {list(used_nf_name2.keys())[0].upper(): used_nf_name2[list(used_nf_name2.keys())[0]]}
        used_nf_name = {}
        for k in used_nf_name2.keys():
            used_nf_name[(k[0].upper(), k[1], k[2])] = used_nf_name2[k]
        
        if len(splitEvents) > 0:

            # print('正在处理fn标签......')
            # fn_lines: 带有fn标签的行数与该行的完整特效标签，一项一个 [ [行数, { 标签1 : ( ASS内部字体名称, Italic, Bold )}], ... ]
          
            for fnLine in splitEvents.keys():
                if len(splitEvents[fnLine]) == 1: continue
                for l in splitEvents[fnLine]:
                    fname = l['Fontname'].lstrip('@')
                    #if len(fname) > 0 and not used_nf_name[fname.upper()][1] is None:
                    if len(fname) > 0:
                        replaceFn = None
                        flKey = (fname.upper(), abs(l['Italic']), abs(l['Bold']))

                        if used_nf_name.get(flKey, False):
                            replaceFn = used_nf_name[flKey]
                        else:
                            print('找不到字体 {0}\n可能是该标签指定字体未被使用，或出现了BUG，请注意检查输出'
                                    .format(fname))
                            os.system('timeout /T 10')
                            continue

                        if replaceFn is not None:
                            if replaceFn[1] is not None:
                                fnTags = re.findall(r'\{.*?\\fn.*?\}', l['Text'])
                                for t in fnTags:
                                    l['Text'] = l['Text'].replace(t, t.replace(l['Fontname'], replaceFn[1])) 
                                l['Fontname'] = replaceFn[1]

                assInfo['Events'][fnLine]['Text'] = ''.join([v['Text'] for v in splitEvents[fnLine]])

        for k in used_nf_name2:
            assInfo['Subset'][newfont_name[k][1]] = k[0].split('|')[0]

        if embedded:

            for k in newfont_name.keys():

                f = newfont_name[k][0]
                if path.basename(f) in assInfo['Fonts']: continue

                emFont = open(f, mode='rb')
                assInfo['Subset'][newfont_name[k][1]] = k[0]
                assInfo['Fonts'][path.basename(f)] = emFont.read()
                emFont.close()
                try:
                    os.remove(f)
                except:
                    pass

                del emFont
            
            print(path.dirname(list(newfont_name.values())[0][0]))
            try:
                os.rmdir(path.dirname(list(newfont_name.values())[0][0]))
            except:
                pass

        del used_nf_name, used_nf_name2

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
    if path.exists(newasspath) and not cover:
        newasspath = dExistsPath(newasspath)

    assFileWrite(newasspath, assInfo, False, False, ignoreExists=cover)

    if embedded:
        print('\033[1;32m字幕嵌入成功\033[0m')
    
    return newasspath


def ffASFMKV(file: str, outfile: str = '', asslangs: list = [], asspaths: list = [], fontpaths: list = []) -> int:
    """
ffASFMKV，将媒体文件、字幕、字体封装到一个MKV文件，需要ffmpeg、ffprobe命令行支持

需要以下输入
    file: 媒体文件绝对路径
    outfile: 输出文件的绝对路径，如果该选项空缺，默认为 输入媒体文件.muxed.mkv
    asslangs: 赋值给字幕轨道的语言，如果字幕轨道多于asslangs的项目数，超出部分将全部应用asslangs的末项
    asspaths: 字幕绝对路径列表
    fontpaths: 字体列表，格式为 [字体1绝对路径, 字体2绝对路径, ...]
将会返回以下
    ffmr: ffmpeg命令行的返回值    
    """
    # print(fontpaths)
    global rmAssIn, rmAttach, mkvout, notfont
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

    rawIdx = json.loads(os.popen('ffprobe -print_format json -show_streams -hide_banner -v 0 -i \"{}\"'.format(file), mode='r').read())
    firstIdx = len(rawIdx['streams'])
    ffargs = []
    metaList = []
    mapList = []
    copyList = []
    
    for f in rawIdx['streams']:
        if f['codec_type'] == 'video':
            if '-map 0:v' not in mapList:
                mapList.append('-map 0:v')
                copyList.append('-c:v copy')
        elif f['codec_type'] == 'audio':
            if '-map 0:a' not in mapList:
                mapList.append('-map 0:a')
                copyList.append('-c:a copy')
        elif f['codec_type'] == 'attachment':
            if not rmAttach:
                if '-map 0:t' not in mapList:
                    mapList.append('-map 0:t')
                    copyList.append('-c:t copy')
            else:
                firstIdx -= 1
        elif f['codec_type'] == 'subtitle':
            if rmAssIn: 
                firstIdx -= 1
            elif '-map 0:s' not in mapList:
                mapList.append('-map 0:s')
    ffargs.append('-i \"{}\"'.format(file))
  
    fn = path.splitext(path.basename(file))[0]
    metaList.append('-metadata:g title=\"{}\"'.format(fn))
    if len(asspaths) > 0:
        for i in range(0, len(asspaths)):
            s = asspaths[i]
            assfn = path.splitext(path.basename(s))[0]
            assnote = assfn[(assfn.find(fn) + len(fn)):].replace('.subset', '')
            # print(assfn, fn, assnote)
            metadata = None
            metadata = []
            if len(assnote) > 1:
                metadata.append('title=\"{}\"'.format(assnote.lstrip('.')))
            if len(asslangs) > 0 and path.splitext(s)[1][1:].lower() not in ['idx']:
                if i < len(asslangs):
                    metadata.append('language=\"{}\"'.format(asslangs[i]))
                else:
                    metadata.append('language=\"{}\"'.format(asslangs[len(asslangs) - 1]))
            if len(metadata) > 0:
                for m in metadata:
                    metaList.append('-metadata:s:{1} {0}'.format(m, firstIdx))
            ffargs.append('-i \"{}\"'.format(s))
            mapList.append('-map {}'.format(i + 1))
            firstIdx += 1
    if len(fontpaths) > 0:
        extMime = {
            'ttf': 'font/ttf',
            'otf': 'font/otf'
            }
        for s in fontpaths:
            metaList.append('-metadata:s:{2} title=\"{0}\" -metadata:s:{2} mimetype=\"{1}\"'.format(path.splitext(path.basename(s))[0].split('.')[-1], 
                                                                              extMime[path.splitext(s)[1][1:].lower()], firstIdx))
            ffargs.append('-attach \"{}\"'.format(s))
            firstIdx += 1
    ffmr = os.system('ffmpeg {0} {3} {4} -c:s copy {2} \"{1}\"'.format(' '.join(ffargs), outfile, ' '.join(metaList), ' '.join(mapList), ' '.join(copyList)))
    # print('ffmpeg {0} {3} {4} -c:s copy {2} \"{1}\"'.format(' '.join(ffargs), outfile, ' '.join(metaList), ' '.join(mapList), ' '.join(copyList)))
    if ffmr >= 1:
        print('\n\033[1;31m[ERROR] 检测到不正常的mkvmerge返回值，重定向输出...\033[0m')
        os.system('set \"FFREPORT=file=ffreport_cache.log:level=32\" & chcp 65001 & ffmpeg {0} {2} {3} -c:s copy {1} NUL'.format(' '.join(ffargs), ' '.join(metaList), ' '.join(mapList), ' '.join(copyList)))
        try:
            shutil.move('ffreport_cache.log', '{0}.{1}.log'.format(path.splitext(file)[0], datetime.now().strftime('%Y-%m%d-%H%M-%S_%f')))
        except:
            print('\033[1;33m[ERROR] \"ffreport_cache.log\"移动失败\033[0m')
    elif not notfont:
        for p in asspaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(p))
            if path.splitext(p)[1][1:].lower() in ['ass', 'ssa']:
                try:
                    os.remove(p)
                except:
                    print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(p))
        for f in fontpaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(f))
            try:
                os.remove(f[0])
            except:
                print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(f))
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    else:
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    os.system('pause')
    return ffmr


def ASFMKV(file: str, outfile: str = '', asslangs: list = [], asspaths: list = [], fontpaths: list = []) -> int:
    """
ASFMKV，将媒体文件、字幕、字体封装到一个MKV文件，需要mkvmerge命令行支持

需要以下输入

  file: 媒体文件绝对路径
  outfile: 输出文件的绝对路径，如果该选项空缺，默认为 输入媒体文件.muxed.mkv
  asslangs: 赋值给字幕轨道的语言，如果字幕轨道多于asslangs的项目数，超出部分将全部应用asslangs的末项
  asspaths: 字幕绝对路径列表
  fontpaths: 字体列表，格式为 [字体1绝对路径, 字体2绝对路径, ...]

将会返回以下
  mkvmr: mkvmerge命令行的返回值    
    """
    # print(fontpaths)
    global rmAssIn, rmAttach, mkvout, notfont
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
            mkvargs.extend(['--attachment-name', path.basename(s)])
            mkvargs.extend(['--attach-file', s])
    mkvargs.extend(['--title', fn])
    mkvjsonp = path.splitext(file)[0] + '.mkvmerge.json'
    mkvjson = open(mkvjsonp, mode='w', encoding='utf-8')
    json.dump(mkvargs, fp=mkvjson, sort_keys=True, indent=2, separators=(',', ': '))
    mkvjson.close()
    mkvmr = os.system('mkvmerge @\"{0}\" -o \"{1}\"'.format(mkvjsonp, outfile))
    if mkvmr > 1:
        print('\n\033[1;31m[ERROR] 检测到不正常的mkvmerge返回值，重定向输出...\033[0m')
        os.system('mkvmerge -r \"{0}\" @\"{1}\" -o NUL'.format('{0}.{1}.log'
                                                               .format(path.splitext(file)[0],
                                                                       datetime.now().strftime('%Y-%m%d-%H%M-%S_%f')),
                                                               mkvjsonp))
    elif not notfont:
        for p in asspaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(p))
            if path.splitext(p)[1][1:].lower() in ['ass', 'ssa']:
                try:
                    os.remove(p)
                except:
                    print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(p))
        for f in fontpaths:
            print('\033[1;32m封装成功: \033[1;37m\"{0}\"\033[0m'.format(f))
            try:
                os.remove(f[0])
            except:
                print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(f))
        try:
            os.remove(mkvjsonp)
        except:
            print('\033[1;33m[ERROR] 文件\"{0}\"删除失败\033[0m'.format(mkvjsonp))
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    else:
        print('\033[1;32m输出成功:\033[0m \033[1m\"{0}\"\033[0m'.format(outfile))
    return mkvmr


def getFileList(inPath: str, extFilter: list[str] = [], subdir: bool = False) -> list[str]:
    """
从输入的目录中获取文件列表
需要以下输入
  inPath: 要搜索的目录
  extFilter: 扩展名过滤器
  subdir: 是否搜索子目录
返回以下结果
  fileList: 文件列表
      结构: [[ 文件名（无扩展名）, 绝对路径 ], ...]
    """
    fileList = []
    extFilter = [s.lower().strip('. ') for s in extFilter if len(s.strip(' ')) != 0]

    def checkExt(ext: str) -> bool:
        if len(extFilter) > 0:
            if ext.lower().lstrip('.') in extFilter:
                return True
        return False

    if path.isdir(inPath):
        if subdir:
            for r, ds, fs in os.walk(inPath):
                for f in fs:
                    if checkExt(path.splitext(f)[1]):
                        fileList.append((path.splitext(f)[0], path.join(r, f)))
        else:
            for f in os.listdir(inPath):
                cPath = path.join(inPath, f)
                if path.isfile(cPath) and checkExt(path.splitext(f)[1]):
                    fileList.append((path.splitext(f)[0], path.join(cPath)))
    else:
        fileList = [(path.splitext(inPath)[0], inPath)]

    return fileList


def getSubtitles(cpath: str, medias: list = []) -> dict:
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
                        vdir = path.dirname(li[1]).lower()
                        sdir = path.dirname(f).lower()
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
                        if path.dirname(li[1]).lower() == path.dirname(f).lower():
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

将会返回以下
  media_ass: 媒体文件与字幕文件对应词典
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
            elif m[0][mStart:mEndI].lower().rjust(sEndI-sStart, '0') == sName[sStart:sEndI].lower():
                media_ass[m].append(s)
            elif m[0][mStart:mEndI].lower() == sName[sStart:sEndI].lower().rjust(mEndI-mStart, '0'):
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
    if len(medias) == 0:
        print('找不到媒体文件')
        return None
    mStart, mEnd = namePosition([s[0] for s in medias])
    subs = getSubtitles(spath)['none']
    if len(subs) == 0:
        print('找不到字幕文件')
        return None
    sStart, sEnd = namePosition([path.splitext(path.basename(s))[0] for s in subs])
    sStartA, sEndA = namePosition([path.splitext(path.basename(s))[0][::-1] for s in subs])
    # 生成第一次的 media_ass 和 renList，并询问
    media_ass = nameMatchingProgress(medias, mStart, mEnd, sStart, sEnd, subs)
    renList, isSkip = renamePreview(media_ass, sStartA, sEndA)
    annl = []
    ann = ''
    if len(renList) > 0:
        print('''请确认重命名是否正确
如果注释部分不正确您可以选择\033[33;1m[A]\033[0m手动输入字幕注释
如果文件匹配不正确您可以选择\033[33;1m[F]\033[0m手动输入匹配规则''')
        rt = os.system('choice /M 请选择 /C YNAF')
    elif not isSkip:
        cls()
        print('视频-字幕匹配失败，请手动输入匹配规则\n')
        rt = 4
    else:
        rt = 2
    while rt in [0, 3, 4]:
        if rt == 4:
            if len(renList) > 0: cls()
            print('''请将剧集所在位置用\"\033[33;1m:\033[0m\"代替
如 \"\033[30;1m[VCB-Studio] Accel World [\033[33;1m01\033[30;1m][720p][x265_aac].mp4\033[0m\"
替换为 \"\033[30;1m[VCB-Studio] Accel World [\033[33;1m:\033[30;1m][720p][x265_aac].mp4\033[0m\"
请务必保留剧集位置后的一个字符以以便程序正常工作''')
            vSearch = input('请输入视频匹配规则:').strip(' \"\'')
            if len(vSearch) == 0:
                print('输入为空，退出')
                break
            sSearch = input('请输入字幕匹配规则:').strip(' \"\'')
            if len(sSearch) == 0:
                print('输入为空，退出')
                break
            vSplit = vSearch.split(':')
            if len(vSplit) != 2:
                print('输入不正确，退出')
                break
            sSplit = sSearch.split(':')
            if len(sSplit) != 2:
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
                annl = [re.sub(cillegal, '_', ann)]
        else:
            break
        renList, isSkip = renamePreview(media_ass, annl=annl)
        if len(renList) > 0:
            print('''请确认重命名是否正确
如果注释部分不正确您可以选择\033[33;1m[A]\033[0m手动输入字幕注释
如果文件匹配不正确您可以选择\033[33;1m[F]\033[0m手动输入匹配规则''')
            rt = os.system('choice /M 请选择 /C YNAF')
        elif not isSkip:
            cls()
            print('视频-字幕匹配失败，请手动输入匹配规则\n')
            rt = 4
        else:
            rt = 2
    else:
        if rt == 1:
            logfile = open(path.join(spath, 'subsRename_{0}.log'.format(datetime.now().strftime('%Y%m%d-%H%M%S'))),
                           mode='w', encoding='utf-8')
            for r in renList:
                try:
                    os.rename(r[0], path.join(path.dirname(r[0]), r[1]))
                except:
                    print('\033[31;1m[ERROR] \"{0}\"重命名失败\033[0m'.format(r[0]))
                else:
                    print('\033[32;1m[SUCCESS]\033[33;0m \"{0}\"\n          \033[0m>>> \033[1m\"{1}\"\033[0m'.format(
                        path.basename(r[0]), r[1]))
                    try:
                        logfile.write('{0}|{1}\n'.format(r[0], r[1]))
                    except:
                        print('\033[31;1m[ERROR] 日志文件写入失败\033[0m')
            logfile.close()
            del logfile
    del rt, mStart, mEnd, sStart, sEnd, sStartA, sEndA, media_ass, renList, annl, ann


def renamePreview(media_ass: dict, sStartA: int = 0, sEndA: str = '', ann: str = '', annl: list = []) -> list:
    """
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
    """
    global matchStrict
    cls()
    print('[重命名预览]\n')
    renList = []
    isSkip = False
    for k in media_ass.keys():
        sl = media_ass[k]
        annlUp = 0
        for s in sl:
            bname = path.basename(s)
            if matchStrict:
                if len(bname) >= len(k[0]):
                    if k[0] == bname[:len(k[0])]:
                        print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                        isSkip = True
                        continue
            else:
                if k[0] in bname:
                    print('\033[33;1m跳过已匹配字幕\033[0m \033[1m\"{0}\"\033[0m'.format(bname))
                    isSkip = True
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
        print('字幕的注释是指类似以下文件名的高亮部分\n\"\033[30;1m[VCB-Studio] Accel World [01][720p][x265_aac]'
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


def main(font_info: list, asspath: list, outdir: list = ['', '', ''], mux: bool = False, vpath: str = '',
         asslangs: list = [], FFmuxer: int = 0, fontline: int = -1):
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
  FFmuxer: 0 = mkvmerge, 1 = FFmpeg, 2 = ASS/SSA Embedded

将会返回以下
  newasspath: 列表，新生成的字幕文件绝对路径
  newfont_name: 词典，{ 原字体名 : [ 新字体绝对路径, 新字体名 ] }
  ??? : 数值，mkvmerge的返回值；如果 mux = False，返回-1
    """
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
                else:
                    outdir[i] = path.join(os.getcwd(), s)
    # print(outdir)
    # os.system('pause')
    global notfont
    # multiAss 多ASS文件输入记录词典
    # 结构: { ASS文件绝对路径 : [ 完整ASS文件内容(fullass), 样式位置(styleline), 字体在样式行中的位置(font_pos) ]}
    multiAss = {}
    assfont = {}
    newasspath = []
    fo = ''
    if not notfont:
        # print('\n字体名称总数: {0}'.format(len(font_name.keys())))
        # noass = False
        for i in range(0, len(asspath)):
            s = asspath[i]
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                multiAss[s] = (None, None)
            else:
                # print('正在分析字幕文件: \"{0}\"'.format(path.basename(s)))
                assInfo = assAnalyseV2(s)
                tagSplit = eventTagsSplit(assInfo, False)
                fontlist = assFontList(assInfo, tagSplit)
                if fontlist is None:
                    print('\033[1;31m[ERROR] 缺少参数的ASS/SSA文件！\"{0}\"\033[0m'.format(path.basename(s)))
                    fontlist = []
                    continue
                multiAss[s] = (assInfo, tagSplit)
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
        newfont_name = None
        newfont_name = assFontSubset(assfont, fo, FFmuxer==2)
        if newfont_name is None: return [], {}, -2
        for s in asspath:
            if path.splitext(s)[1][1:].lower() not in ['ass', 'ssa']:
                newasspath.append(s)
            elif len(multiAss[s][0]['Events']) == 0:
                continue
            else:
                newasspath.append(
                    assFontChange(newfont_name, s, multiAss[s][0], multiAss[s][1], outdir[0], embedded=FFmuxer==2))
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
        if FFmuxer == 1:
            mkvr = ffASFMKV(vpath, path.join(outdir[2], path.splitext(path.basename(vpath))[0] + '.mkv'),
                      asslangs=asslangs, asspaths=newasspath, fontpaths=list(set([f[0] for f in newfont_name.values()])))
        elif FFmuxer == 0:
            mkvr = ASFMKV(vpath, path.join(outdir[2], path.splitext(path.basename(vpath))[0] + '.mkv'),
                      asslangs=asslangs, asspaths=newasspath, fontpaths=list(set([f[0] for f in newfont_name.values()])))
        else:
            mkvr = 0
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
    global s_fontload
    font_list = []
    print('\033[1;32m正在尝试从字幕所在文件夹载入字体...\033[0m')
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
    return font_info


def cls():
    os.system('cls')


def cListAssFont(font_info):
    global resultw, s_subdir, copyfont, fontload, exact_lost, char_lost, embeddedFontExtract
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
[5] 显示缺少字体的字幕: \033[1;33m{4}\033[0m
[6] 显示字体缺字数(慢): \033[1;33m{5}\033[0m
[7] 抽出内嵌字体(如果有): \033[1;33m{6}\033[0m
'''.format(resultw, copyfont, s_subdir, fontload, exact_lost, char_lost, embeddedFontExtract))
        work = os.system('choice /M 请输入 /C AB1234567L')
        if work == 1:
            cls()
            wfilep = path.join(os.getcwd(), datetime.now().strftime('ASFMKV_FullFont_%Y-%m%d-%H%M-%S_%f.log'))
            if resultw:
                wfile = open(wfilep, mode='w', encoding='utf-8-sig')
            else:
                wfile = None
            fn = ''
            print('FontList', file=wfile)
            for s in font_name.keys():
                nfn = path.basename(font_name[s][0])
                if fn != nfn:
                    if wfile is not None:
                        print('>\n{0} <{1}'.format(nfn, s), end='', file=wfile)
                    else:
                        print('>\033[0m\n\033[1;36m{0}\033[0m \033[1m<{1}'.format(nfn, s), end='')
                else:
                    print(', {0}'.format(s), end='', file=wfile)
                fn = nfn
            if wfile is not None:
                print(wfilep)
                print('>', file=wfile)
                wfile.close()
            else:
                print('>\033[0m')
        elif work == 2:
            cls()
            cpath = ''
            directout = False
            while not path.exists(cpath) and not directout:
                directout = True
                cpath = input('请输入目录路径或字幕文件路径: ').strip('\" ')
                if cpath == '':
                    print('没有输入，回到上级菜单')
                elif not path.exists(cpath):
                    print('\033[1;31m[ERROR] 找不到路径: \"{0}\"\033[0m'.format(cpath))
                elif not path.isfile(cpath) and not path.isdir(cpath):
                    print('\033[1;31m[ERROR] 输入的必须是文件或目录！: \"{0}\"\033[0m'.format(cpath))
                elif not path.isabs(cpath):
                    print('\033[1;31m[ERROR] 路径必须为绝对路径！: \"{0}\"\033[0m'.format(cpath))
                elif path.isdir(cpath):
                    directout = False
                elif not path.splitext(cpath)[1][1:].lower() in ['ass', 'ssa']:
                    print('\033[1;31m[ERROR] 输入的必须是ASS/SSA字幕文件！: \"{0}\"\033[0m'.format(cpath))
                else:
                    directout = False
            # clist = []
            if not directout:

                def writeNonSubset(filePath: str, assInfo: dict):
                    if len(assInfo['Subset']) > 0:
                        nsdir = path.join(path.dirname(filePath), 'NoSubsetSub')
                        nsfile = path.join(nsdir, path.basename(filePath))
                        try:
                            if not path.exists(nsdir): os.mkdir(nsdir)
                            assFileWrite(nsfile, assInfo)
                            print('\033[1;33m已恢复字幕: \033[0m\033[1m\"{0}\"\033[0m'.format(nsfile))
                        except:
                            print('\033[1;31m[ERROR] 恢复的子集化字幕写入失败\n{0}\033[0m'.format(sys.exc_info()))

                def writeEmbeddedFonts(filePath: str, assInfo: dict):
                    cacheDir = path.join(path.dirname(filePath), 'ExtractFonts', path.basename(filePath))
                    if not path.isdir(path.dirname(cacheDir)):
                        os.mkdir(path.dirname(cacheDir))
                    if not path.isdir(cacheDir):
                        os.mkdir(cacheDir)
                    for f in assInfo['Fonts'].keys():
                        fontFile = open(path.join(cacheDir, f), mode='wb')
                        fontFile.write(assInfo['Fonts'][f])

                if path.isdir(cpath):
                    # print(cpath)
                    if fontload:
                        font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                    else:
                        font_info2 = font_info
                    for f in getFileList(cpath, ['ass', 'ssa'], s_subdir):
                        assInfo = assAnalyseV2(f[1], embeddedFontExtract)
                        writeNonSubset(f[1], assInfo)
                        fontlist = assFontList(assInfo, eventTagsSplit(assInfo, False))

                        # 记得写字体抽出
                        if embeddedFontExtract:
                            writeEmbeddedFonts(cpath, assInfo)

                        if len(assInfo['Styles']) == 0:
                            print('\033[1;31m[ERROR] 缺少参数的ASS/SSA文件！\"{0}\"\033[0m'.format(f[0]))
                            continue
                        assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True, checkf=f[0])
                        assfont = updateAssFont(assfont, assfont2)
                        del assfont2
                                    
                    fd = path.join(cpath, 'Fonts')
                else:
                    if fontload:
                        font_info2 = templeFontLoad(path.dirname(cpath), copy.deepcopy(font_info))
                    else:
                        font_info2 = font_info
                    assInfo = assAnalyseV2(cpath, embeddedFontExtract)
                    writeNonSubset(cpath, assInfo)
                    fontlist = assFontList(assInfo, eventTagsSplit(assInfo, False))

                    # 记得写字体抽出
                    if embeddedFontExtract:
                        writeEmbeddedFonts(cpath, assInfo)

                    if len(assInfo['Styles']) == 0:
                        print('\033[1;31m[ERROR] 缺少参数的ASS/SSA文件！\"{0}\"\033[0m'.format(cpath))
                        continue
                    else:
                        assfont2, font_info2 = checkAssFont(fontlist, font_info2, onlycheck=True, checkf=cpath)
                        assfont = updateAssFont(assfont, assfont2)
                        del assfont2
                        fd = path.join(path.dirname(cpath), 'Fonts')
            
                if len(assfont.keys()) < 1:
                    print('\033[1;31m[ERROR] 目标路径没有ASS/SSA字幕文件\033[0m')
                elif not fontlist is None:
                    wfile = None
                    print('')
                    if copyfont or resultw:
                        if not path.isdir(fd): os.mkdir(fd)
                    if resultw:
                        wfile = open(path.join(cpath, 'Fonts', 'fonts.txt'), mode='w', encoding='utf-8-sig')
                    maxnum = 0
                    out_of_ranges = {}
                    if char_lost:
                        print('\033[1;33m检查字体缺字情况，请稍等...\033[0m')
                    for s in assfont.keys():
                        
                        if not s[0] == s[1]:
                            lx = len(assfont[s][0])
                            if lx > maxnum:
                                maxnum = lx
                            if char_lost:
                                print('\033[1;33m正在检查:\033[0m \033[1m{0}\033[0m'.format(assfont[s][1]))
                                a, b, out_of_range = charExistCheck(s[0], int(s[1]), assfont[s][0])
                                if len(out_of_range) > 0:
                                    out_of_ranges[s] = [str(len(out_of_range)), out_of_range.strip(' ')]
                                elif len(a) == 0:
                                    print('\033[1;31m检查失败:\033[0m \033[1m{0}\033[0m'.format(assfont[s][1]))
                                    out_of_ranges[s] = ['N', '']
                                del a, b, out_of_range
                    print('【详情】')
                    maxnum = len(str(maxnum))
                    for s in assfont.keys():
                        if not s[0] == s[1]:
                            fp = s[0]
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
                                if not out_of_ranges.get(s) is None:
                                    if out_of_ranges[s][0] != 'N':
                                        print('{0}{2} chars no found: {1}'.format(''.center(maxnum, ' '), out_of_ranges[s][1], out_of_ranges[s][0]), file=wfile)
                            if errshow:
                                if char_lost:
                                    print('\033[1;32m[{3}\033[0m-\033[1;31m{4}\033[1;32m]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;31m{2}\033[0m'.format(
                                        assfont[s][2], path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum), out_of_ranges.get(s, ['0'])[0].ljust(maxnum)))
                                    if not out_of_ranges.get(s) is None:
                                        if out_of_ranges[s][0] != 'N':
                                            print('{0}[ \033[1;31m{1}\033[0m ]'.format(''.center(maxnum*2+3, ' '), out_of_ranges[s][1]))
                                else:
                                    print('\033[1;32m[{3}]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;31m{2}\033[0m'.format(
                                        assfont[s][2], path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum)))
                            else:
                                if char_lost:
                                    print('\033[1;32m[{3}\033[0m-\033[1;31m{4}\033[1;32m]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;32m{2}\033[0m'
                                    .format(assfont[s][2], path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum), out_of_ranges.get(s, ['0'])[0].ljust(maxnum)))
                                    if not out_of_ranges.get(s) is None:
                                        if out_of_ranges[s][0] != 'N':
                                            print('{0}[ \033[1;31m{1}\033[0m ]'.format(''.center(maxnum*2+3, ' '), out_of_ranges[s][1]))
                                else:
                                    print('\033[1;32m[{3}]\033[0m \033[1;36m{0}\033[0m \033[1m<{1}>\033[1;32m{2}\033[0m'
                                    .format(assfont[s][2], path.basename(fn), ann, str(len(assfont[s][0])).rjust(maxnum), out_of_ranges.get(s, ['0'])[0].ljust(maxnum)))
                            # print(assfont[s][0])
                        else:
                            if resultw:
                                print('{0} - No Found'.format(s[0]), file=wfile)
                                if exact_lost:
                                    for f in assfont[s][3]:
                                        print('{0}\"{1}\"'.format(''.center(4, ' '), f), file=wfile)
                            if char_lost: 
                                print('\033[1;31m[{1}]\033[1;36m {0}\033[1;31m - No Found\033[0m'
                                .format(s[0],'N'.center(maxnum*2+1,'N')))
                            else: print('\033[1;31m[{1}]\033[1;36m {0}\033[1;31m - No Found\033[0m'
                                .format(s[0],'N'.center(maxnum,'N')))
                            if exact_lost:
                                for f in assfont[s][3]:
                                    print('{0}\"\033[1;31m{1}\033[0m\"'.format(''.center(maxnum, ' '), f))
                    if resultw: wfile.close()
                    print('\n【概览】')
                    for s in assfont.keys():
                        if not s[0] == s[1]:
                            print('\033[1;32m{}\033[0m'.format(assfont[s][2]))
                    for s in assfont.keys():
                        if s[0] == s[1]:
                            print('\033[1;31m{}\033[0m'.format(s[0]))
                    print('')
                else:
                    fontlist = []
                del font_info2
        elif work == 3:
            if resultw:
                resultw = False
            else:
                resultw = True
        elif work == 4:
            if copyfont:
                copyfont = False
            else:
                copyfont = True
        elif work == 5:
            if s_subdir:
                s_subdir = False
            else:
                s_subdir = True
        elif work == 6:
            if not o_fontload:
                if fontload:
                    fontload = False
                else:
                    fontload = True
            else:
                cls()
                print('在禁用系统字体源的情况下，fontload必须为True')
                os.system('pause')
        elif work == 7:
            if exact_lost:
                exact_lost = False
            else:
                exact_lost = True
        elif work == 8:
            if char_lost:
                char_lost = False
            else:
                char_lost = True
        elif work == 9:
            if embeddedFontExtract:
                embeddedFontExtract = False
            else:
                embeddedFontExtract = True
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
    font_already = []
    for nf in newfont_name.keys():
        if path.exists(newfont_name[nf][0]) and not newfont_name[nf][1] in font_already:
            if newfont_name[nf][1] is None:
                print('\033[1;31m失败:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\"\033[0m'
                .format(nf[0], path.basename(newfont_name[nf][0])))
            else:
                print('\033[1;32m成功:\033[1m \"{0}\"\033[0m >> \033[1m\"{1}\" ({2})\033[0m'
                .format(nf[0], path.basename(newfont_name[nf][0]), newfont_name[nf][1]))
        font_already.append(newfont_name[nf][1])
    del font_already


translationLang = {}

def getSubsLangs(media_ass: dict) -> list:
    '''获取字幕语言代码的用户交互部分'''
    global langlist, no_mkvm, iso639_all, preferLang, translationLang
    subs_count = {}
    sublangs = []
    for m in media_ass.keys():
        if subs_count.get(len(media_ass[m]), -1) == -1:
            subs_count[len(media_ass[m])] = [m]
        else:
            subs_count[len(media_ass[m])].append(m)
    # minsubs = min(list(subs_count.keys()))
    maxsubs = max(list(subs_count.keys()))

    cls()
    if len(iso639_all) == 0:
        getISOLangs()

    if len(translationLang) == 0:
        translationLang = iso639_all.get(preferLang, {})
        if len(translationLang) == 0:
            translationLang = iso639_all.get(preferLang.split('_')[0], {})
            if len(translationLang) == 0:
                for lk0 in iso639_all.keys():
                    if lk0.split('_')[0] == preferLang.split('_')[0]:
                        translationLang = iso639_all.get(lk0, {})
                        break
                if len(translationLang) == 0:
                    for lk0 in iso639_all.keys():
                        if preferLang.split('_')[0] in lk0:
                            translationLang = iso639_all.get(lk0, {})
                            break
                        # for lk in iso639_all[lk0].keys():
                        #     if lk in translationLang:
                        #         translationLang[lk].append(iso639_all[lk0][lk])
                        #     else:
                        #         translationLang[lk] = [iso639_all[lk0][lk]]

        if not no_mkvm:
            print('正在从mkvmerge获取ISO-639语言列表，请稍等...')
            langmkv = os.popen('mkvmerge --list-languages', mode='r')
            for s in langmkv.buffer.read().decode('utf-8').splitlines()[2:]:
                s = s.replace('\n', '').split('|')
                for si in range(1, len(s)):
                    ss = s[si]
                    if len(ss.strip(' ')) > 0:
                        langlist[ss.strip(' ').lower()] = s[0].strip(' ')
            langmkv.close()
        else:
            langlist = translationLang

    noLang = False
    if len(langlist) == 0: noLang = True
    
    reStart = True
    while reStart:
        reStart = False
        for i in range(0, maxsubs):
            
            def getLangName(code: str) -> str:
                '''获取语言代码的系统语言'''
                global iso639_all, preferLang
                showlang = translationLang.get(code, '')
                if len(showlang) == 0:
                    showlang = langlist.get(code, '')
                # elif showlang is not str:
                #     showlang = '|'.join(showlang)
                if len(showlang) == 0: return code
                return showlang

            def showFocusSub(i: int):
                print('【字幕示例】')
                for ii in range(0, len(media_ass[subs_count[maxsubs][0]])):
                    showlang = ''
                    showISO = ''
                    if ii < len(sublangs):
                        showlang = getLangName(sublangs[ii].lower())
                        if showlang == sublangs[ii].lower():
                            showlang = getLangName(sublangs[ii])
                        showISO = '{}: '.format(sublangs[ii].lower())

                    if ii == i:  
                        print('\"\033[1;33m{0}\033[0m\" \033[1;34m{2}{1}\033[0m'.format(path.basename(media_ass[subs_count[maxsubs][0]][ii]), showlang, showISO))
                    else:
                        print('\"{0}\" \033[1;34m{2}{1}\033[0m'.format(path.basename(media_ass[subs_count[maxsubs][0]][ii]), showlang, showISO))

            lang = ''
            while not (lang.lower() in langlist or lang in langlist):
                cls()
                showFocusSub(i)
                print('')
                if no_mkvm and len(translationLang) > 0:
                    searchLang = '仅本地语言搜索'
                elif no_mkvm:
                    searchLang = '该状态下不可用'
                elif len(translationLang) == 0:
                    searchLang = '仅英文English搜索'
                else:
                    searchLang = 'English/本地语言搜索'
                if len(sublangs) > i:
                    if len(sublangs[i]) > 0:
                        print('该项之前已经有赋值了，如果您不想修改，请直接回车')
                    elif no_mkvm:
                        print('可接受语言代码: ISO-639-2T/3，不区分大小写')
                    else:
                        print('可接受语言代码: ISO-639-1/2/3，不区分大小写')
                elif no_mkvm:
                    print('可接受语言代码: ISO-639-2T/3，不区分大小写')
                else:
                    print('可接受语言代码: ISO-639-1/2/3，不区分大小写')
                print('常用语言代码: [chi]中 [eng]英 [jpn]日 [kor]韩 [und]未知')
                if len(langlist) == 0: print('[WANRING] 您正在无语言编码参考文件的环境下运行，程序无法确认您输入的正确性')
                lang = input('请输入语言(或搜索语言名称，{}):'.format(searchLang)).strip(' \'\"')
                if len(lang) == 0:
                    if len(sublangs) > i:
                        if len(sublangs[i]) > 0:
                            lang = sublangs[i]
                    continue
                if noLang:
                    if len(lang) in [2,3] and len(''.join(re.findall(r'[A-Za-z]', lang))) == len(lang):
                        lang = lang.lower()
                        pass
                elif lang.lower() in langlist:
                    pass
                else:
                    hitDict = {}
                    for l in langlist.keys():
                        hits = 0
                        langKey = lang.lower().split(' ')
                        langStr = langlist[l].lower()

                        if l in translationLang:
                            transStr = translationLang[l].lower()
                            for lk in langKey:
                                if lk in transStr:
                                    hits += 1
                                    transStr = transStr.replace(lk, '')
                                if hitDict.get(hits, -1) == -1:
                                    hitDict[hits] = {langlist[l]: l}
                                else:
                                    if not hitDict[hits].get(langlist[l], False):
                                        hitDict[hits][langlist[l]] = l
                        
                        hits = 0
                        for lk in langKey:
                            if lk in langStr:
                                hits += 1
                                langStr = langStr.replace(lk, '')
                        if hitDict.get(hits, -1) == -1:
                            hitDict[hits] = {langlist[l]: l}
                        else:
                            if not hitDict[hits].get(langlist[l], False):
                                hitDict[hits][langlist[l]] = l

                    if max(list(hitDict.keys())) == 0:
                        print('找不到对应语言')
                        os.system('pause')
                        continue
                    hitDictOut = {}
                    hitIndex = -1
                    maxHits = max(list(hitDict.keys()))
                    for ki in range(1, maxHits + 1, 1):
                        print('【关键字匹配数 \033[1;33m{}\033[0m】'.format(ki))
                        for kii in range(0, len(hitDict[ki])):
                            hitIndex += 1
                            print('[\033[1;33m{0}\033[0m] {1}'.format(hitIndex, getLangName(list(hitDict[ki].values())[kii])))
                            # print(getLangName(list(hitDict[ki].values())[kii]))
                            # print('[\033[1;33m{0}\033[0m] {1}'.format(hitIndex, list(hitDict[ki].keys())[kii]))
                            hitDictOut[hitIndex] = (list(hitDict[ki].keys())[kii], list(hitDict[ki].values())[kii])
                    if len(hitDictOut) == 1:
                        lang = hitDictOut[0][1]
                    else:
                        langIndex = -1
                        while int(langIndex) not in range(0, hitIndex + 1):
                            langIndex = input('请输入方括号内数字选择语言:')
                            if not langIndex.isnumeric or len(langIndex) == 0: 
                                langIndex = -1
                                continue
                            if int(langIndex) in range(0, hitIndex + 1):
                                lang = hitDictOut[int(langIndex)][1]

                cls()
                showFocusSub(i)
                lN = getLangName(lang.lower())
                if lN == lang.lower():
                    lN = getLangName(lang)
                print('\n语言为\"\033[1;33m{0}\033[0m\"，确定吗？'.format(lN))
                if os.system('choice') == 1:
                    if i >= len(sublangs):
                        sublangs.append(lang)
                    else:
                        sublangs[i] = lang
                    cls()
                    showFocusSub(i)
                    if i == maxsubs - 1:
                        cls()
                        print('【完整预览】')
                        for subs in media_ass.values():
                            for si in range(0, len(subs)):
                                lN = getLangName(sublangs[si].lower())
                                if lN == sublangs[si].lower():
                                    lN = getLangName(sublangs[si])
                                print('\"{0}\" \033[1;34m{1}\033[0m'.format(path.basename(subs[si]), getLangName(sublangs[si].lower())))
                            print('')
                        print('【完整预览】')
                        print('请检查您的输入是否正确，要重新开始吗(R)？或是放弃(C)？按Y继续封装')
                        work = os.system('choice /C RYC')
                        if work == 1:
                            reStart = True
                        elif work == 3:
                            return []
                        else:
                            cls()
                else:
                    lang = ''
    return sublangs


def cFontSubset(font_info):
    global extlist, v_subdir, s_subdir, rmAssIn, rmAttach, fontload, \
        mkvout, assout, fontout, matchStrict, no_mkvm, notfont, warningStop, errorStop, ignoreLost, char_compatible, \
        insteadFF, noRequestFont
    leave = True
    while leave:
        cls()
        showFF = '子集化并封装'
        showFFKey = ''
        if insteadFF and not no_mkvm:
            showFF = '子集化并封装(mkvmerge)\n[D] 子集化并封装(FFmpeg)'
            showFFKey = 'D'
        elif no_mkvm and insteadFF:
            showFF = '子集化并封装(FFmpeg)'
        elif not insteadFF:
            showFF = '子集化并封装(mkvmerge)'
        print('''ASFMKV & ASFMKV-FontSubset
选择功能:
[A] 子集化字体
[B] 子集化并封装(ASS/SSA内嵌)
[C] {13}
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
[U] 广兼容性子集化: \033[1;33m{9}\033[0m
[W] 忽略字体所缺字: \033[1;33m{12}\033[0m
[X] 子集化失败中断: \033[1;33m{10}\033[0m
[Y] 忽略字幕缺少字体：\033[1;33m{14}\033[0m
[Z] 使用工作目录字体: \033[1;33m{11}\033[0m
'''.format(v_subdir, s_subdir, rmAssIn, rmAttach, mkvout, assout,
           fontout, matchStrict, notfont, char_compatible, errorStop, fontload, ignoreLost, showFF, noRequestFont))
        work = os.system('choice /M 请输入 /C AC1234567890UWXYZLB{}'.format(showFFKey))

        if work == 2 and (no_mkvm and not insteadFF):

            print('[ERROR] 在您的系统中找不到 mkvmerge 或 ffmpeg, 该功能不可用')
            work = -1

        if work == 3:
            # [1] 检视媒体扩展名列表 及 语言编码列表
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
        elif work == 4:
            # [2] 搜索子目录(视频)
            if v_subdir:
                v_subdir = False
            else:
                v_subdir = True
        elif work == 5:
            # [3] 搜索子目录(字幕)
            if s_subdir:
                s_subdir = False
            else:
                s_subdir = True
        elif work == 6:
            # [4] (封装)移除内挂字幕
            if rmAssIn:
                rmAssIn = False
            else:
                rmAssIn = True
        elif work == 7:
            # [5] (封装)移除原有附件
            if rmAttach:
                rmAttach = False
            else:
                rmAttach = True
        elif work == 8:
            # [6] (封装)不封装字体
            if notfont:
                notfont = False
            else:
                notfont = True
        elif work == 9:
            # [7] 严格字幕匹配
            if matchStrict:
                matchStrict = False
            else:
                matchStrict = True
        elif work in [10, 11, 12]:
            # [8/9/10] 输出文件夹
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
            # [U] 广兼容性子集化
            if char_compatible:
                char_compatible = False
            else:
                char_compatible = True
        elif work == 14:
            # [W] 忽略字体所缺字
            if ignoreLost:
                ignoreLost = False
            else:
                ignoreLost = True
        elif work == 15:
            # [X] 子集化失败中断
            if errorStop:
                errorStop = False
            else:
                errorStop = True
        elif work == 16:
            # [Y] 忽略字幕缺少字体
            if noRequestFont:
                noRequestFont = False
            else:
                noRequestFont = True
        elif work == 17:
            # [Z] 使用工作目录字体
            if not o_fontload:
                if fontload:
                    fontload = False
                else:
                    fontload = True
            else:
                cls()
                print('在禁用系统字体源的情况下，fontload必须为True')
                os.system('pause')
        
        elif work in [1, 2, 19, 20]:
            cls()
            if work == 1:
                print('''子集化字体
搜索子目录(视频): \033[1;33m{0}\033[0m
搜索子目录(字幕): \033[1;33m{1}\033[0m
严格字幕匹配: \033[1;33m{2}\033[0m
字幕文件输出文件夹: \033[1;33m{3}\033[0m
字体文件输出文件夹: \033[1;33m{4}\033[0m
广兼容性子集化: \033[1;33m{5}\033[0m
忽略字体所缺字: \033[1;33m{8}\033[0m
子集化失败中断: \033[1;33m{6}\033[0m
使用工作目录字体：\033[1;33m{7}\033[0m
'''.format(v_subdir, s_subdir, matchStrict, assout, fontout, char_compatible, errorStop, fontload, ignoreLost))
            else:
                print('''子集化字体并封装
搜索子目录(视频): \033[1;33m{4}\033[0m
搜索子目录(字幕): \033[1;33m{5}\033[0m
移除内挂字幕: \033[1;33m{0}\033[0m
移除原有附件: \033[1;33m{1}\033[0m
不封装字体: \033[1;33m{2}\033[0m
严格字幕匹配: \033[1;33m{6}\033[0m
媒体文件输出文件夹: \033[1;33m{3}\033[0m
广兼容性子集化: \033[1;33m{8}\033[0m
忽略字体所缺字: \033[1;33m{10}\033[0m
子集化失败中断: \033[1;33m{7}\033[0m
使用工作目录字体：\033[1;33m{9}\033[0m
'''.format(rmAssIn, rmAttach, notfont, mkvout, v_subdir, s_subdir, matchStrict, errorStop, char_compatible, fontload, ignoreLost))
            cpath = ''
            directout = False
            subonly = False
            subonlyp = []
            while not path.exists(cpath) and not directout:
                directout = True
                cpath = input('不输入任何值 直接回车回到上一页面\n请输入文件或目录路径: ').strip('\" ')
                if cpath == '':
                    print('没有输入，回到上级菜单')
                elif not path.isabs(cpath):
                    print('\033[1;31m[ERROR] 输入的必须是绝对路径！: \"{0}\"\033[0m'.format(cpath))
                elif not path.exists(cpath):
                    print('\033[1;31m[ERROR] 找不到路径: \"{0}\"\033[0m'.format(cpath))
                elif path.isfile(cpath):
                    testext = path.splitext(cpath)[1][1:].lower()
                    if testext in extlist:
                        directout = False
                    elif testext in ['ass', 'ssa'] and work in [1, 19]:
                        directout = False
                        subonly = True
                    else:
                        print('\033[1;31m[ERROR] 扩展名不正确: \"{0}\"\033[0m'.format(cpath))
                elif not path.isdir(cpath):
                    print('\033[1;31m[ERROR] 输入的应该是目录或媒体文件！: \"{0}\"\033[0m'.format(cpath))
                else:
                    directout = False
            # print(directout)
            medias = []
            if not directout:
                if path.isfile(cpath):
                    if not subonly:
                        medias = [[path.splitext(path.basename(cpath))[0], cpath]]
                    else:
                        subonlyp = [cpath]
                    cpath = path.dirname(cpath)
                else:
                    if work != 19: medias = getFileList(cpath, extlist, v_subdir)
                    else: medias = []
                    if len(medias) == 0:
                        subonlyp = getFileList(cpath, ['ass', 'ssa'], s_subdir)
                        if work == 19:
                            subonly = True
                        elif len(subonlyp) > 0:
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
                if work in [2, 20]: domux = True

                if len(medias) > 0:
                    media_ass = getSubtitles(cpath, medias)
                    if len(media_ass.values()) > 0:
                        sublangs = None
                        sublangs = []
                        muxer = 0
                        cls()
                        if domux: 
                            print('您需要为字幕轨道添加语言信息吗？')
                            if os.system('choice') == 1:
                                sublangs = getSubsLangs(media_ass)
                        # print(media_ass)
                        if work == 20 or (work == 2 and no_mkvm):
                            muxer = 1
                        elif work == 19:
                            muxer = 2
                            domux = False
                        for k in media_ass.keys():
                            # print(k)
                            # print([assout_cache, fontout_cache, mkvout_cache])
                            if fontload:
                                font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                            else:
                                font_info2 = font_info
                            newasspaths, newfont_name, mkvr = main(font_info2, media_ass[k],
                                                                mux=domux,
                                                                outdir=[assout_cache, fontout_cache, mkvout_cache],
                                                                vpath=k,
                                                                asslangs=sublangs,
                                                                FFmuxer = muxer)
                            if mkvr != -2:
                                showMessageSubset(newasspaths, newfont_name)
                            else:
                                break
                    else:
                        print('\033[1;31m[ERROR] 找不到视频对应字幕\033[0m')
                elif subonly:
                    if fontload:
                        font_info2 = templeFontLoad(cpath, copy.deepcopy(font_info))
                    else:
                        font_info2 = font_info
                    muxer = 0
                    if work == 19:
                        muxer = 2
                    for subp in subonlyp:
                        newasspaths, newfont_name, mkvr = main(font_info2, [subp],
                                                               mux=False,
                                                               outdir=[assout_cache, fontout_cache, mkvout_cache],
                                                               FFmuxer = muxer)
                        if mkvr != -2:
                            showMessageSubset(newasspaths, newfont_name)
                        else:
                            break

        else:
            leave = False
        if work < 4 or work >= 19:
            os.system('pause')


def cLicense():
    cls()
    print('''AddSubFontMKV Python Remake Preview 21

Apache-2.0 License
https://www.apache.org/licenses/

Copyright(c) 2022-2024 yyfll

Requirements:
fontTools     |  MIT License
chardet       |  LGPL-2.1 License
colorama      |  BSD-3 License

Selective Requirements:
mkvmerge      |  GPL-2 License
ffmpeg        |  LGPL-2.1 License
language-list |  MIT License
''')
    os.system('pause')


def cFontSearch(font_info: list):
    leave = True
    while leave:
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
        elif work in [2, 3]:
            cls()
            if work == 3: print('AND = \" \"（半角空格）\n仅支持AND')
            f_n = input('请输入字体名称:\n').strip(' \"').lower()
            print('')
            if len(f_n) > 0:
                hadget = False
                if work == 2:
                    f_nt = font_info[1].get(f_n)
                    if f_nt is not None:
                        f_nget = font_info[0].get(f_nt)
                        if f_nget is not None:
                            print('\033[1;31m[{0}]\033[0m {2}\033[1m\n\"{1}\"\033[0m\n'.format(f_nt, f_nget[0],
                                                                                               f_nget[2]))
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


def cSVMatching():
    global extlist, v_subdir
    cls()
    print('Subtitles-Videos Filename Matching (Experiment)')
    nomatching = 0
    sPath = input('字幕目录:').strip(' \"\'')
    while not (path.exists(sPath) and path.isdir(sPath)):
        nomatching += 1
        if nomatching > 2:
            print('回到主界面')
            break
        sPath = input('字幕目录:').strip(' \"\'')
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
                        for i in range(0, logCount + 1, -1):
                            logfile = open(path.join(sPath, logfiles[i]), mode='r', encoding='utf-8')
                            for li in logfile.readlines():
                                li = li.strip('\n').split('|')
                                if len(li) < 2: continue
                                if li[1] in renCache:
                                    renCache[path.basename(li[0])] = (renCache[li[1]][0], li[0])
                                    del renCache[li[1]]
                                else:
                                    renCache[path.basename(li[0])] = (path.join(path.dirname(li[0]), li[1]), li[0])
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
                                    print('\033[31;1m[ERROR]\033[0m 文件\"\033[1m{0}\033[0m\"重命名失败'.format(
                                        path.basename(r[0])))
                                    nodel = True
                                else:
                                    print('''\033[32;1m[SUCCESS]\033[0m \"\033[1m{0}\033[0m\"
        >>>> \"\033[33;1m{1}\033[0m\"'''.format(path.basename(r[0]), path.basename(r[1])))
                        nomatching = 3
                        if not nodel:
                            for li in logfiles:
                                try:
                                    os.remove(path.join(sPath, li))
                                except:
                                    pass
            if nomatching > 2:
                print('回到主界面')
            else:
                nomatching = 0
                vPath = input('视频目录:').strip('\"\' ')
                while not (path.exists(vPath) and path.isdir(vPath)):
                    nomatching += 1
                    if nomatching > 2:
                        print('回到主界面')
                        break
                    vPath = input('视频目录:').strip('\"\' ')
                if nomatching <= 2:
                    nameMatching(sPath, getFileList(vPath, extlist, v_subdir))
                else:
                    print('回到主界面')
        else:
            print('回到主界面')
    os.system('pause')


def TCDetect_Trans(f: str, of: str):
    t = open(f, mode='rb')
    # 使用 UniversalDetector 确定字幕文本编码（防止非标准 UTF-16LE 字幕）
    detector = UniversalDetector()
    for dt in t:
        detector.feed(dt)
        if detector.done:
            tc = detector.result['encoding']
            break
    detector.reset()
    t.close()
    if tc.lower() == 'utf-8-sig':
        del t,tc
        print('[Already UTF-8-BOM] \033[33;1m\"{0}\"\033[0m'.format(path.basename(f)))
        return None
    elif tc.lower() == 'gb2312':
        tc = 'gbk'
    t = open(f, mode='r', encoding=tc)
    o = open(of, mode='w', encoding='utf-8-sig')
    o.write(t.read())
    t.close()
    o.close()
    print('[{0} >>> UTF-8-BOM] \033[33;1m\"{1}\"\033[0m'.format(tc.upper(), path.basename(f)))
    del t,tc,o,detector


def cTextCodingTranscode():
    global s_subdir
    cls()
    print('Simple ASS/SSA/SRT to UTF-8-BOM Batch')
    sPath = ''
    loop = 0
    while not (path.isdir(sPath) or (path.isfile(sPath) and path.splitext(sPath)[1][1:].lower() in ['ass', 'ssa'])):
        loop += 1
        if loop > 4: return None
        sPath = input('字幕路径或字幕所在目录路径:').strip('\"\' ')
    loop = 0
    if path.isdir(sPath):
        while len(input('''
搜索子目录: \033[33;1m{0}\033[0m (再不输入任何值\033[33;1m{1}\033[0m次回到主菜单)
是否要将目标目录的子目录的字幕一起转换?
\033[33;1m输入任意键(除空格)并回车\033[0m进入到下一步
\033[33;1m不输入任何值\033[0m切换是否搜索子目录
'''.format(s_subdir, 7-loop)).strip(' ')) == 0:
            loop += 1
            if loop > 6: return None
            s_subdir = not s_subdir
    outPath = ''
    fn = 0
    if not path.exists(outPath): 
        try:
            os.mkdir(outPath)
        except Exception as e:
            print('\033[31;1m{}\033[0m'.format(e))
    else: fn = len(os.listdir(outPath))
    outPath = path.join(sPath, 'UTF8-Subtitles')
    if not path.exists(outPath):
        try:
            os.mkdir(outPath)
        except Exception as e:
            print('\033[31;1m{}\033[0m'.format(e))
            return None
    
    if path.isdir(sPath):
        for f in getFileList(sPath, ['ass', 'ssa', 'srt'], s_subdir):
            try:
                TCDetect_Trans(f[1], path.join(outPath, '.utf8'.join(path.splitext(path.basename(f)))))
            except Exception as e:
                print('\033[31;1m{}\033[0m'.format(e))
    else:
        try:
            TCDetect_Trans(sPath, path.join(outPath, '.utf8'.join(path.splitext(path.basename(sPath)))))
        except Exception as e:
            print('\033[31;1m{}\033[0m'.format(e))
    if not len(os.listdir(outPath)):
        try:
            os.rmdir(outPath)
        except Exception as e:
            print('\033[31;1m{}\033[0m'.format(e))
    elif not len(os.listdir(outPath)) == fn:
        print('输出文件夹: \033[33;1m\"{}\"\033[0m'.format(outPath))
    os.system('pause')

no_mkvm = False
no_cmdc = False
mkvmv = ''
ffmv = ''
font_info = [{}, {}, {}, {}]

def checkFF():
    '''检查ffmpeg和ffprobe可用性'''
    global ffmv, insteadFF
    if os.system('ffmpeg -version 1>nul') == 0 and os.system('ffprobe -version 1>nul') == 0:
        ffmv = '\n' + ' '.join(os.popen('ffmpeg -version', mode='r').readlines()[0].rstrip('\n').split(' ')[:3])
        insteadFF = True

# font_info 列表结构
#   [ font_name(dict), font_n_lower(dict), font_family(dict), warning_font(dict) ]
def loadMain(reload: bool = False):
    global extlist, no_mkvm, no_cmdc, dupfont, mkvmv, font_info, fontin, langlist, ffmv, insteadFF
    # 初始化字体列表 和 mkvmerge 相关参数
    os.system('title ASFMKV Python Remake Pre21 ^| (c) 2022-2024 yyfll ^| Apache-2.0')
    if not reload:
        if not o_fontload:
            font_list = getFontFileList(fontin)
            font_info = fontProgress(font_list, [{}, {}, {}, [], {}], f_priority)
            del font_list
        mkvmv = '\n\033[1;33m没有检测到 mkvmerge\033[0m'

        if os.system('mkvmerge -V 1>nul 2>nul') > 0:
            no_mkvm = True
            checkFF()
        else:
            print('\n\n\033[1;33m正在获取 mkvmerge 支持格式列表，请稍等...\033[0m')
            mkvmv = '\n' + os.popen('mkvmerge -V --ui-language en', mode='r').read().replace('\n', '')
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
            if usingFF: checkFF()

        if os.system('choice /? 1>nul 2>nul') > 0:
            no_cmdc = True
        # no_mkvm = True

    cls()
    if o_fontload:
        fltip = '\033[1;35m已禁用系统字体读取\033[0m'
    else:
        fltip = '包含乱码的名称'
    ffSelect = ''
    ffMessage = ''
    if not len(ffmv) > 0:
        ffMessage = '\n[F] 检查并启用FFmpeg'
        ffSelect = 'F'
    print('''ASFMKV Python Remake Pre21 | (c) 2022-2024 yyfll{0}{5}
字体名称数: [\033[1;33m{2}\033[0m]（{4}）
请选择功能:
[A] 列出字幕所用字体
[B] 字体子集化 & MKV封装
[M] 字幕-视频名称匹配
[S] 搜索字体
[W] 检视旧格式字体
[U] 字幕批量转换UTF-8-BOM
[C] 检视重复字体: 重复名称[\033[1;33m{1}\033[0m]
其他:
[R] 字体信息重载{6}
[D] 依赖与许可证
[L] 直接退出'''.format(mkvmv, len(dupfont.keys()), len(font_info[0].keys()), len(font_info[3]), fltip, ffmv, ffMessage))
    print('')
    work = os.system('choice /M 请选择: /C ABCSWDLRMU{}'.format(ffSelect))
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
        if not o_fontload:
            cFontSearch(font_info)
        else:
            cls()
            print('您已禁用系统字体读取，该功能不可用')
            os.system('pause')
    elif work == 5:
        # cOutOfDateFontsTrans()
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
            cls()
            appdata = os.getenv('APPDATA')
            fontCacheDir = path.join(appdata, 'ASFMKVpy')
            if not path.isdir(fontCacheDir): fontCacheDir = None
            print('''字体信息重载
缓存目录: \"\033[1;33m{}\033[0m\"

[0] 回到主菜单
[1] 更新缓存
[2] 仅覆盖当前使用的缓存文件并重载
[3] 删除所有缓存文件并重载
[4] 新增自定义字体目录并重载

请选择:'''.format(fontCacheDir))
            work2 = os.system('choice /C 01234')

            def reloadFont(usingCache: bool = True):
                global dupfont, fontin, font_info
                del font_info
                dupfont = None
                dupfont = {}
                font_list = getFontFileList(fontin)
                font_info = fontProgress(font_list, usingCache=usingCache)
                del font_list

            if work2 not in [2, 3, 4, 5] :
                pass
            elif work2 == 2:
                reloadFont()
            elif work2 == 3:
                reloadFont(True)
            elif work2 == 4:
                if not fontCacheDir is None:
                    for f in os.listdir(fontCacheDir):
                        if path.splitext(f)[1][1:].lower() == 'json':
                            fp = path.join(fontCacheDir, f)
                            os.remove(fp)
                reloadFont(True)
            elif work2 == 5:
                cls()
                fontPath = input('目录路径，直接回车结束:').strip(' \"\'')
                fontPaths = []
                while len(fontPath) > 0:
                    if path.isdir(fontPath):
                        fontPaths.append(fontPath)
                    else: print('\033[1;31m[ERROR]\033[0m \"\033[1;33m{}\033[0m\"非有效目录'.format(fontPath))
                    print('目录列表')
                    for p in fontPaths:
                        print('\"{}\"'.format(p))
                    print('')
                    fontPath = input('目录路径:').strip(' \"\'')
                fontin.extend(fontPaths)
                reloadFont()
    elif work == 6:
        cLicense()
    elif work == 7:
        work2 = os.system('choice /M 您真的要退出吗 /C YN')
        if work2 != 2:
            exit()
    elif work == 9:
        cSVMatching()
    elif work == 10:
        cTextCodingTranscode()
    elif work == 11:
        checkFF()
    else:
        exit()
    cls()


# loadMain()
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
