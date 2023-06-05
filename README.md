# AddSubFontMKV Python Remake (ASFMKV Py)
**Copyright(c) 2022 yyfll**

**将您的字幕和字体通过mkvmerge快速批量封装到Matroska容器**

**或是查看您的系统是否拥有字幕所需的字体**

**亦或是将您的字体按照字幕进行子集化**

浪费时间打磨5年的ASFMKV批处理的正统后继者（X）

本脚本使用 Apache-2.0 许可证
## 关于Preview
由于部分功能是新开发的，目前仍存在多多少少的小问题<br>在本脚本最需要版本号蹭蹭蹭蹦的时候我个人非常忙，真是非常抱歉

**尽管目前的最新版本仍然是Preview版本，但是它们都相对稳定而可靠**

**您不应该使用过时的Preview16之前的任何版本，它们甚至不能正确分析ASS文件**

## 目录
**新版本在Release，仓库还没得空整理暂时不放**
| [最新更新](#CLI命令行版本更新) | [能干什么](#%E8%83%BD%E5%B9%B2%E4%BB%80%E4%B9%88) | [运行环境](#%E5%AE%89%E8%A3%85%E4%BE%9D%E8%B5%96%E7%BB%84%E4%BB%B6%E5%92%8C%E7%A8%8B%E5%BA%8F) | [功能介绍](#%E5%8A%9F%E8%83%BD%E4%BB%8B%E7%BB%8D) | [自定义变量](#%E8%87%AA%E5%AE%9A%E4%B9%89%E5%8F%98%E9%87%8F) | [字体信息缓存](#字体信息缓存) | [sublangs本地化](#sublangs本地化)
| --- | --- | --- | --- | --- | --- | --- |

## 版本前瞻（画大饼）
* PyQt Beta3版本

## GUI图形化版本更新
**注意: GUI版本目前仍未修复Preview14及之后的修复内容，极不推荐使用**
### PyQt Beta2 (实验性)
* 更新了 Py1.02 Preview10 中修复的内容
* 添加了目录选择窗口
#### Fix 2
* 修复了QLineEdit无法接收拖入路径的问题
#### Fix 3
* 补充了忘记添加的全部数字保留
* 修复了在仅有字幕目录而无视频目录时，“子集化字体”导致程序崩溃的问题
#### Fix 4
* **重要：修复了单行多fn标签时fn标签切割错误的问题**
## CLI命令行版本更新
### Preview 17
* 字体信息缓存
* 可禁用启动时字体名称修正信息、字体重复信息
* 字幕轨道语言编码询问

## 能干什么？
* 字体查重（借由**fontTools**实现，谢谢。）
* ASS/SSA 依赖字体确认
* 字体子集化（借由**fontTools**实现，谢谢。字体命名上参考了现有的子集化程序。）
* 字幕字体批量封装到Matroska（借由**mkvmerge**实现，谢谢）
* 字幕-视频文件名称匹配（**[RenameSubtitles_RE](https://github.com/DYY-Studio/RenameSubtitles_RE)** 1.1方法的改进版本）
* ASS/SSA 批量转换为 UTF-8-BOM
## 安装依赖组件和程序
#### [Python](https://www.python.org/downloads/windows/)
#### 运行库
* CLI版本: `pip3 install chardet fontTools colorama --upgrade`
* GUI版本: `pip3 install chardet fontTools colorama PyQt5 --upgrade`

#### [mkvmerge](https://mkvtoolnix.download/) (可选)

下载并安装/解压 MKVToolNix，将安装目录添加到系统变量path中
#### [FFmpeg](https://ffmpeg.org/download.html)（Preview18+ 可选）

下载并解压 FFmpeg与FFprobe，将可执行文件所在目录添加到系统变量path中，或直接与ASFMKVpy放置在同一目录

推荐使用 FFmpeg 5.0 以上版本以避免兼容性问题
## 系统要求
**Windows 7 SP1 专业版及以上，python 3.7 及以上，cmd 必须支持 choice 命令**

** 不支持 Linux **
#### 测试环境
* Windows10 21H2、21H1、20H1(2004) 及 Windows7 SP1 (已打全部补丁)
* Python 3.9.5 / 3.8.10
* mkvmerge v48.0.0 / v65.0.0
* FFmpeg v5.1.2

# 功能介绍
| K | 功能 | K | 功能 |
|---------|---------|---------|---------|
| A | [列出所需字体](#列出所需字体) | W | [检视旧格式字体](#检视旧格式字体) |
| B | [字体子集化 & MKV封装](#字体子集化及封装)<br>[**sublangs教程(Preview17+)**](#sublangs教程) | U | [字幕批量转换UTF-8-BOM](#字幕文本编码批量转换) |
| M | [字幕-视频名称匹配](#字幕-视频名称匹配) | C | [检视重复名称字体](#检视重复名称字体) |
| S | [搜索字体(Debug)](#搜索字体) | R | [字体信息重载](#字体信息重载) |

启动时需要先扫描注册表里的字体和用户给的自定义字体目录（如果有）
## 列出所需字体
（名称来源于前辈的软件 **ListAssFont**，这名字很精炼，我找不到词代替它）

（这一部分的输出格式很大程度上参考了前辈的设计，谢谢！）

为了更快的找到字幕中所述字体名对应的字体，ASFMKV采用了以字体名称作为Key的词典

因而这里的格式与前辈的 ListAssFont 不同
### [A] 列出全部字体
会输出 字体文件名（无路径） 与 字体名称 的对应关系
### [B] 检查并列出
需要用户输入ASS/SSA字幕文件的路径或其所在目录

会输出 字幕文件内规定的字体名称 与 字体文件名 的对应关系
## 字体子集化及封装
ASFMKV的传统功能以及字体子集化功能

**请参见[字体黑名单](https://github.com/DYY-Studio/AddSubFontMKV_py/wiki/%E5%AD%97%E4%BD%93%E9%BB%91%E5%90%8D%E5%8D%95-%E4%B9%B1%E7%A0%81---Fonts-Blacklist---Decode-Error)以减少子集化错误的可能性**
### [A] 子集化字体
需要用户输入有外挂字幕的视频文件的路径或其所在目录

会创建子文件夹 `subs` 和 `Fonts`（如果用户没有规定）输出子集化完成的字幕和字体

字体会在 `Fonts` 下的以各视频文件匹配到的第一个字幕的文件名作为名称文件夹下

字幕会带有`.subset`标注
### [B/C] 子集化并封装
需要用户输入有外挂字幕的视频文件的路径或其所在目录

会自动完成封装流程，并删除子集化的字体和修改过的字幕

默认情况下输出到源文件夹，文件名加有 `.muxed`

**也可以把 notfont 设置为 True，只封装字幕，不封装字体，不进行子集化**
> mkvmerge

传统的使用mkvmerge封装，经过大量测试，比较稳定

需要 `mkvmerge.exe` 在`path`系统变量中的路径下或与ASFMKVpy在同一目录
> FFmpeg

**Preview 18**新增的使用ffmpeg进行MKV封装，测试较少，仅保证部分情况正常使用

默认用于没有mkvmerge的环境，并且由于ffmpeg的元数据写入方式，同时需要ffprobe支持

需要您的ffmpeg版本较新，支持`-attach`，建议使用5.0以上版本

需要 `ffmpeg.exe`和`ffprobe.exe` 在`path`系统变量中的路径下或与ASFMKVpy在同一目录

### sublangs教程
在 Preview 17 中，我加入了拖了大半年的sublangs询问输入功能，在用户交互上废了很多心思

但是使用起来还是比较麻烦，所以在这里特别讲一下

* *演示用视频文件: [SFEO-Raws] ACCA13区监察课 ACCA 13-ku Kansatsu-ka 01-12+SP (BD 720P x264 10bit AAC)*
* *演示用字幕文件: 动漫国字幕组 ACCA13区监察课*

1. 当你在Preview17+使用 **子集化字体并封装** 功能时，输入路径回车后，会询问您是否要为字幕轨道添加语言信息<br>
![图片](https://github.com/DYY-Studio/AddSubFontMKV_py/assets/48157880/37558f01-038c-40c3-ade7-3f9c6867fc7c)
2. 选择`Y`后，程序需要从mkvmerge命令行获取ISO-639语言列表，由于语言较多，可能比较耗时
3. 然后您就进入了`sublangs`输入界面<br>![图片](https://github.com/DYY-Studio/AddSubFontMKV_py/assets/48157880/6429dc2e-9ad1-4aa1-a1e1-dbce4d821c56)<br>
示例字幕是您输入的目录中，**外挂字幕最多的视频文件**对应的外挂字幕<br>
在该界面，高亮的示例字幕代表您目前输入的语言代码所**要应用到的字幕**<br>其它视频文件的外挂字幕，也将按照这个名称顺序应用，**请避免在使用该功能时混用名称排序不同的字幕**
4. 您可以输入 ISO-639-1/2/3 语言代码，也可以输入该语言对应的英文名称（mkvmerge的语言列表没有其它语言支持）
<br>如要输入中文，直接输入ISO-639-3 `chi`，或ISO-639-2 `zh`
<br>也可以可以搜索`Chinese`
<br>![图片](https://github.com/DYY-Studio/AddSubFontMKV_py/assets/48157880/bc99406b-94bc-4104-a353-cb18e5a9a964)
<br>然后选择`4`
<br>![图片](https://github.com/DYY-Studio/AddSubFontMKV_py/assets/48157880/59767022-e1a0-4606-a8ab-76c0eefba8c0)
5. 当您确认了输入的语言后，在刚刚的字幕示例的末尾，会显示您选择的语言的英文名称，程序则会向您请求下一个字幕的语言
<br>![图片](https://github.com/DYY-Studio/AddSubFontMKV_py/assets/48157880/99be82a4-7384-458b-9b91-59e8fe0cd37c)
6. 在全部字幕示例的语言代码输入完毕后，程序会向您展示全部字幕的语言代码应用情况
<br>如果您认为之前的输入有误，您可以在这里输入`R`，如果您突然不想用语言代码了，可以输入`C`，否则输入`Y`继续封装

附上一些常用的ISO-639-3编码对应
| 语言 | ISO-639-3 | 语言 | ISO-639-3
| --- | --- | --- | --- |
| 中文 | chi | 未知 | und |
| 日文 | jpn | 韩文 | kor |
| 英文 | eng | 法语 | fre |

您可以使用 `mkvmerge --list-languages` 来查看全部语言

## 字幕-视频名称匹配
本功能是 **[RenameSubtitles_RE](https://github.com/DYY-Studio/RenameSubtitles_RE)** 项目1.1方法的python实现

通过简单运算得到视频文件名和字幕文件名中剧集数所在位置，然后直接套用，具体算法参见隔壁项目页

功能特色
* 非常快
* 有一定的准确度
* 支持多字幕
* 支持自动字幕后缀信息，或者手动输入多字幕后缀信息

本功能需要您的字幕文件和视频文件使用相同的集数表达，`5.1`和`5.5`，`OVA01`和`OVA1`是无法匹配的

## 搜索字体
DEBUG用的功能，可以在程序读取的所有字体信息中搜索字体
### 完全匹配
就是完全匹配，只有输入的名称与字体信息中的名称完全相符才会显示
### 部分匹配
只支持` `（半角空格）作为"与（AND）"操作符

在字体名称中有对应关键词就会输出

## 检视重复名称字体
列出同一字体名称下的字体文件

注意：必须满足以下条件才会显示在这里
1. 文件名不相同（包括扩展名）
2. 必须是从注册表读取的系统已安装字体

## 检视旧格式字体
在fontTools对字体的读取与子集化进行下一步的修正前，这里会列出与fontTools不兼容的旧格式字体

这些字体很有可能不能正常的被子集化，您可能需要对这些字体使用字体编辑软件重新保存来兼容fontTools

## 字幕文本编码批量转换
就是单纯的把你的 `.ass` `.ssa` `.srt`文件转为UTF-8-BOM以便其它应用使用，基于chardet编码探测

## 字体信息重载
### 1 更新缓存
以标准启动方式重新加载字体信息
### 2 仅覆盖当前使用的缓存文件并重载
完全重建当前使用的字体信息缓存
### 3 删除所有缓存文件并重载
把现在没使用的字体信息缓存也删了
### 4 新增自定义字体目录并重载
允许用户添加多个自定义字体目录

## 高级
### 崩溃反馈
1. 打开 命令提示行/cmd ，输入`py "脚本绝对路径"`，重复您崩溃前所做的操作<br>或 使用有断点调试功能的编辑器（如VS Code），对该脚本进行调试，重复您崩溃前所做的操作

2. 发送崩溃信息 及 致使崩溃的操作 到作者邮箱，或在GitHub提出issue
### 自定义变量
自定义变量是沿袭自ASFMKV批处理版本的高级自定义操作

在Python Remake中，大部分自定义变量都可以在运行过程中快捷更改

如果您想要更改这些变量的默认值，或是更改部分无法在运行中更改的变量，请往下看
#### 变更自定义变量
1. 使用编辑器打开本脚本<br>（注意，您的编辑器必须支持 UTF-8 NoBOM，不建议Win7用户使用记事本打开）
2. 在`import`部分的下方就是本脚本的自定义变量<br>编辑时请注意语法，不要移除单引号，绝对路径需要添加r在左侧引号前，右侧引号前不应有反斜杠
3. 保存
#### 运行时不能变更的自定义变量
| 变量名 | 作用 |
| --- | --- |
| extlist | 指定视频文件的扩展名 |
| fontin | 外部字体文件夹的路径 |
| sublang | 给字幕轨道赋予语言编码 |
| o_fontload | 禁用注册表字体库和自定义字体库 |
| s_fontload | 加载工作目录的子目录中的字体 |
| f_priority | 控制各字体源的优先顺序 |
| no_extcheck | 关闭mkvmerge兼容扩展名检查 |
## 字体信息缓存
该部分将介绍 Preview 17 版本中字体信息缓存的实现方式

使用JSON配置实在是太简单了，所以缓存以JSON形式保存，在程序中则保存为一个词典

缓存的JSON保存在`%APPDATA%/ASFMKVpy`下
### 目录索引
Preview 17 字体信息缓存采用目录形式缓存，将同一目录（不包括子目录）下的字体信息包含在一个JSON文件中

JSON文件名称即为索引，是 **目录绝对路径(转小写)再使用UTF-8编码** 的 **CRC32** 校验值

因此缓存采用随用随读的形式，程序校验文件所在目录路径再寻找对应的缓存文件。

### 字体索引
为了防止字体变更导致信息失效，字体索引是由三个成分构成一段**CRC32校验值**

**字体文件绝对路径(转小写) + 字体文件大小(Byte) + 文件变更时间(Float) 再使用 UTF-8 编码**

## sublangs本地化
从 **Preview 18** 版本开始，sublangs选择器支持ISO-639语言的本地化名称读取

您需要将本地化文件放入 `%APPDATA%\ASFMKVpy\isoLang\对应语言的ISO639-1语言代码`

如简体中文应将本地化文件放入 `%APPDATA%\ASFMKVpy\isoLang\zh_CN`

程序在启动时，会通过即将过时的`locale.getdefaultlocale()`获取当前系统所用语言的ISO639-1语言代码

在执行sublangs选择器时，程序会扫描上述文件夹寻找本地语言对应的本地化文件，如果没有子集(如zh_CN)，则寻找同一语言下的其它子集(zh、zh_Hans...)

每次在Release时我都会提供一个全语言本地化压缩包，请根据自己所在地语言选择放入

**提供的本地化语言包是 [language-list](https://github.com/umpirsky/language-list) 项目和 [pycountry](https://github.com/flyingcircusio/pycountry) 项目的混合**<br>
**前者提供了ISO-639-1对应的翻译文本，而后者提供了ISO-639-2T/2B/3中对应的语言代码**<br>
**由衷地感谢这两个项目**

### 自定义本地化文件

本地化JSON应遵循以下语法:
`{"[ISO-639-1/2B/2T/3]": "[对应的本地化名称]"}`

## 未测试功能
#### mkvout, fontout, assout
设置为绝对路径的情况暂未测试
## 缺点
### 程序结构复杂
写到后面逻辑乱了，很多可以函数化的东西被反复用了
### 错误易崩溃
对错误的预防只有最低限度，一旦错了就炸了，没有Module_DEBUG的第一天……
### 注释不够多
自己读起来都痛苦
## 致谢
### [fontTools](https://github.com/fonttools/fonttools) (MIT Licence)
### [MKVToolNix/mkvmerge](https://mkvtoolnix.download/) (GPLv2 Licence)
### [chardet](https://github.com/chardet/chardet) (LGPLv2.1 Licence)
### [colorama](https://github.com/tartley/colorama) (BSD-3-Clause License)
### [language-list](https://github.com/umpirsky/language-list) (MIT License)
### [pycountry](https://github.com/flyingcircusio/pycountry) (LGPLv2.1 License)
