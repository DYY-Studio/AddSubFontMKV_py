# AddSubFontMKV Python Remake (ASFMKV Py)
**Copyright(c) 2022 yyfll**

**将您的字幕和字体通过mkvmerge快速批量封装到Matroska容器**

**或是查看您的系统是否拥有字幕所需的字体**

**亦或是将您的字体按照字幕进行子集化**

浪费时间打磨5年的ASFMKV批处理的正统后继者（X）

本脚本使用 Apache-2.0 许可证
## 能干什么？
* 字体查重（借由**fontTools**实现，谢谢。）
* ASS/SSA 依赖字体确认（前辈的**ListAssFont**这个名字真的非常好概括这个功能）
* 字体子集化（借由**fontTools**实现，谢谢。字体命名上参考了现有的子集化程序。）
* 字幕字体批量封装到Matroska（借由**mkvmerge**实现，谢谢）
## 安装依赖组件和程序
#### [Python](https://www.python.org/downloads/windows/)
#### 运行库
`pip3 install chardet fontTools colorama --upgrade`
#### [mkvmerge](https://mkvtoolnix.download/) (可选)

下载并安装/解压 MKVToolNix，将安装目录添加到系统变量path中
## 系统要求
**Windows 7 SP1 专业版及以上，python 3.7 及以上，cmd 必须支持 choice 命令**

** 不支持 Linux **
#### 测试环境
* Windows10 21H1、2004 及 Windows7 SP1 (已打全部补丁)
* Python 3.9.5 / 3.8.10
* mkvmerge v48.0.0
# 功能介绍
启动时需要先扫描注册表里的字体和用户给的自定义字体目录（如果有）
## ASFMKV-ListAssFont
（名称来源于前辈的软件 **ListAssFont**，这名字很精炼，我找不到词代替它）

（这一部分的输出格式很大程度上参考了前辈的设计，谢谢！）

为了更快的找到字幕中所述字体名对应的字体，ASFMKV采用了以字体名称作为Key的词典

因而这里的格式与前辈的 ListAssFont 不同
### [A] 列出全部字体
会输出 字体文件名（无路径） 与 字体名称 的对应关系
### [B] 检查并列出
需要用户输入ASS/SSA字幕文件的路径或其所在目录

会输出 字幕文件内规定的字体名称 与 字体文件名 的对应关系
## ASFMKV & ASFMKV-FontSubset
ASFMKV的传统功能以及字体子集化功能

**请参见[字体黑名单](https://github.com/DYY-Studio/AddSubFontMKV_py/wiki/%E5%AD%97%E4%BD%93%E9%BB%91%E5%90%8D%E5%8D%95-%E4%B9%B1%E7%A0%81---Fonts-Blacklist---Decode-Error)以减少子集化错误的可能性**
### [A] 子集化字体
需要用户输入有外挂字幕的视频文件的路径或其所在目录

会创建子文件夹 `subs` 和 `Fonts`（如果用户没有规定）输出子集化完成的字幕和字体

字体会在 `Fonts` 下的以各视频文件匹配到的第一个字幕的文件名作为名称文件夹下

字幕会带有`.subset`标注
### [B] 子集化并封装
需要用户输入有外挂字幕的视频文件的路径或其所在目录

会自动完成封装流程，并删除子集化的字体和修改过的字幕

默认情况下输出到源文件夹，文件名加有 `.muxed`

**也可以把 notfont 设置为 True，只封装字幕，不封装字体，不进行子集化**

## 检视重复字体: 重复名称
列出同一字体名称下的字体文件

注意：必须满足以下条件才会显示在这里
1. 文件名不相同（包括扩展名）
2. 必须是从注册表读取的系统已安装字体

## 特点
### 完全命令行界面
继承了ASFMKV的光荣传统，用choice命令实现了更流畅的命令行操作 ~~不会做也懒得做图形界面~~
### 自定义变量
继承自ASFMKV的特性，方便用户把ASFMKV调教成自己喜欢的样子。
### 多色高亮显示
借由 **colorama** 实现 Windows 7 的高亮多色显示支持，谢谢。~~要不是要支持Win7这玩意根本不会用到~~
### 低ASS/SSA格式要求
由于测试用的字幕都是十多年前的老物，各种奇葩格式已经被免疫了

（我第一次见那个……Format行到第一个Dialogue能空10多行的字幕）
### 自动字幕轨道名称
`[SweetSub&VCB-Studio] Kaijuu no Kodomo [Ma10p_1080p][x265_flac].chs.ass`中

`chs`会被作为新Matroska文件中对应的字幕轨道的名称
### 乱码字体名称正确读取
经过两日废寝忘食的研究，发现了字体名称乱码的奥秘，它能够读出准确的名称了

*参见[字体黑名单](https://github.com/DYY-Studio/AddSubFontMKV_py/wiki/%E5%AD%97%E4%BD%93%E9%BB%91%E5%90%8D%E5%8D%95-%E4%B9%B1%E7%A0%81---Fonts-Blacklist---Decode-Error)*

（但是由于使用的是自动化的 **pyftsubset**，子集化会出现故障，程序会帮你塞个原字体进去）
### 高速字体扫描
谢谢！谢谢 **fontTools**！从来没这么爽过
### 额外字体目录支持
想要子集化系统里没装有的字体？没问题！

（我自己试着扫了 超级字体整合包XZ，果然还是还是硬盘太慢的锅……）
### 字体查重
想知道你的Windows里用户字体文件夹和系统字体文件夹里到底重了多少？它可以做到！
### 字体缺失询问
子集化的时候系统缺字体了？没问题！您只要把它的路径输入进去就行了！
### 多语种字体名称读取
俄语？希腊语？或是意大利语？没问题！统统都给你读出来！~~其实是懒得做语言过滤~~
## 缺点
### 程序结构复杂
写到后面逻辑乱了，很多可以函数化的东西被反复用了
### 错误易崩溃
对错误的预防只有最低限度，一旦错了就炸了，没有Module_DEBUG的第一天……
### 注释不够多
自己读起来都痛苦
## 高级
### 崩溃反馈
1. 打开 命令提示行/cmd ，输入`py "脚本绝对路径"`，重复您崩溃前所做的操作<br>或 使用有断点调试功能的编辑器（如VS Code），对该脚本进行调试，重复您崩溃前所做的操作

2. 发送崩溃信息 及 致使崩溃的操作 到作者邮箱
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
## 未实装功能
#### sublang
目前暂不能在运行过程中询问用户来获取语言编码
## 未测试功能
#### mkvout, fontout, assout
设置为绝对路径的情况暂未测试
## 致谢
### [fontTools](https://github.com/fonttools/fonttools) (MIT Licence)
### [MKVToolNix/mkvmerge](https://mkvtoolnix.download/) (GPLv2 Licence)
### [chardet](https://github.com/chardet/chardet) (LGPLv2.1 Licence)
### [colorama](https://github.com/tartley/colorama) (BSD-3-Clause License)
