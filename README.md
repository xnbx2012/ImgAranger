# ImgArranger

> 一个简单的，用Python写的图片排版工具。

> A simple images arrange tool written in Python



## 主要功能 Main Function

从docx文档中读取图片并按照图片在文档中出现的顺序，统一宽度，按照用户设定的分栏和页边距排版到新的docx文档中。原本设计是应用于错题收集，从错题软件导出的Word文档中每一道错题都是一张图片，因此需要进行合理的排版进行打印，一般是分为两栏。当然，也可应用于其它图片排版的需求当中



## 直接运行
在Windows平台下，克隆本仓库，双击`run.bat`即可运行



## 打包编译

### Linux

在项目根目录执行以下命令编译二进制文件

```bash
pyinstaller -F -p ./venvl/lib/python3.7/site-packages/ mistake_arrange.py
```
打包成二进制文件后需要将`mistake_arr.qss`放在

