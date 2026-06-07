<div align="center">

# ImgArranger

**一个用 Python 写的图片排版工具 · A simple images arrange tool written in Python**

[![License](https://img.shields.io/github/license/xnbx2012/ImgAranger)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)](#)
[![GitHub stars](https://img.shields.io/github/stars/xnbx2012/ImgAranger.svg)](https://github.com/xnbx2012/ImgAranger/stargazers)

[简体中文](#简体中文) · [English](#english)

</div>

---

## 简体中文

### 项目简介

**ImgArranger** 是一个使用 Python 编写的简单图片排版工具。它能够从 `.docx` 文档中提取所有图片，并按照它们在原文档中出现的顺序，以统一宽度重新排版到一份新的 `.docx` 文档中。

### 主要功能

- 📄 **从 docx 提取图片**：自动读取 Word 文档中嵌入的所有图片
- 🧩 **多栏排版**：支持自定义分栏数（默认 2 栏，适合错题打印）
- 📐 **自定义页边距**：灵活调整上下左右页边距
- 🖨️ **支持多种纸张尺寸**：内置 A4、A5、B5 等常用纸张规格
- 💾 **可视化操作**：基于 PyQt5 的图形界面，操作简单直观
- 📦 **可打包为二进制**：支持使用 PyInstaller 打包为独立可执行文件

### 适用场景

原本设计是应用于**错题整理**——从错题软件导出的 Word 文档中，每道错题都是一张图片，使用本工具可以将其合理排版后打印。当然，也可以应用于其他任何需要批量图片排版的场景。

### 安装要求

- Python 3.7+
- 操作系统：Windows / Linux / macOS

### 快速开始

#### Windows 用户（推荐）

直接双击 `run.bat` 即可运行：

```bat
run.bat
```

#### 命令行运行

```bash
# 1. 克隆仓库
git clone https://github.com/xnbx2012/ImgAranger.git
cd ImgAranger

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python mistake_arrange.py
```

### 使用说明

1. 点击 **「选择文件」** 选取包含图片的 `.docx` 源文件
2. 点击 **「另存为」** 选择输出文件路径
3. 设置分栏数、页边距、纸张大小等参数
4. 点击 **「开始排版」** 等待处理完成

### 打包编译

#### Linux

在项目根目录执行：

```bash
pyinstaller -F -p ./venvl/lib/python3.7/site-packages/ mistake_arrange.py
```

打包完成后需要将 `mistake_arr.qss` 样式文件放在可执行文件同目录下。

#### Windows

```bash
pyinstaller -F -w mistake_arrange.py
```

### 项目结构

```
ImgAranger/
├── mistake_arrange.py      # 主程序入口
├── mistake_arr.py          # PyQt5 UI 代码（由 .ui 文件生成）
├── mistake_arr.ui          # Qt Designer 界面文件
├── mistake_arr.qss         # 样式表
├── requirements.txt        # Python 依赖
├── run.bat                 # Windows 启动脚本
├── demo.docx               # 示例文档
└── LICENSE                 # MIT 许可证
```

### 依赖说明

| 包名 | 用途 |
| ---- | ---- |
| PyQt5 | 图形界面框架 |
| python-docx | 操作 Word 文档 |
| Pillow | 图片处理 |
| lxml | XML 解析（python-docx 依赖） |

### 许可证

本项目基于 [MIT License](LICENSE) 开源。

### 贡献

欢迎提交 Issue 和 Pull Request！

---

## English

### Introduction

**ImgArranger** is a simple image arrangement tool written in Python. It extracts all images embedded in a `.docx` document, then re-arranges them into a new `.docx` file with a unified width, preserving the original order.

### Features

- 📄 **Extract images from docx**: Automatically read all images embedded in a Word document
- 🧩 **Multi-column layout**: Support custom number of columns (default 2, ideal for printing mistake collections)
- 📐 **Custom margins**: Flexibly adjust top, bottom, left, and right margins
- 🖨️ **Multiple paper sizes**: Built-in support for A4, A5, B5, and other common paper sizes
- 💾 **GUI-based**: Intuitive graphical interface built with PyQt5
- 📦 **Packaging support**: Can be packaged as a standalone executable using PyInstaller

### Use Case

The tool was originally designed for **collecting and printing wrong/mistake questions** — each question in a Word document exported from a mistake-collection app is a single image. This tool helps you arrange them in a printable layout. Of course, it can also be used in any other scenario that requires batch image arrangement.

### Requirements

- Python 3.7+
- OS: Windows / Linux / macOS

### Quick Start

#### Windows (Recommended)

Simply double-click `run.bat` to run.

#### Run from command line

```bash
# 1. Clone the repository
git clone https://github.com/xnbx2012/ImgAranger.git
cd ImgAranger

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the application
python mistake_arrange.py
```

### Usage

1. Click **「Select File」** to choose a `.docx` source file containing images
2. Click **「Save As」** to choose the output file path
3. Configure columns, margins, and paper size
4. Click **「Start」** and wait for processing to complete

### Build / Package

#### Linux

Run in the project root:

```bash
pyinstaller -F -p ./venvl/lib/python3.7/site-packages/ mistake_arrange.py
```

After packaging, place `mistake_arr.qss` in the same directory as the executable.

#### Windows

```bash
pyinstaller -F -w mistake_arrange.py
```

### Project Structure

```
ImgAranger/
├── mistake_arrange.py      # Main entry point
├── mistake_arr.py          # PyQt5 UI code (generated from .ui file)
├── mistake_arr.ui          # Qt Designer UI definition
├── mistake_arr.qss         # Stylesheet
├── requirements.txt        # Python dependencies
├── run.bat                 # Windows launcher
├── demo.docx               # Sample document
└── LICENSE                 # MIT License
```

### Dependencies

| Package | Purpose |
| ------- | ------- |
| PyQt5 | GUI framework |
| python-docx | Word document manipulation |
| Pillow | Image processing |
| lxml | XML parsing (required by python-docx) |

### License

This project is open-sourced under the [MIT License](LICENSE).

### Contributing

Issues and Pull Requests are welcome!

---

<div align="center">

⭐ 如果这个项目对你有帮助，请给它一个 Star！
⭐ If this project helps you, please give it a star!

</div>
