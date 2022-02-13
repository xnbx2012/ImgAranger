#-- coding:utf8 --
import os
import sys
import shutil
from PyQt5 import QtWidgets
import PyQt5.sip
import PyQt5
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import docx
from docx.oxml.ns import qn
from pathlib import Path

from mistake_arr import Ui_Form
import zipfile
from docx import Document
from docx.shared import Mm

qt_form = ""


class businessLogic():
    list_to_delete = []
    global qt_form

    def log(self, content):
        print("[INFO]" + content)

    def unZip(self, file_name):
        """unzip zip file"""
        self.log("正在解压文件")
        zip_file = zipfile.ZipFile(file_name)
        if os.path.isdir(file_name + "_files"):
            pass
        else:
            os.mkdir(file_name + "_files")
        for names in zip_file.namelist():
            zip_file.extract(names, file_name + "_files/")
        zip_file.close()
        self.list_to_delete.append(file_name + "_files")
        self.log("文件解压完成")

    def picGet(self, sourceFile, targetDir, leftMargin, rightMargin, topMargin, bottomMargin, column, size):
        delList=[sourceFile + ".zip_files/",sourceFile + ".zip"]
        self.log("正在提取Word文档中的图片")
        self.log("正在复制并重命名文件为zip")
        shutil.copyfile(sourceFile, sourceFile + ".zip")
        self.list_to_delete.append(sourceFile + ".zip")
        self.unZip(sourceFile + ".zip")
        document = Document()
        section = document.sections[0]
        sectPr = section._sectPr
        cols = sectPr.xpath('./w:cols')[0]
        self.log("分栏："+str(column))
        cols.set(qn('w:num'), str(column))
        index = 1
        section.top_margin = Mm(topMargin)
        section.bottom_margin = Mm(bottomMargin)
        section.left_margin = Mm(leftMargin)
        section.right_margin = Mm(rightMargin)
        paper_size=size.split("|")[1]
        width=int(paper_size.split("*")[0])
        #height=int(paper_size.split("*")[1])
        section.page_width = Mm(width)
        #section.page_height = Mm(height)
        img_width = (width-leftMargin-rightMargin-3)/column
        #img_height = (height-topMargin-bottomMargin-5)/5
        while index < 10:
            if os.path.isfile(sourceFile + ".zip" + "_files/word/media/image" + str(index) + ".png"):
                self.log("正在重命名排序，当前处理序号：" + str(index))
                shutil.move(sourceFile + ".zip" + "_files/word/media/image" + str(index) + ".png",
                            sourceFile + ".zip" + "_files/word/media/image0" + str(index) + ".png")
            index = index + 1
        fileList = os.listdir(sourceFile + ".zip" + "_files/word/media")
        total= len(fileList)
        current=0
        for file in fileList:
            current=current+1
            qt_form.progressBar.setProperty("value", (current/total)*100)
            if file.startswith("image"):
                self.log("图片" + file + "已提取")
                document.add_picture(sourceFile + ".zip" + "_files/word/media/" + file, width=Mm(img_width))
            else:
                self.log(file + "非图片，已跳过")
        self.log("正在保存文档，并删除垃圾文件")
        document.save(targetDir)
        for trash in delList:
            tFile = Path(trash)
            if tFile.is_dir():
                shutil.rmtree(trash)
            else:
                os.remove(trash)

        self.log("任务完成")


class uiEvent(QtWidgets.QWidget, Ui_Form):
    inputPath = ""
    outputPath = ""

    def __init__(self):
        super(uiEvent, self).__init__()
        self.setupUi(self)

    # 实现pushButton_click()函数，textEdit是我们放上去的文本框的id
    def inputSelect(self):
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", os.getcwd(),
                                                                   "All Files(*);;Text Files(*.txt)")
        self.input_box.setText(fileName)
        self.inputPath = fileName

    def outputSelect(self):
        file_path, file_type = QtWidgets.QFileDialog.getSaveFileName(self, "文件另存为", "C:/Users/PeterZhong/Desktop",
                                                          "docx files (*.docx);;All files(*.*)")
        self.output_box.setText(file_path)
        self.outputPath = file_path

    def start(self):
        print("主UI开始处理")
        bl = businessLogic()
        bl.picGet(self.inputPath, self.outputPath,self.doubleSpinBox.value(),self.doubleSpinBox_2.value(),self.doubleSpinBox_3.value(),self.doubleSpinBox_4.value(),self.spinBox.value(),self.comboBox.currentText())
        self.hint("文档保存成功！")

    def hint(self, text):
        QMessageBox.information(self, '提示', text)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    my_pyqt_form = uiEvent()
    qt_form = my_pyqt_form
    my_pyqt_form.show()
    sys.exit(app.exec_())
