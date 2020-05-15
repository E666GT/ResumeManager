# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'MainUI.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(679, 495)
        Form.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor))
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(310, 170, 221, 321))
        self.label.setMouseTracking(True)
        self.label.setObjectName("label")
        self.label_doc_img = QtWidgets.QLabel(Form)
        self.label_doc_img.setGeometry(QtCore.QRect(40, 350, 101, 101))
        self.label_doc_img.setObjectName("label_doc_img")
        self.label_excel_img = QtWidgets.QLabel(Form)
        self.label_excel_img.setGeometry(QtCore.QRect(40, 140, 111, 131))
        self.label_excel_img.setObjectName("label_excel_img")
        self.button_word2excel = QtWidgets.QPushButton(Form)
        self.button_word2excel.setGeometry(QtCore.QRect(40, 270, 21, 91))
        self.button_word2excel.setObjectName("button_word2excel")
        self.button_excel2word = QtWidgets.QPushButton(Form)
        self.button_excel2word.setGeometry(QtCore.QRect(120, 270, 21, 91))
        self.button_excel2word.setObjectName("button_excel2word")
        self.button_excel2cv = QtWidgets.QPushButton(Form)
        self.button_excel2cv.setGeometry(QtCore.QRect(170, 200, 101, 31))
        self.button_excel2cv.setObjectName("button_excel2cv")
        self.check_ispdf = QtWidgets.QCheckBox(Form)
        self.check_ispdf.setGeometry(QtCore.QRect(230, 150, 41, 16))
        self.check_ispdf.setObjectName("check_ispdf")
        self.formLayoutWidget = QtWidgets.QWidget(Form)
        self.formLayoutWidget.setGeometry(QtCore.QRect(40, 30, 251, 100))
        self.formLayoutWidget.setObjectName("formLayoutWidget")
        self.formLayout = QtWidgets.QFormLayout(self.formLayoutWidget)
        self.formLayout.setContentsMargins(0, 0, 0, 0)
        self.formLayout.setObjectName("formLayout")
        self.excelpath_label = QtWidgets.QLabel(self.formLayoutWidget)
        self.excelpath_label.setObjectName("excelpath_label")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.excelpath_label)
        self.cvtype_label = QtWidgets.QLabel(self.formLayoutWidget)
        self.cvtype_label.setObjectName("cvtype_label")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.cvtype_label)
        self.templeteword_path_label = QtWidgets.QLabel(self.formLayoutWidget)
        self.templeteword_path_label.setObjectName("templeteword_path_label")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.templeteword_path_label)
        self.excelpath_lineEdit = QtWidgets.QLineEdit(self.formLayoutWidget)
        self.excelpath_lineEdit.setObjectName("excelpath_lineEdit")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.excelpath_lineEdit)
        self.wordpath_lineEdit = QtWidgets.QLineEdit(self.formLayoutWidget)
        self.wordpath_lineEdit.setObjectName("wordpath_lineEdit")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.wordpath_lineEdit)
        self.cvtype_lineEdit = QtWidgets.QLineEdit(self.formLayoutWidget)
        self.cvtype_lineEdit.setObjectName("cvtype_lineEdit")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.cvtype_lineEdit)
        self.templete_word_lineEdit = QtWidgets.QLineEdit(self.formLayoutWidget)
        self.templete_word_lineEdit.setObjectName("templete_word_lineEdit")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.templete_word_lineEdit)
        self.wordpath_label = QtWidgets.QLabel(self.formLayoutWidget)
        self.wordpath_label.setObjectName("wordpath_label")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.wordpath_label)
        self.DebugBrowser = QtWidgets.QTextBrowser(Form)
        self.DebugBrowser.setGeometry(QtCore.QRect(420, 30, 241, 151))
        self.DebugBrowser.setObjectName("DebugBrowser")
        self.checkBox_LangChinese = QtWidgets.QCheckBox(Form)
        self.checkBox_LangChinese.setGeometry(QtCore.QRect(230, 130, 71, 16))
        self.checkBox_LangChinese.setObjectName("checkBox_LangChinese")
        self.label_editdoc = QtWidgets.QLabel(Form)
        self.label_editdoc.setGeometry(QtCore.QRect(20, 450, 201, 16))
        self.label_editdoc.setObjectName("label_editdoc")
        self.label_allinone = QtWidgets.QLabel(Form)
        self.label_allinone.setGeometry(QtCore.QRect(20, 250, 201, 16))
        self.label_allinone.setObjectName("label_allinone")
        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(310, 10, 54, 12))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setGeometry(QtCore.QRect(420, 10, 54, 12))
        self.label_5.setObjectName("label_5")
        self.listWidget_labels = QtWidgets.QListWidget(Form)
        self.listWidget_labels.setGeometry(QtCore.QRect(310, 30, 101, 151))
        self.listWidget_labels.setObjectName("listWidget_labels")
        self.checkBox_export_all_labels_cv = QtWidgets.QCheckBox(Form)
        self.checkBox_export_all_labels_cv.setGeometry(QtCore.QRect(520, 200, 141, 16))
        self.checkBox_export_all_labels_cv.setObjectName("checkBox_export_all_labels_cv")
        self.pushButton_openExcel = QtWidgets.QPushButton(Form)
        self.pushButton_openExcel.setGeometry(QtCore.QRect(40, 220, 91, 31))
        self.pushButton_openExcel.setMinimumSize(QtCore.QSize(91, 0))
        self.pushButton_openExcel.setObjectName("pushButton_openExcel")
        self.pushButton_openDoc = QtWidgets.QPushButton(Form)
        self.pushButton_openDoc.setGeometry(QtCore.QRect(50, 420, 91, 31))
        self.pushButton_openDoc.setMinimumSize(QtCore.QSize(91, 0))
        self.pushButton_openDoc.setObjectName("pushButton_openDoc")
        self.label_block = QtWidgets.QLabel(Form)
        self.label_block.setGeometry(QtCore.QRect(0, 0, 681, 501))
        self.label_block.setStyleSheet("background-color : red; color : blue;")
        self.label_block.setObjectName("label_block")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "CV Version Control"))
        self.label.setText(_translate("Form", "<html><head/><body><img src=\"icv.png\" width=200/></body></html>"))
        self.label_doc_img.setText(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<br/><img src=\"iword.png\" width=100/></body></html>"))
        self.label_excel_img.setText(_translate("Form", "<html><head/><body><img src=\"iexcel.png\" width=100/></body></html>"))
        self.button_word2excel.setText(_translate("Form", "↑\n"
"↑\n"
"↑\n"
"↑"))
        self.button_excel2word.setText(_translate("Form", "↓\n"
"↓\n"
"↓\n"
"↓"))
        self.button_excel2cv.setText(_translate("Form", "→ → → →"))
        self.check_ispdf.setText(_translate("Form", "PDF"))
        self.excelpath_label.setText(_translate("Form", "Excel Path"))
        self.cvtype_label.setText(_translate("Form", "CV Type"))
        self.templeteword_path_label.setText(_translate("Form", "Templete Word"))
        self.wordpath_label.setText(_translate("Form", "Word Path"))
        self.DebugBrowser.setHtml(_translate("Form", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
"<p style=\" margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Debug Area</p></body></html>"))
        self.checkBox_LangChinese.setText(_translate("Form", "中文"))
        self.label_editdoc.setText(_translate("Form", "Modified Time"))
        self.label_allinone.setText(_translate("Form", "Modified Time"))
        self.label_4.setText(_translate("Form", "Labels"))
        self.label_5.setText(_translate("Form", "Debug"))
        self.checkBox_export_all_labels_cv.setText(_translate("Form", "导出所有Label的简历"))
        self.pushButton_openExcel.setText(_translate("Form", "Open"))
        self.pushButton_openDoc.setText(_translate("Form", "Open"))
        self.label_block.setText(_translate("Form", "<html><head/><body><p align=\"center\"><span style=\" font-size:16pt; font-weight:600; text-decoration: underline;\">关闭文件后才可以进行下一步操作</span></p></body></html>"))

