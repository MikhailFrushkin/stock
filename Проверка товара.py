import glob
import os
import random
from pathlib import Path
from sys import argv, executable

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QGraphicsPixmapItem, QGraphicsScene

from utils import check_stock


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setEnabled(True)
        MainWindow.setFixedSize(450, 431)
        MainWindow.setWindowIcon(QtGui.QIcon("images/icons/favicon.ico"))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 0, 441, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setGeometry(QtCore.QRect(10, 260, 300, 17))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.checkBox.setFont(font)
        self.checkBox.setObjectName("checkBox")
        self.checkBox_2 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_2.setGeometry(QtCore.QRect(10, 290, 300, 16))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.checkBox_2.setFont(font)
        self.checkBox_2.setObjectName("checkBox_2")
        self.checkBox_3 = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_3.setGeometry(QtCore.QRect(10, 320, 300, 17))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        self.checkBox_3.setFont(font)
        self.checkBox_3.setObjectName("checkBox_3")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 120, 449, 51))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 60, 161, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 190, 161, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(150, 360, 161, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(10, 30, 521, 31))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setItalic(False)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(10, 150, 445, 41))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setItalic(False)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")

        self.graphicsView = QtWidgets.QGraphicsView(self.centralwidget)
        self.graphicsView.setGeometry(QtCore.QRect(-10, -9, 461, 451))
        self.graphicsView.setAutoFillBackground(True)
        self.graphicsView.setObjectName("graphicsView")

        self.graphicsView.raise_()
        self.label.raise_()
        self.checkBox.raise_()
        self.checkBox_2.raise_()
        self.checkBox_3.raise_()
        self.label_2.raise_()
        self.pushButton.raise_()
        self.pushButton_2.raise_()
        self.pushButton_3.raise_()
        self.label_3.raise_()
        self.label_4.raise_()

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Проверка остатков товара"))
        self.label.setText(_translate("MainWindow", "Файл отстаков с 6.1 Складские лоты"))
        self.checkBox.setText(_translate("MainWindow", "Сформировать минусовые RDiff"))
        self.checkBox_2.setText(_translate("MainWindow", "Сформировать плюсовые RDiff"))
        self.checkBox_3.setText(_translate("MainWindow", "Сформировать файлы пст мин.витрины"))
        self.label_2.setText(_translate("MainWindow", "Файл Мин.витрины"))
        self.pushButton.setText(_translate("MainWindow", "Выбрать файл"))
        self.pushButton_2.setText(_translate("MainWindow", "Выбрать файл"))
        self.pushButton_3.setText(_translate("MainWindow", "Выполнить"))
        self.label_3.setText(_translate("MainWindow", "Файл не выбран"))
        self.label_4.setText(_translate("MainWindow", "Файл не выбран"))


class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()

        self.current_dir = Path.cwd()
        self.setupUi(self)
        self.pushButton.clicked.connect(self.evt_btn_open_file_clicked)
        self.pushButton_2.clicked.connect(self.evt_btn_open_file_clicked2)
        self.pushButton_3.clicked.connect(self.evt_btn_clicked)
        files_pic = glob.glob('images/*.*')
        if len(files_pic) > 0:
            pic = QGraphicsPixmapItem()
            pic.setPixmap(QPixmap(random.choice(files_pic)).scaled(459, 431))
            self.graphicsView.setScene(QGraphicsScene())
            self.graphicsView.scene().addItem(pic)

    def evt_btn_open_file_clicked(self):
        res = QFileDialog.getOpenFileName(self, 'Открыть файл', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if res[0] != '':
            self.label_3.setText(res[0])

    def evt_btn_open_file_clicked2(self):
        res = QFileDialog.getOpenFileName(self, 'Открыть файл', f'{self.current_dir}', 'Лист XLSX (*.xlsx)')
        if res[0] != '':
            self.label_4.setText(res[0])

    def evt_btn_clicked(self):
        flag = True
        min_vitrina = False
        plus = False
        minus = False
        name_file_min_vitrina = None

        if self.label_3.text() != 'Файл не выбран':
            file_name_stock = self.label_3.text()
            if self.checkBox.checkState() == 2:
                minus = True
            if self.checkBox_2.checkState() == 2:
                plus = True
            if self.checkBox_3.checkState() == 2:
                if self.label_4.text() != 'Файл не выбран':
                    min_vitrina = True
                    name_file_min_vitrina = self.label_4.text()
                else:
                    flag = False
                    QMessageBox.critical(self, 'Не выбран файл мин.витрины', 'Не выбран файл мин.витрины')
            if flag:
                QMessageBox.information(self, 'Завершено', '{}'.format(check_stock(self=self,
                                                                                   file_path=file_name_stock,
                                                                                   min_vitrina=min_vitrina,
                                                                                   plus=plus, minus=minus,
                                                                                   name_file_min_vitrina=
                                                                                   name_file_min_vitrina)))
        else:
            QMessageBox.critical(self, 'Не выбран файл с остатками', 'Не выбран файл с остатками')

    def restart1(self):
        os.execl(executable, os.path.abspath(__file__), *argv)


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
