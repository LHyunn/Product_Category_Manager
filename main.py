import os
import pandas as pd
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QTableWidgetItem, QAbstractItemView
from PyQt5.QtGui import QPixmap, QIcon
from qt_material import apply_stylesheet

global backup_company, backup_product, backup_main, backup_main, backup_mid, backup_sub, backup_opt1, backup_opt2
global company, product, main, mid, sub, opt1, opt2
global code_log, code_name_log, code_log_reset, list_xlsx_name, list_option

os.environ['QT_API'] = 'pyqt5'

os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"


class Ui_mainWindow(object):
    def setupUi(self, mainWindow):

        mainWindow.setObjectName("mainWindow")
        mainWindow.resize(940, 721)
        mainWindow.setMinimumSize(QtCore.QSize(940, 721))
        mainWindow.setMaximumSize(QtCore.QSize(940, 721))
        self.statusbar = QtWidgets.QStatusBar(mainWindow)
        self.statusbar = QtWidgets.QStatusBar(mainWindow)


        self.centralwidget = QtWidgets.QWidget(mainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalFrame = QtWidgets.QFrame(self.centralwidget)
        self.horizontalFrame.setGeometry(QtCore.QRect(0, 0, 940, 721))
        self.horizontalFrame.setObjectName("horizontalFrame")
        self.tableWidget99 = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget99.setRowCount(1)
        self.tableWidget99.setColumnCount(1)  # 이부분 tableWidget의 none값의 타입을 모르겠고 인터넷에서도 관련 문서를 찾을 수가 없어서 이렇게 나둠. 비어있는지 비교하기 위한 값.
        self.tableWidget99.hide()

        self.Frame_1 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_1.setGeometry(QtCore.QRect(10, 60, 321, 211))
        self.Frame_1.setObjectName("Frame_1")
        self.label_1 = QtWidgets.QLabel(self.Frame_1)
        self.label_1.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_1.setObjectName("label_1")
        self.checkBox = QtWidgets.QCheckBox(self.Frame_1)
        self.checkBox.setGeometry(QtCore.QRect(220, 10, 100, 20))
        self.checkBox.setObjectName("checkBox")
        # self.checkBox.toggle()
        self.checkBox.hide()
        self.pushButton_11 = QtWidgets.QPushButton(self.Frame_1)
        self.pushButton_11.setGeometry(QtCore.QRect(185, 5, 40, 30))
        self.pushButton_11.setObjectName("pushButton_2")

        self.pushButton_11.setCheckable(True)


        self.pushButton_12 = QtWidgets.QPushButton(self.Frame_1)
        self.pushButton_12.setGeometry(QtCore.QRect(230,5, 40, 30))
        self.pushButton_12.setObjectName("pushButton_4")

        self.pushButton_12.setDisabled(True)



        self.pushButton_13 = QtWidgets.QPushButton(self.Frame_1)
        self.pushButton_13.setGeometry(QtCore.QRect(275,5, 40, 30))
        self.pushButton_13.setObjectName("pushButton_5")


        self.pushButton_13.setDisabled(True)

        self.tableWidget = QtWidgets.QTableWidget(self.Frame_1)
        self.tableWidget.setGeometry(QtCore.QRect(0, 40, 321, 171))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(2)


        self.tableWidget.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget.setAlternatingRowColors(True)
        self.tableWidget.resizeColumnToContents(True)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)






        self.Frame_2 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_2.setGeometry(QtCore.QRect(10, 275, 321, 211))
        self.Frame_2.setObjectName("Frame_2")
        self.label_2 = QtWidgets.QLabel(self.Frame_2)
        self.label_2.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_2.setObjectName("label_2")
        self.checkBox_2 = QtWidgets.QCheckBox(self.Frame_2)
        self.checkBox_2.setGeometry(QtCore.QRect(100, 10, 100, 20))
        self.checkBox_2.setObjectName("checkBox_2")
        self.checkBox_2.toggle()
        self.pushButton_21 = QtWidgets.QPushButton(self.Frame_2)
        self.pushButton_21.setGeometry(QtCore.QRect(185,5,40,30))
        self.pushButton_21.setObjectName("pushButton_2")

        self.pushButton_21.setCheckable(True)


        self.pushButton_22 = QtWidgets.QPushButton(self.Frame_2)
        self.pushButton_22.setGeometry(QtCore.QRect(230,5, 40, 30))
        self.pushButton_22.setObjectName("pushButton_4")

        self.pushButton_22.setDisabled(True)

        self.pushButton_23 = QtWidgets.QPushButton(self.Frame_2)
        self.pushButton_23.setGeometry(QtCore.QRect(275,5, 40, 30))
        self.pushButton_23.setObjectName("pushButton_5")

        self.pushButton_23.setDisabled(True)



        self.tableWidget_2 = QtWidgets.QTableWidget(self.Frame_2)
        self.tableWidget_2.setGeometry(QtCore.QRect(0, 40, 321, 171))
        self.tableWidget_2.setObjectName("tableWidget_2")
        self.tableWidget_2.setColumnCount(2)
        self.tableWidget_2.setRowCount(2)
        self.tableWidget_2.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_2.setAlternatingRowColors(True)
        self.tableWidget_2.resizeColumnToContents(True)
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_2.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.Frame_3 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_3.setGeometry(QtCore.QRect(10, 490, 321, 211))
        self.Frame_3.setObjectName("Frame_3")
        self.label_3 = QtWidgets.QLabel(self.Frame_3)
        self.label_3.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_3.setObjectName("label_3")
        self.checkBox_3 = QtWidgets.QCheckBox(self.Frame_3)
        self.checkBox_3.setGeometry(QtCore.QRect(100, 10, 100, 20))
        self.checkBox_3.setObjectName("checkBox_3")
        self.checkBox_3.toggle()
        self.pushButton_31 = QtWidgets.QPushButton(self.Frame_3)
        self.pushButton_31.setGeometry(QtCore.QRect(185,5,40,30))
        self.pushButton_31.setObjectName("pushButton_2")




        self.pushButton_32 = QtWidgets.QPushButton(self.Frame_3)
        self.pushButton_32.setGeometry(QtCore.QRect(230,5, 40, 30))
        self.pushButton_32.setObjectName("pushButton_4")

        self.pushButton_33 = QtWidgets.QPushButton(self.Frame_3)
        self.pushButton_33.setGeometry(QtCore.QRect(275,5, 40, 30))
        self.pushButton_33.setObjectName("pushButton_5")

        self.pushButton_31.setCheckable(True)
        self.pushButton_32.setDisabled(True)
        self.pushButton_33.setDisabled(True)


        self.tableWidget_3 = QtWidgets.QTableWidget(self.Frame_3)
        self.tableWidget_3.setGeometry(QtCore.QRect(0, 40,  321, 171))
        self.tableWidget_3.setObjectName("tableWidget_3")
        self.tableWidget_3.setColumnCount(2)
        self.tableWidget_3.setRowCount(1)

        self.tableWidget_3.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_3.setAlternatingRowColors(True)
        self.tableWidget_3.resizeColumnToContents(True)
        self.tableWidget_3.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_3.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.Frame_4 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_4.setGeometry(QtCore.QRect(340, 60, 291, 321))
        self.Frame_4.setObjectName("Frame_4")
        self.label_4 = QtWidgets.QLabel(self.Frame_4)
        self.label_4.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_4.setObjectName("label_4")
        self.checkBox_4 = QtWidgets.QCheckBox(self.Frame_4)
        self.checkBox_4.setGeometry(QtCore.QRect(250, 5, 87, 20))
        self.checkBox_4.setObjectName("checkBox_4")
        # self.checkBox_4.toggle()
        self.checkBox_4.hide()
        self.pushButton_41 = QtWidgets.QPushButton(self.Frame_4)
        self.pushButton_41.setGeometry(QtCore.QRect(155, 5, 40, 30))
        self.pushButton_41.setObjectName("pushButton_2")




        self.pushButton_42 = QtWidgets.QPushButton(self.Frame_4)
        self.pushButton_42.setGeometry(QtCore.QRect(200, 5, 40, 30))
        self.pushButton_42.setObjectName("pushButton_4")

        self.pushButton_43 = QtWidgets.QPushButton(self.Frame_4)
        self.pushButton_43.setGeometry(QtCore.QRect(245, 5, 40, 30))
        self.pushButton_43.setObjectName("pushButton_5")

        self.pushButton_41.setCheckable(True)
        self.pushButton_42.setDisabled(True)
        self.pushButton_43.setDisabled(True)


        self.tableWidget_4 = QtWidgets.QTableWidget(self.Frame_4)
        self.tableWidget_4.setGeometry(QtCore.QRect(0, 40, 291, 281))
        self.tableWidget_4.setObjectName("tableWidget_4")
        self.tableWidget_4.setColumnCount(2)
        self.tableWidget_4.setRowCount(1)

        self.tableWidget_4.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_4.setAlternatingRowColors(True)
        self.tableWidget_4.resizeColumnToContents(True)
        self.tableWidget_4.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_4.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.Frame_5 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_5.setGeometry(QtCore.QRect(340, 390,291, 311))
        self.Frame_5.setObjectName("Frame_5")
        self.label_5 = QtWidgets.QLabel(self.Frame_5)
        self.label_5.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_5.setObjectName("label_5")
        self.checkBox_5 = QtWidgets.QCheckBox(self.Frame_5)
        self.checkBox_5.setGeometry(QtCore.QRect(250, 5, 87, 20))
        self.checkBox_5.setObjectName("checkBox_5")
        # self.checkBox_5.toggle()
        self.checkBox_5.hide()
        self.pushButton_51 = QtWidgets.QPushButton(self.Frame_5)
        self.pushButton_51.setGeometry(QtCore.QRect(155, 5, 40, 30))
        self.pushButton_51.setObjectName("pushButton_2")

        self.pushButton_52 = QtWidgets.QPushButton(self.Frame_5)
        self.pushButton_52.setGeometry(QtCore.QRect(200, 5, 40, 30))
        self.pushButton_52.setObjectName("pushButton_4")

        self.pushButton_53 = QtWidgets.QPushButton(self.Frame_5)
        self.pushButton_53.setGeometry(QtCore.QRect(245, 5, 40, 30))
        self.pushButton_53.setObjectName("pushButton_5")

        self.pushButton_51.setCheckable(True)
        self.pushButton_52.setDisabled(True)
        self.pushButton_53.setDisabled(True)

        self.tableWidget_5 = QtWidgets.QTableWidget(self.Frame_5)
        self.tableWidget_5.setGeometry(QtCore.QRect(0, 40, 291, 271))
        self.tableWidget_5.setObjectName("tableWidget_5")
        self.tableWidget_5.setColumnCount(2)
        self.tableWidget_5.setRowCount(1)



        self.tableWidget_5.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_5.setAlternatingRowColors(True)
        self.tableWidget_5.resizeColumnToContents(True)
        self.tableWidget_5.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_5.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.Frame_6 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_6.setGeometry(QtCore.QRect(640, 60, 291, 321))
        self.Frame_6.setObjectName("Frame_6")
        self.label_6 = QtWidgets.QLabel(self.Frame_6)
        self.label_6.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_6.setObjectName("label_6")
        self.checkBox_6 = QtWidgets.QCheckBox(self.Frame_6)
        self.checkBox_6.setGeometry(QtCore.QRect(250, 5, 87, 20))
        self.checkBox_6.setObjectName("checkBox_6")
        # self.checkBox_6.toggle()
        self.checkBox_6.hide()

        self.pushButton_61 = QtWidgets.QPushButton(self.Frame_6)
        self.pushButton_61.setGeometry(QtCore.QRect(155, 5, 40, 30))
        self.pushButton_61.setObjectName("pushButton_2")

        self.pushButton_61.setCheckable(True)


        self.pushButton_62 = QtWidgets.QPushButton(self.Frame_6)
        self.pushButton_62.setGeometry(QtCore.QRect(200, 5, 40, 30))
        self.pushButton_62.setObjectName("pushButton_4")

        self.pushButton_63 = QtWidgets.QPushButton(self.Frame_6)
        self.pushButton_63.setGeometry(QtCore.QRect(245, 5, 40, 30))
        self.pushButton_63.setObjectName("pushButton_5")

        self.pushButton_61.setCheckable(True)
        self.pushButton_62.setDisabled(True)
        self.pushButton_63.setDisabled(True)


        self.tableWidget_6 = QtWidgets.QTableWidget(self.Frame_6)
        self.tableWidget_6.setGeometry(QtCore.QRect(0, 40, 291, 281))
        self.tableWidget_6.setObjectName("tableWidget_6")
        self.tableWidget_6.setColumnCount(2)
        self.tableWidget_6.setRowCount(1)

        self.tableWidget_6.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_6.setAlternatingRowColors(True)
        self.tableWidget_6.resizeColumnToContents(True)
        self.tableWidget_6.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_6.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.Frame_7 = QtWidgets.QFrame(self.horizontalFrame)
        self.Frame_7.setGeometry(QtCore.QRect(640, 390, 291, 311))
        self.Frame_7.setObjectName("Frame_7")
        self.label_7 = QtWidgets.QLabel(self.Frame_7)
        self.label_7.setGeometry(QtCore.QRect(10, 10, 60, 16))
        self.label_7.setObjectName("label_7")
        self.checkBox_7 = QtWidgets.QCheckBox(self.Frame_7)
        self.checkBox_7.setGeometry(QtCore.QRect(250, 5, 87, 20))
        self.checkBox_7.setObjectName("checkBox_7")
        # self.checkBox_7.toggle()
        self.checkBox_7.hide()
        self.pushButton_71 = QtWidgets.QPushButton(self.Frame_7)
        self.pushButton_71.setGeometry(QtCore.QRect(155, 5, 40, 30))
        self.pushButton_71.setObjectName("pushButton_2")

        self.pushButton_71.setCheckable(True)


        self.pushButton_72 = QtWidgets.QPushButton(self.Frame_7)
        self.pushButton_72.setGeometry(QtCore.QRect(200, 5, 40, 30))
        self.pushButton_72.setObjectName("pushButton_4")

        self.pushButton_73 = QtWidgets.QPushButton(self.Frame_7)
        self.pushButton_73.setGeometry(QtCore.QRect(245, 5, 40, 30))
        self.pushButton_73.setObjectName("pushButton_5")

        self.pushButton_71.setCheckable(True)
        self.pushButton_72.setDisabled(True)
        self.pushButton_73.setDisabled(True)

        self.tableWidget_7 = QtWidgets.QTableWidget(self.Frame_7)
        self.tableWidget_7.setGeometry(QtCore.QRect(0, 40, 291, 271))
        self.tableWidget_7.setObjectName("tableWidget_7")
        self.tableWidget_7.setColumnCount(2)
        self.tableWidget_7.setRowCount(1)

        self.tableWidget_7.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.tableWidget_7.setAlternatingRowColors(True)
        self.tableWidget_7.resizeColumnToContents(True)
        self.tableWidget_7.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_7.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)


        self.label_8 = QtWidgets.QLabel(self.horizontalFrame)
        self.label_8.setGeometry(QtCore.QRect(10, 10, 81, 16))
        self.label_8.setObjectName("label_8")
        self.label_9 = QtWidgets.QLabel(self.horizontalFrame)
        self.label_9.setGeometry(QtCore.QRect(180, 10, 81, 16))
        self.label_9.setObjectName("label_9")

        self.comboBox = QtWidgets.QComboBox(self.horizontalFrame)
        self.comboBox.setGeometry(QtCore.QRect(700, 15, 101, 31))
        self.comboBox.setObjectName("comboBox")

        self.textEdit = QtWidgets.QTextEdit(self.horizontalFrame)
        self.textEdit.setGeometry(QtCore.QRect(10, 30, 161, 20))
        self.textEdit.setObjectName("textEdit")

        self.textEdit_2 = QtWidgets.QTextEdit(self.horizontalFrame)
        self.textEdit_2.setGeometry(QtCore.QRect(180, 30, 251, 20))
        self.textEdit_2.setObjectName("textEdit_2")


        self.pushButton = QtWidgets.QPushButton(self.horizontalFrame)
        self.pushButton.setGeometry(QtCore.QRect(810, 15, 121, 31))
        self.pushButton.setObjectName("pushButton")



        self.pushButton_3 = QtWidgets.QPushButton(self.horizontalFrame)
        self.pushButton_3.setGeometry(QtCore.QRect(770, 30, 101, 31))
        self.pushButton_3.setObjectName("pushButton_3")

        self.pushButton_3.setCheckable(True)

        self.pushButton_3.hide()

        mainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(mainWindow)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)


        self.label_logo = QtWidgets.QLabel(self.horizontalFrame)
        self.label_logo.setGeometry(435, 5, 81, 43)
        self.label_logo.setObjectName("label_logo")
        ######################################################################
        self.display_items_company()
        self.tableWidget.cellClicked.connect(self.display_items_product)
        self.tableWidget_2.cellClicked.connect(self.display_items_main)
        self.tableWidget_3.cellClicked.connect(self.display_items_mid)
        self.tableWidget_4.cellClicked.connect(self.display_items_sub)
        self.tableWidget_5.cellClicked.connect(self.display_items_opt1)
        self.tableWidget_6.cellClicked.connect(self.display_items_opt2)
        self.tableWidget_7.cellClicked.connect(self.opt2_select)
        self.pushButton_11.clicked.connect(self.modify_company)
        self.pushButton_12.clicked.connect(self.delete_row_company)
        self.pushButton_13.clicked.connect(self.save_company)
        self.pushButton_21.clicked.connect(self.modify_product)
        self.pushButton_22.clicked.connect(self.delete_row_product)
        self.pushButton_23.clicked.connect(self.save_product)
        self.pushButton_31.clicked.connect(self.modify_main)
        self.pushButton_32.clicked.connect(self.delete_row_main)
        self.pushButton_33.clicked.connect(self.save_main)
        self.pushButton_41.clicked.connect(self.modify_mid)
        self.pushButton_42.clicked.connect(self.delete_row_mid)
        self.pushButton_43.clicked.connect(self.save_mid)
        self.pushButton_51.clicked.connect(self.modify_sub)
        self.pushButton_52.clicked.connect(self.delete_row_sub)
        self.pushButton_53.clicked.connect(self.save_sub)
        self.pushButton_61.clicked.connect(self.modify_opt1)
        self.pushButton_62.clicked.connect(self.delete_row_opt1)
        self.pushButton_63.clicked.connect(self.save_opt1)
        self.pushButton_71.clicked.connect(self.modify_opt2)
        self.pushButton_72.clicked.connect(self.delete_row_opt2)
        self.pushButton_73.clicked.connect(self.save_opt2)
        self.tableWidget.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget.setAlternatingRowColors(True)

        self.tableWidget_2.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_2.setAlternatingRowColors(True)

        self.tableWidget_3.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_3.setAlternatingRowColors(True)

        self.tableWidget_4.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_4.setAlternatingRowColors(True)

        self.tableWidget_5.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_5.setAlternatingRowColors(True)

        self.tableWidget_6.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_6.setAlternatingRowColors(True)

        self.tableWidget_7.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tableWidget_7.setAlternatingRowColors(True)

        qPixmapVar = QPixmap()
        qPixmapVar.load("logo.png")
        self.label_logo.setPixmap(qPixmapVar)
        mainWindow.setWindowIcon(QIcon('icon.ico'))
    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "품목분류코드 생성기"))
        self.label_1.setText(_translate("mainWindow", "회사"))
        self.label_2.setText(_translate("mainWindow", "품목군"))
        self.label_3.setText(_translate("mainWindow", "대분류"))
        self.label_4.setText(_translate("mainWindow", "중분류"))
        self.label_5.setText(_translate("mainWindow", "소분류"))
        self.label_6.setText(_translate("mainWindow", "옵션1"))
        self.label_7.setText(_translate("mainWindow", "옵션2"))
        self.label_8.setText(_translate("mainWindow", "품목 코드"))
        self.label_9.setText(_translate("mainWindow", "품목명 코드"))
        self.pushButton.setText(_translate("mainWindow", "엑셀 파일 열기"))
        self.pushButton_3.setText(_translate("mainWindow", "수정 모드 꺼짐"))
        self.pushButton_11.setText(_translate("mainWindow", "수정"))
        self.pushButton_12.setText(_translate("mainWindow", "삭제"))
        self.pushButton_13.setText(_translate("mainWindow", "저장"))
        self.pushButton_21.setText(_translate("mainWindow", "수정"))
        self.pushButton_22.setText(_translate("mainWindow", "삭제"))
        self.pushButton_23.setText(_translate("mainWindow", "저장"))
        self.pushButton_31.setText(_translate("mainWindow", "수정"))
        self.pushButton_32.setText(_translate("mainWindow", "삭제"))
        self.pushButton_33.setText(_translate("mainWindow", "저장"))
        self.pushButton_41.setText(_translate("mainWindow", "수정"))
        self.pushButton_42.setText(_translate("mainWindow", "삭제"))
        self.pushButton_43.setText(_translate("mainWindow", "저장"))
        self.pushButton_51.setText(_translate("mainWindow", "수정"))
        self.pushButton_52.setText(_translate("mainWindow", "삭제"))
        self.pushButton_53.setText(_translate("mainWindow", "저장"))
        self.pushButton_61.setText(_translate("mainWindow", "수정"))
        self.pushButton_62.setText(_translate("mainWindow", "삭제"))
        self.pushButton_63.setText(_translate("mainWindow", "저장"))
        self.pushButton_71.setText(_translate("mainWindow", "수정"))
        self.pushButton_72.setText(_translate("mainWindow", "삭제"))
        self.pushButton_73.setText(_translate("mainWindow", "저장"))

        self.comboBox.setItemText(0, _translate("MainWindow", "전체"))
        self.comboBox.setItemText(1, _translate("MainWindow", "회사"))
        self.comboBox.setItemText(2, _translate("MainWindow", "품목군"))
        self.comboBox.setItemText(3, _translate("MainWindow", "대분류"))
        self.comboBox.setItemText(4, _translate("MainWindow", "중분류"))
        self.comboBox.setItemText(5, _translate("MainWindow", "소분류"))
        self.comboBox.setItemText(6, _translate("MainWindow", "옵션1"))
        self.comboBox.setItemText(7, _translate("MainWindow", "옵션2"))
        self.checkBox.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_4.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_5.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_2.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_3.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_7.setText(_translate("mainWindow", "품목명 표시"))
        self.checkBox_6.setText(_translate("mainWindow", "품목명 표시"))

    def display_items_company(self):
        global company
        company = pd.read_excel("Companies.xlsx", index_col=0)
        self.cleaner_combobox_tableWidget(1)
        self.tableWidget.setRowCount(len(company.index.tolist()))
        for i in range(0, len(company.index.tolist()), 1):
            for j in range(0, 2, 1):
                if str(company.iat[i, j]) != 'nan':
                    item = QTableWidgetItem(str(company.iat[i, j]))
                    self.tableWidget.setItem(i, j, item)

    def display_items_product(self):
        global product, code_log, code_name_log
        if not True == self.pushButton_11.isChecked():
            row = self.tableWidget.currentIndex().row()
            if self.tableWidget.item(row, 1) != self.tableWidget99.item(1, 1):
                self.cleaner_combobox_tableWidget(2)
                self.code_name_logging(1, self.tableWidget.item(row, 1).text())
                self.code_logging(1, self.tableWidget.item(row, 0).text())
                product = pd.read_excel("Products.xlsx", index_col=0)
                product_selected = product.loc[product['회사']== code_log[0]]
                self.tableWidget_2.setRowCount(len(product_selected.index.tolist()))
                for i in range(0, len(product_selected.index.tolist()), 1):
                    for j in range(0,2, 1):
                        if str(product_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(product_selected.iat[i, j]))
                            self.tableWidget_2.setItem(i, j, item)
            else:
                self.cleaner_combobox_tableWidget(2)


    def display_items_main(self):
        global main, code_log, code_name_log
        if not True == self.pushButton_21.isChecked():
            row = self.tableWidget_2.currentIndex().row()
            if self.tableWidget_2.item(row, 1) != self.tableWidget99.item(1, 1):
                self.cleaner_combobox_tableWidget(3)
                self.code_name_logging(2, self.tableWidget_2.item(row, 1).text())
                self.code_logging(2, self.tableWidget_2.item(row, 0).text())
                main = pd.read_excel("MainCategory.xlsx", index_col=0)
                main_selected = main.loc[main['회사']== code_log[0]]
                main_selected = main_selected.loc[main_selected['품목군코드'] == code_log[1]]
                self.tableWidget_3.setRowCount(len(main_selected.index.tolist()))
                for i in range(0, len(main_selected.index.tolist()), 1):
                    for j in range(0, 2, 1):
                        if str(main_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(main_selected.iat[i, j]))
                            self.tableWidget_3.setItem(i, j, item)
            else:
                self.cleaner_combobox_tableWidget(3)

    def display_items_mid(self):
        global mid, code_log, code_name_log
        if not True == self.pushButton_31.isChecked():
            row = self.tableWidget_3.currentIndex().row()
            if self.tableWidget_3.item(row, 1) != self.tableWidget99.item(1, 1):
                self.cleaner_combobox_tableWidget(4)
                self.code_name_logging(3, self.tableWidget_3.item(row, 1).text())
                self.code_logging(3, self.tableWidget_3.item(row, 0).text())
                mid = pd.read_excel("MiddleCategory.xlsx", index_col=0)
                mid_selected = mid.loc[mid['회사']== code_log[0]]
                mid_selected = mid_selected.loc[mid_selected['품목군코드'] == code_log[1]]
                mid_selected = mid_selected.loc[mid_selected['대분류코드'] == code_log[2]]
                self.tableWidget_4.setRowCount(len(mid_selected.index.tolist()))
                for i in range(0, len(mid_selected.index.tolist()), 1):
                    for j in range(0, 2, 1):
                        if str(mid_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(mid_selected.iat[i, j]))
                            self.tableWidget_4.setItem(i, j, item)
            else:
                self.cleaner_combobox_tableWidget(4)

    def display_items_sub(self):
        global sub, code_log, code_name_log
        if not True == self.pushButton_41.isChecked():
            row = self.tableWidget_4.currentIndex().row()
            if self.tableWidget_4.item(row, 1) != self.tableWidget99.item(1, 1):
                self.code_name_logging(4, self.tableWidget_4.item(row, 1).text())
                self.code_logging(4, self.tableWidget_4.item(row, 0).text())
                sub = pd.read_excel("Subcategory.xlsx", index_col=0)
                self.cleaner_combobox_tableWidget(5)
                sub_selected = sub.loc[sub['회사']== code_log[0]]
                sub_selected = sub_selected.loc[sub_selected['품목군코드'] == code_log[1]]
                sub_selected = sub_selected.loc[sub_selected['대분류코드'] == code_log[2]]
                sub_selected = sub_selected.loc[sub_selected['중분류코드'] == code_log[3]]
                self.tableWidget_5.setRowCount(len(sub_selected.index.tolist()))
                for i in range(0, len(sub_selected.index.tolist()), 1):
                    for j in range(0, 2, 1):
                        if str(sub_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(sub_selected.iat[i, j]))
                            self.tableWidget_5.setItem(i, j, item)
            else:
                self.cleaner_combobox_tableWidget(5)

    def display_items_opt1(self):
        global opt1, code_log, code_name_log
        if not True == self.pushButton_51.isChecked():
            row = self.tableWidget_5.currentIndex().row()
            if self.tableWidget_5.item(row, 1) != self.tableWidget99.item(1, 1):
                self.code_name_logging(5, self.tableWidget_5.item(row, 1).text())
                self.code_logging(5, self.tableWidget_5.item(row, 0).text())
                opt1 = pd.read_excel("Options1.xlsx", index_col=0)
                self.cleaner_combobox_tableWidget(6)
                opt1_selected = opt1.loc[opt1['회사']== code_log[0]]
                opt1_selected = opt1_selected.loc[opt1_selected['품목군코드'] == code_log[1]]
                opt1_selected = opt1_selected.loc[opt1_selected['대분류코드'] == code_log[2]]
                opt1_selected = opt1_selected.loc[opt1_selected['중분류코드'] == code_log[3]]
                opt1_selected = opt1_selected.loc[opt1_selected['소분류코드'] == code_log[4]]
                self.tableWidget_6.setRowCount(len(opt1_selected.index.tolist()))
                for i in range(0, len(opt1_selected.index.tolist()), 1):
                    for j in range(0, 2, 1):
                        if str(opt1_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(opt1_selected.iat[i, j]))
                            self.tableWidget_6.setItem(i, j, item)
            else:
                self.cleaner_combobox_tableWidget(6)

    def display_items_opt2(self):
        global opt2, code_log, code_name_log
        if not True == self.pushButton_61.isChecked():
            row = self.tableWidget_6.currentIndex().row()
            if self.tableWidget_6.item(row, 1) != self.tableWidget99.item(1, 1):
                self.code_name_logging(6,self.tableWidget_6.item(row, 1).text())
                self.code_logging(6,self.tableWidget_6.item(row, 0).text())
                opt2 = pd.read_excel("Options2.xlsx", index_col=0)
                self.cleaner_combobox_tableWidget(7)
                opt2_selected = opt2.loc[opt2['회사'] == code_log[0]]
                opt2_selected = opt2_selected.loc[opt2_selected['품목군코드'] == code_log[1]]
                opt2_selected = opt2_selected.loc[opt2_selected['대분류코드'] == code_log[2]]
                opt2_selected = opt2_selected.loc[opt2_selected['중분류코드'] == code_log[3]]
                opt2_selected = opt2_selected.loc[opt2_selected['소분류코드'] == code_log[4]]
                opt2_selected = opt2_selected.loc[opt2_selected['옵션1'] == code_log[5]]
                self.tableWidget_7.setRowCount(len(opt2_selected.index.tolist()))
                for i in range(0, len(opt2_selected.index.tolist()), 1):
                    for j in range(0, 2, 1):
                        if str(opt2_selected.iat[i, j]) != 'nan':
                            item = QTableWidgetItem(str(opt2_selected.iat[i, j]))
                            self.tableWidget_7.setItem(i, j, item)

    def opt2_select(self):
        if not True == self.pushButton_71.isChecked():
            row = self.tableWidget_7.currentIndex().row()
            if self.tableWidget_7.item(row, 1) != self.tableWidget99.item(1, 1):
                self.code_name_logging(7,self.tableWidget_7.item(row, 1).text())
                self.code_logging(7,self.tableWidget_7.item(row, 0).text())

    def modify_company(self):
        global backup_company
        backup_company = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_11.isChecked():

            for i in range(self.tableWidget.rowCount()):
                if self.tableWidget.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_company.at[i, '코드'] = self.tableWidget.item(i, 0).text()
                    backup_company.at[i, '코드명'] = self.tableWidget.item(i, 1).text()
            print("백업된 테이블 : ")
            print(backup_company)
            self.tableWidget.setRowCount(self.tableWidget.rowCount()+10)
            self.pushButton_12.setEnabled(True)
            self.pushButton_13.setEnabled(True)
            self.tableWidget.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget.setRowCount(self.tableWidget.rowCount() - 10)
            self.pushButton_12.setDisabled(True)
            self.pushButton_13.setDisabled(True)
            self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget.clear()
            for i in range(0, len(backup_company.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_company.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_company.iat[i, j]))
                        self.tableWidget.setItem(i, j, item)
            self.display_items_company()

    def delete_row_company(self):
        self.tableWidget.removeRow(self.tableWidget.currentRow())

    def save_company(self):
        global company
        company = pd.concat([backup_company, company])
        company = pd.concat([backup_company, company])
        company = company.drop_duplicates(keep=False)
        print("company에서 백업 삭제")
        print(company)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget.rowCount()):
            if self.tableWidget.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget.item(i, 1).text()
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        company = pd.concat([new_df, company])
        company.to_excel("Companies.xlsx")
        company = pd.read_excel("Companies.xlsx", index_col = 0)
        self.tableWidget.setRowCount(self.tableWidget.rowCount() - 10)
        self.pushButton_11.toggle()
        self.pushButton_12.setDisabled(True)
        self.pushButton_13.setDisabled(True)
        self.tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_company()

    def modify_product(self):
        global backup_product
        backup_product = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_21.isChecked():

            for i in range(self.tableWidget_2.rowCount()):
                if self.tableWidget_2.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_product.at[i, '코드'] = self.tableWidget_2.item(i, 0).text()
                    backup_product.at[i, '코드명'] = self.tableWidget_2.item(i, 1).text()
                    backup_product.at[i, '회사'] = code_log[0]
            print("백업된 테이블 : ")
            print(backup_product)
            self.tableWidget_2.setRowCount(self.tableWidget_2.rowCount() + 10)
            self.pushButton_22.setEnabled(True)
            self.pushButton_23.setEnabled(True)
            self.tableWidget_2.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_2.setRowCount(self.tableWidget_2.rowCount() - 10)
            self.pushButton_22.setDisabled(True)
            self.pushButton_23.setDisabled(True)
            self.tableWidget_2.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_2.clear()
            for i in range(0, len(backup_product.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_product.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_product.iat[i, j]))
                        self.tableWidget_2.setItem(i, j, item)
            self.display_items_product()

    def delete_row_product(self):
        self.tableWidget_2.removeRow(self.tableWidget_2.currentRow())
    def save_product(self):
        global product
        product = pd.concat([backup_product, product])
        product = pd.concat([backup_product, product])
        product = product.drop_duplicates(keep=False)
        print("product에서 백업 삭제")
        print(product)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_2.rowCount()):
            if self.tableWidget_2.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_2.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_2.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        product = pd.concat([new_df, product])
        product.to_excel("Products.xlsx")
        product = pd.read_excel("Products.xlsx", index_col = 0)
        self.tableWidget_2.setRowCount(self.tableWidget_2.rowCount() - 10)
        self.pushButton_21.toggle()
        self.pushButton_22.setDisabled(True)
        self.pushButton_23.setDisabled(True)
        self.tableWidget_2.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_product()

    def modify_main(self):
        global backup_main
        backup_main = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_31.isChecked():

            for i in range(self.tableWidget_3.rowCount()):
                if self.tableWidget_3.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_main.at[i, '코드'] = self.tableWidget_3.item(i, 0).text()
                    backup_main.at[i, '코드명'] = self.tableWidget_3.item(i, 1).text()
                    backup_main.at[i, '회사'] = code_log[0]
                    backup_main.at[i, '품목군코드'] = code_log[1]
            print("백업된 테이블 : ")
            print(backup_main)
            self.tableWidget_3.setRowCount(self.tableWidget_3.rowCount() + 10)
            self.pushButton_32.setEnabled(True)
            self.pushButton_33.setEnabled(True)
            self.tableWidget_3.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_3.setRowCount(self.tableWidget_3.rowCount() - 10)
            self.pushButton_32.setDisabled(True)
            self.pushButton_33.setDisabled(True)
            self.tableWidget_3.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_3.clear()
            for i in range(0, len(backup_main.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_main.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_main.iat[i, j]))
                        self.tableWidget_3.setItem(i, j, item)
            self.display_items_main()

    def delete_row_main(self):
        self.tableWidget_3.removeRow(self.tableWidget_3.currentRow())

    def save_main(self):
        global main
        main = pd.concat([backup_main, main])
        main = pd.concat([backup_main, main])
        main = main.drop_duplicates(keep=False)
        print("main에서 백업 삭제")
        print(main)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_3.rowCount()):
            if self.tableWidget_3.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_3.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_3.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                    new_df.at[i, '품목군코드'] = code_log[1]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        main = pd.concat([new_df, main])
        main.to_excel("MainCategory.xlsx")
        main = pd.read_excel("MainCategory.xlsx", index_col=0)
        self.tableWidget_3.setRowCount(self.tableWidget_3.rowCount() - 10)
        self.pushButton_41.toggle()
        self.pushButton_42.setDisabled(True)
        self.pushButton_43.setDisabled(True)
        self.tableWidget_3.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_main()

    def modify_mid(self):
        global backup_mid
        backup_mid = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_41.isChecked():

            for i in range(self.tableWidget_4.rowCount()):
                if self.tableWidget_4.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_mid.at[i, '코드'] = self.tableWidget_4.item(i, 0).text()
                    backup_mid.at[i, '코드명'] = self.tableWidget_4.item(i, 1).text()
                    backup_mid.at[i, '회사'] = code_log[0]
                    backup_mid.at[i, '품목군코드'] = code_log[1]
                    backup_mid.at[i, '대분류코드'] = code_log[2]
            print("백업된 테이블 : ")
            print(backup_mid)
            self.tableWidget_4.setRowCount(self.tableWidget_4.rowCount() + 10)
            self.pushButton_42.setEnabled(True)
            self.pushButton_43.setEnabled(True)
            self.tableWidget_4.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_4.setRowCount(self.tableWidget_4.rowCount() - 10)
            self.pushButton_42.setDisabled(True)
            self.pushButton_43.setDisabled(True)
            self.tableWidget_4.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_4.clear()
            for i in range(0, len(backup_mid.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_mid.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_mid.iat[i, j]))
                        self.tableWidget_4.setItem(i, j, item)
            self.display_items_mid()

    def delete_row_mid(self):
        self.tableWidget_4.removeRow(self.tableWidget_4.currentRow())

    def save_mid(self):
        global mid
        mid = pd.concat([backup_mid, mid])
        mid = pd.concat([backup_mid, mid])
        mid = mid.drop_duplicates(keep=False)
        print("mid에서 백업 삭제")
        print(mid)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_4.rowCount()):
            if self.tableWidget_4.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_4.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_4.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                    new_df.at[i, '품목군코드'] = code_log[1]
                    new_df.at[i, '대분류코드'] = code_log[2]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        mid = pd.concat([new_df, mid])
        mid.to_excel("MiddleCategory.xlsx")
        mid = pd.read_excel("MiddleCategory.xlsx", index_col=0)
        self.tableWidget_4.setRowCount(self.tableWidget_4.rowCount() - 10)
        self.pushButton_41.toggle()
        self.pushButton_42.setDisabled(True)
        self.pushButton_43.setDisabled(True)
        self.tableWidget_4.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_mid()

    def modify_sub(self):
        global backup_sub
        backup_sub = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_51.isChecked():
            for i in range(self.tableWidget_5.rowCount()):
                if self.tableWidget_5.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_sub.at[i, '코드'] = self.tableWidget_5.item(i, 0).text()
                    backup_sub.at[i, '코드명'] = self.tableWidget_5.item(i, 1).text()
                    backup_sub.at[i, '회사'] = code_log[0]
                    backup_sub.at[i, '품목군코드'] = code_log[1]
                    backup_sub.at[i, '대분류코드'] = code_log[2]
                    backup_sub.at[i, '중분류코드'] = code_log[3]
            print("백업된 테이블 : ")
            print(backup_sub)
            self.tableWidget_5.setRowCount(self.tableWidget_5.rowCount() + 10)
            self.pushButton_52.setEnabled(True)
            self.pushButton_53.setEnabled(True)
            self.tableWidget_5.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_5.setRowCount(self.tableWidget_5.rowCount() - 10)
            self.pushButton_52.setDisabled(True)
            self.pushButton_53.setDisabled(True)
            self.tableWidget_5.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_5.clear()
            for i in range(0, len(backup_sub.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_sub.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_sub.iat[i, j]))
                        self.tableWidget_5.setItem(i, j, item)
            self.display_items_sub()

    def delete_row_sub(self):
        self.tableWidget_5.removeRow(self.tableWidget_5.currentRow())

    def save_sub(self):
        global sub
        sub = pd.concat([backup_sub, sub])
        sub = pd.concat([backup_sub, sub])
        sub = sub.drop_duplicates(keep=False)
        print("sub에서 백업 삭제")
        print(sub)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_5.rowCount()):
            if self.tableWidget_5.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_5.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_5.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                    new_df.at[i, '품목군코드'] = code_log[1]
                    new_df.at[i, '대분류코드'] = code_log[2]
                    new_df.at[i, '중분류코드'] = code_log[3]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        sub = pd.concat([new_df, sub])
        sub.to_excel("SubCategory.xlsx")
        sub = pd.read_excel("SubCategory.xlsx", index_col=0)
        self.tableWidget_5.setRowCount(self.tableWidget_5.rowCount() - 10)
        self.pushButton_51.toggle()
        self.pushButton_52.setDisabled(True)
        self.pushButton_53.setDisabled(True)
        self.tableWidget_5.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_sub()

    def modify_opt1(self):
        global backup_opt1

        backup_opt1 = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_61.isChecked():

            for i in range(self.tableWidget_6.rowCount()):
                if self.tableWidget_6.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_opt1.at[i, '코드'] = self.tableWidget_6.item(i, 0).text()
                    backup_opt1.at[i, '코드명'] = self.tableWidget_6.item(i, 1).text()
                    backup_opt1.at[i, '회사'] = code_log[0]
                    backup_opt1.at[i, '품목군코드'] = code_log[1]
                    backup_opt1.at[i, '대분류코드'] = code_log[2]
                    backup_opt1.at[i, '중분류코드'] = code_log[3]
                    backup_opt1.at[i, '소분류코드'] = code_log[4]
            print("백업된 테이블 : ")
            print(backup_opt1)
            self.tableWidget_6.setRowCount(self.tableWidget_6.rowCount() + 10)
            self.pushButton_62.setEnabled(True)
            self.pushButton_63.setEnabled(True)
            self.tableWidget_6.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_6.setRowCount(self.tableWidget_6.rowCount() - 10)
            self.pushButton_62.setDisabled(True)
            self.pushButton_63.setDisabled(True)
            self.tableWidget_6.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_6.clear()
            for i in range(0, len(backup_opt1.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_opt1.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_opt1.iat[i, j]))
                        self.tableWidget_6.setItem(i, j, item)
            self.display_items_opt1()

    def delete_row_opt1(self):
        self.tableWidget_6.removeRow(self.tableWidget_6.currentRow())

    def save_opt1(self):
        global opt1
        opt1 = pd.concat([backup_opt1, opt1])
        opt1 = pd.concat([backup_opt1, opt1])
        opt1 = opt1.drop_duplicates(keep=False)
        print("opt1에서 백업 삭제")
        print(opt1)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_6.rowCount()):
            if self.tableWidget_6.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_6.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_6.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                    new_df.at[i, '품목군코드'] = code_log[1]
                    new_df.at[i, '대분류코드'] = code_log[2]
                    new_df.at[i, '중분류코드'] = code_log[3]
                    new_df.at[i, '소분류코드'] = code_log[4]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        opt1 = pd.concat([new_df, opt1])
        opt1.to_excel("Options1.xlsx")
        opt1 = pd.read_excel("Options1.xlsx", index_col=0)
        self.tableWidget_6.setRowCount(self.tableWidget_6.rowCount() - 10)
        self.pushButton_61.toggle()
        self.pushButton_62.setDisabled(True)
        self.pushButton_63.setDisabled(True)
        self.tableWidget_6.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_opt1()

    def modify_opt2(self):
        global backup_opt2
        backup_opt2 = pd.read_excel("case.xlsx", index_col=0)
        if self.pushButton_71.isChecked():
            for i in range(self.tableWidget_7.rowCount()):
                if self.tableWidget_7.item(i, 0) != self.tableWidget99.item(1, 1):
                    backup_opt2.at[i, '코드'] = self.tableWidget_7.item(i, 0).text()
                    backup_opt2.at[i, '코드명'] = self.tableWidget_7.item(i, 1).text()
                    backup_opt2.at[i, '회사'] = code_log[0]
                    backup_opt2.at[i, '품목군코드'] = code_log[1]
                    backup_opt2.at[i, '대분류코드'] = code_log[2]
                    backup_opt2.at[i, '중분류코드'] = code_log[3]
                    backup_opt2.at[i, '소분류코드'] = code_log[4]
                    backup_opt2.at[i, '옵션1'] = code_log[5]
            print("백업된 테이블 : ")
            print(backup_opt2)
            self.tableWidget_7.setRowCount(self.tableWidget_7.rowCount() + 10)
            self.pushButton_72.setEnabled(True)
            self.pushButton_73.setEnabled(True)
            self.tableWidget_7.setEditTriggers(QAbstractItemView.DoubleClicked)
        else:
            self.tableWidget_7.setRowCount(self.tableWidget_7.rowCount() - 10)
            self.pushButton_72.setDisabled(True)
            self.pushButton_73.setDisabled(True)
            self.tableWidget_7.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tableWidget_7.clear()
            for i in range(0, len(backup_opt2.index.tolist()), 1):
                for j in range(0, 2):
                    if str(backup_opt2.iat[i, j]) != 'nan':
                        item = QTableWidgetItem(str(backup_opt2.iat[i, j]))
                        self.tableWidget_7.setItem(i, j, item)

            self.display_items_opt2()

    def delete_row_opt2(self):
        self.tableWidget_7.removeRow(self.tableWidget_7.currentRow())

    def save_opt2(self):
        global opt2
        opt2 = pd.concat([backup_opt2, opt2])
        opt2 = pd.concat([backup_opt2, opt2])
        opt2 = opt2.drop_duplicates(keep=False)
        print("opt2에서 백업 삭제")
        print(opt2)
        new_df = pd.read_excel("case.xlsx", index_col=0)
        for i in range(self.tableWidget_7.rowCount()):
            if self.tableWidget_7.item(i, 0) != self.tableWidget99.item(1, 1):
                try:
                    new_df.at[i, '코드'] = self.tableWidget_7.item(i, 0).text()
                    new_df.at[i, '코드명'] = self.tableWidget_7.item(i, 1).text()
                    new_df.at[i, '회사'] = code_log[0]
                    new_df.at[i, '품목군코드'] = code_log[1]
                    new_df.at[i, '대분류코드'] = code_log[2]
                    new_df.at[i, '중분류코드'] = code_log[3]
                    new_df.at[i, '소분류코드'] = code_log[4]
                    new_df.at[i, '옵션1'] = code_log[5]
                except:
                    new_df = new_df.drop([new_df.index[i]])
                    pass
        opt2 = pd.concat([new_df, opt2])
        opt2.to_excel("Options2.xlsx")
        opt2 = pd.read_excel("Options2.xlsx", index_col=0)
        self.tableWidget_7.setRowCount(self.tableWidget_7.rowCount() - 10)
        self.pushButton_71.toggle()
        self.pushButton_72.setDisabled(True)
        self.pushButton_73.setDisabled(True)
        self.tableWidget_7.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.display_items_opt2()


    def cleaner_combobox_tableWidget(self, where):
        if where < 7:
            self.tableWidget_7.clear()
            self.tableWidget_7.setRowCount(0)
            if where < 6:
                self.tableWidget_6.clear()
                self.tableWidget_6.setRowCount(0)
                if where < 5:
                    self.tableWidget_5.clear()
                    self.tableWidget_5.setRowCount(0)
                    if where < 4:
                        self.tableWidget_4.clear()
                        self.tableWidget_4.setRowCount(0)
                        if where < 3:
                            self.tableWidget_3.clear()
                            self.tableWidget_3.setRowCount(0)
                            if where < 2:
                                self.tableWidget_2.clear()
                                self.tableWidget_2.setRowCount(0)
                                if where < 1:
                                    self.tableWidget.clear()
                                    self.tableWidget.setRowCount(0)


    def code_logging(self, where, code):
        if where == 0:
            self.textEdit.clear()
            code_log[0] = code_log_reset[0]
            code_log[1] = code_log_reset[1]
            code_log[2] = code_log_reset[2]
            code_log[3] = code_log_reset[3]
            code_log[4] = code_log_reset[4]
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = ""
            self.textEdit.setText(result)
        elif where == 1:
            self.textEdit.clear()
            code_log[0] = code
            code_log[1] = code_log_reset[1]
            code_log[2] = code_log_reset[2]
            code_log[3] = code_log_reset[3]
            code_log[4] = code_log_reset[4]
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = code_log[0]
            self.textEdit.setText(result)
        elif where == 2:
            self.textEdit.clear()
            code_log[1] = code
            code_log[2] = code_log_reset[2]
            code_log[3] = code_log_reset[3]
            code_log[4] = code_log_reset[4]
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = code_log[0] + code_log[1]
            self.textEdit.setText(result)
        elif where == 3:
            self.textEdit.clear()
            code_log[2] = code
            code_log[3] = code_log_reset[3]
            code_log[4] = code_log_reset[4]
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = code_log[0] + code_log[1] + "-" + code_log[2]
            self.textEdit.setText(result)
        elif where == 4:
            self.textEdit.clear()
            code_log[3] = code
            code_log[4] = code_log_reset[4]
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = code_log[0] + code_log[1] + "-" + code_log[2] + code_log[3]
            self.textEdit.setText(result)
        elif where == 5:
            self.textEdit.clear()
            code_log[4] = code
            code_log[5] = code_log_reset[5]
            code_log[6] = code_log_reset[6]
            result = code_log[0] + code_log[1] + "-" + code_log[2] + code_log[3] + code_log[4]
            self.textEdit.setText(result)
        elif where == 6:
            self.textEdit.clear()
            code_log[5] = code
            code_log[6] = code_log_reset[6]
            result = code_log[0] + code_log[1] + "-" + code_log[2] + code_log[3] + code_log[4] + code_log[5]
            self.textEdit.setText(result)
        elif where == 7:
            self.textEdit.clear()
            code_log[6] = code
            result = code_log[0] + code_log[1] + "-" + code_log[2] + code_log[3] + code_log[4] + code_log[5] + code_log[
                6]
            self.textEdit.setText(result)

    def code_name_logging(self, where, codename):
        if where == 0:
            self.textEdit_2.clear()
            code_name_log[0] = code_log_reset[0]
            code_name_log[1] = code_log_reset[1]
            code_name_log[2] = code_log_reset[2]
            code_name_log[3] = code_log_reset[3]
            code_name_log[4] = code_log_reset[4]
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)

        elif where == 1:
            self.textEdit_2.clear()
            if self.checkBox.isChecked() == False:
                codename = ""
            code_name_log[0] = codename
            code_name_log[1] = code_log_reset[1]
            code_name_log[2] = code_log_reset[2]
            code_name_log[3] = code_log_reset[3]
            code_name_log[4] = code_log_reset[4]
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 2:
            self.textEdit_2.clear()
            if self.checkBox_2.isChecked() == False:
                codename = ""
            code_name_log[1] = codename
            code_name_log[2] = code_log_reset[2]
            code_name_log[3] = code_log_reset[3]
            code_name_log[4] = code_log_reset[4]
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 3:
            self.textEdit_2.clear()
            if self.checkBox_3.isChecked() == False:
                codename = ""
            code_name_log[2] = codename
            code_name_log[3] = code_log_reset[3]
            code_name_log[4] = code_log_reset[4]
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 4:
            self.textEdit_2.clear()
            if self.checkBox_4.isChecked() == False:
                codename = ""
            code_name_log[3] = codename
            code_name_log[4] = code_log_reset[4]
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 5:
            self.textEdit_2.clear()
            if self.checkBox_5.isChecked() == False:
                codename = ""
            code_name_log[4] = codename
            code_name_log[5] = code_log_reset[5]
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 6:
            self.textEdit_2.clear()
            if self.checkBox_6.isChecked() == False:
                codename = ""
            code_name_log[5] = codename
            code_name_log[6] = code_log_reset[6]
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)
        elif where == 7:
            self.textEdit_2.clear()
            if self.checkBox_7.isChecked() == False:
                codename = ""
            code_name_log[6] = codename
            result = ' '.join(s for s in code_name_log)
            self.textEdit_2.setText(result)

    def open_button_event(self, bool):
        if bool:
            xlsxindex = self.comboBox.currentIndex() - 1
            if xlsxindex == -1:
                xlsxindex = 7
                os.startfile(list_xlsx_name[xlsxindex])
                status = str(list_xlsx_name[xlsxindex]) + " 여는 중."
                self.statusbar.showMessage(status)

    def open_button_event(self, bool):
        if bool:
            xlsxindex = self.comboBox.currentIndex() - 1
            os.startfile(list_xlsx_name[xlsxindex])



    def xlsx_list(self, MainWindow):

        self.comboBox.clear()
        for i in range(8):
            self.comboBox.insertItem(i, list_option[i])

        self.pushButton.clicked.connect(lambda: self.open_button_event(True))

if __name__ == "__main__":
    import sys

    
    code_log = [' ', ' ', ' ', ' ', ' ', ' ', ' ']
    code_log_reset = [' ', ' ', ' ', ' ', ' ', ' ', ' ']
    code_name_log = [' ', ' ', ' ', ' ', ' ', ' ', ' ']
    list_xlsx_name = ['Companies.xlsx', 'Products.xlsx', 'MainCategory.xlsx', 'MiddleCategory.xlsx', 'SubCategory.xlsx',
                     'Options1.xlsx', 'Options2.xlsx', 'total.xlsx']
    list_option = ['선택', '회사', '품목군', '대분류', '중분류', '소분류', '옵션1', '옵션2']


    app = QtWidgets.QApplication(sys.argv)
    #app.setAttribute(QtCore.Qt.AA_EnableHigwhDpiScaling, True)  # enable highdpi scaling
    #app.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    extra = {

        # Density Scale
        'density_scale': '-2',
    }
    apply_stylesheet(app, 'light_blue.xml', invert_secondary=True, extra=extra)
    mainWindow = QtWidgets.QMainWindow()
    ui = Ui_mainWindow()
    ui.setupUi(mainWindow)
    ui.xlsx_list(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
