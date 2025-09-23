
from re import S
import sys
from tkinter import SE
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import os
from openpyxl import Workbook, load_workbook
from excele_ekle import addExcel
from excelden_oku import  load_excel_to_table


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.showMaximized()

        self.actionswitc_listesi = QAction(MainWindow)
        self.actionswitc_listesi.setObjectName(u"actionswitc_listesi")
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.tabWidget = QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tabWidget.setGeometry(0, 0, 1900,1000)

        self.tabWidget.setLayoutDirection(Qt.LeftToRight)
        self.tabWidget.setAutoFillBackground(False)
        self.tabWidget.setStyleSheet(u"Switch Ekle")
        self.tab = QWidget()
        self.tab.setObjectName(u"tab")
        self.layoutWidget = QWidget(self.tab)
        self.layoutWidget.setObjectName(u"layoutWidget")
        self.layoutWidget.setGeometry(700, 250, 500, 350) #ortala
        self.formLayout = QFormLayout(self.layoutWidget)
        self.formLayout.setObjectName(u"formLayout")
        self.formLayout.setContentsMargins(0, 0, 0, 0)
        self.label = QLabel(self.layoutWidget)
        self.label.setObjectName(u"label")
        self.formLayout.setWidget(0, QFormLayout.LabelRole, self.label)

 
        self.label.setMinimumHeight(50) # height set to 50 pixels 

        self.lineEdit = QLineEdit(self.layoutWidget)
        self.lineEdit.setObjectName(u"lineEdit")
        self.formLayout.setWidget(0, QFormLayout.FieldRole, self.lineEdit)
        self.lineEdit.setMinimumHeight(40) # height set to 50 pixels
       

        self.lineEdit.setStyleSheet("font-size: 11pt;")


        self.label_2 = QLabel(self.layoutWidget)
        self.label_2.setObjectName(u"label_2")
        self.formLayout.setWidget(1, QFormLayout.LabelRole, self.label_2)



     
        self.label_2.setMinimumHeight(50)


        

        self.lineEdit_2 = QLineEdit(self.layoutWidget)
        self.lineEdit_2.setObjectName(u"lineEdit_2")
        self.formLayout.setWidget(1, QFormLayout.FieldRole, self.lineEdit_2)
        self.lineEdit_2.setMinimumHeight(40)
        self.lineEdit_2.setStyleSheet("font-size: 11pt;")

        self.label_3 = QLabel(self.layoutWidget)
        self.label_3.setObjectName(u"label_3")
        self.formLayout.setWidget(2, QFormLayout.LabelRole, self.label_3)

        self.label_3.setMinimumHeight(50)

        self.lineEdit_3 = QLineEdit(self.layoutWidget)
        self.lineEdit_3.setObjectName(u"lineEdit_3")
        self.formLayout.setWidget(2, QFormLayout.FieldRole, self.lineEdit_3)
        self.lineEdit_3.setMinimumHeight(40)
        self.lineEdit_3.setStyleSheet("font-size: 11pt;")

        self.label_4 = QLabel(self.layoutWidget)
        self.label_4.setObjectName(u"label_4")
        self.formLayout.setWidget(3, QFormLayout.LabelRole, self.label_4)

        self.label_4.setMinimumHeight(50)

        self.lineEdit_4 = QLineEdit(self.layoutWidget)
        self.lineEdit_4.setObjectName(u"lineEdit_4")
        self.formLayout.setWidget(3, QFormLayout.FieldRole, self.lineEdit_4)
        self.lineEdit_4.setMinimumHeight(40)
        self.lineEdit_4.setStyleSheet("font-size: 11pt;")

        self.label_5 = QLabel(self.layoutWidget)
        self.label_5.setObjectName(u"label_5")
        self.formLayout.setWidget(4, QFormLayout.LabelRole, self.label_5)

        self.label_5.setMinimumHeight(50)

        self.lineEdit_5 = QLineEdit(self.layoutWidget)
        self.lineEdit_5.setObjectName(u"lineEdit_5")
        self.formLayout.setWidget(4, QFormLayout.FieldRole, self.lineEdit_5)
        self.lineEdit_5.setMinimumHeight(40)
        self.lineEdit_5.setStyleSheet("font-size: 11pt;")

        self.label_6 = QLabel(self.layoutWidget)
        self.label_6.setObjectName(u"label_6")
        self.formLayout.setWidget(5, QFormLayout.LabelRole, self.label_6)

        self.label_6.setMinimumHeight(50)

        self.lineEdit_6 = QLineEdit(self.layoutWidget)
        self.lineEdit_6.setObjectName(u"lineEdit_6")
        self.formLayout.setWidget(5, QFormLayout.FieldRole, self.lineEdit_6)
        self.lineEdit_6.setMinimumHeight(40)
        self.lineEdit_6.setStyleSheet("font-size: 11pt;")

        self.label_7 = QLabel(self.layoutWidget)
        self.label_7.setObjectName(u"label_7")
        self.formLayout.setWidget(6, QFormLayout.LabelRole, self.label_7)

        self.label_7.setMinimumHeight(50)

        self.lineEdit_7 = QLineEdit(self.layoutWidget)
        self.lineEdit_7.setObjectName(u"lineEdit_7")
        self.formLayout.setWidget(6, QFormLayout.FieldRole, self.lineEdit_7)
        self.lineEdit_7.setMinimumHeight(40)
        self.lineEdit_7.setStyleSheet("font-size: 11pt;")

        self.pushButton = QPushButton(self.tab)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(850, 650, 250, 50))
        font = QFont()
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(100)
        self.pushButton.setFont(font)
        self.pushButton.setMouseTracking(True)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        self.tab_2.setObjectName(u"tab_2")
        self.tableWidget = QTableWidget(self.tab_2)
        self.tableWidget.setObjectName(u"tableWidget")
        self.tableWidget.setGeometry(QRect(0, 10, 1800, 850))
        self.tabWidget.addTab(self.tab_2, "")
        MainWindow.setCentralWidget(self.centralwidget)

   

        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        load_excel_to_table(self.tableWidget, "data.xlsx") 
        self.pushButton.clicked.connect(self.add_to_table)
   
        # self.pushButton.clicked.connect(lambda: self.tabWidget.setCurrentWidget(self.tab_2))

        self.lineEdit.returnPressed.connect(self.lineEdit_2.setFocus)
        self.lineEdit_2.returnPressed.connect(self.lineEdit_3.setFocus)
        self.lineEdit_3.returnPressed.connect(self.lineEdit_4.setFocus)
        self.lineEdit_4.returnPressed.connect(self.lineEdit_5.setFocus)
        self.lineEdit_5.returnPressed.connect(self.lineEdit_6.setFocus)
        self.lineEdit_6.returnPressed.connect(self.lineEdit_7.setFocus)
        self.lineEdit_7.returnPressed.connect(self.add_to_table)
        

        self.tabWidget.setCurrentIndex(0)

        QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.actionswitc_listesi.setText(QCoreApplication.translate("MainWindow", u"switc listesi", None))
        self.tabWidget.setWhatsThis(QCoreApplication.translate("MainWindow", u"<html><head/><body><p>Switch Ekle</p></body></html>", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Switch ad\u0131</span></p></body></html>", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Markas\u0131</span></p></body></html>", None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Modeli</span></p></body></html>", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Lokasyonu</span></p></body></html>", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">IP</span></p></body></html>", None))
        self.label_6.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Username</span></p></body></html>", None))
        self.label_7.setText(QCoreApplication.translate("MainWindow", u"<html><head/><body><p><span style=\" font-size:11pt;\">Password</span></p></body></html>", None))
        self.pushButton.setText(QCoreApplication.translate("MainWindow", u"Ekle", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), QCoreApplication.translate("MainWindow", u"Switch Ekle", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), QCoreApplication.translate("MainWindow", u"Switchleri Listele", None))


    def add_to_table(self):
        # Get data from line edits
        switch_adi = self.lineEdit.text()
        markasi = self.lineEdit_2.text()
        modeli = self.lineEdit_3.text()
        lokasyonu = self.lineEdit_4.text()
        ip = self.lineEdit_5.text()
        username = self.lineEdit_6.text()
        password = self.lineEdit_7.text()

        if not all([switch_adi, markasi, modeli, lokasyonu, ip, username, password]):
            QMessageBox.warning(self.centralwidget, "Warning", " Alanlar\u0131 doldurun!")
            return False
        else:
             self.pushButton.clicked.connect(lambda: self.tabWidget.setCurrentWidget(self.tab_2))

        # Add new row
        row_position = self.tableWidget.rowCount()
        self.tableWidget.insertRow(row_position)

         # Add data to the Excel file
        addExcel(switch_adi, markasi, modeli, lokasyonu, ip, username, password)
        load_excel_to_table(self.tableWidget, "data.xlsx")

  

        # Clear all line edits
        self.lineEdit.clear()
        self.lineEdit_2.clear()
        self.lineEdit_3.clear()
        self.lineEdit_4.clear()
        self.lineEdit_5.clear()
        self.lineEdit_6.clear()
        self.lineEdit_7.clear()

        # Switch to the second tab
        self.tabWidget.setCurrentWidget(self.tab_2)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

