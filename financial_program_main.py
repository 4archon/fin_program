from PyQt5 import QtCore, QtGui, QtWidgets
import openpyxl
from financial_calculations import FinIndicators
import pandas as pd


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(720, 434)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_3.addWidget(self.lineEdit)
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.horizontalLayout_3.addWidget(self.pushButton)
        self.pushButton.clicked.connect(self.action_for_button_1)
        self.verticalLayout_5.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_2.addWidget(self.lineEdit_2)
        self.verticalLayout_5.addLayout(self.horizontalLayout_2)
        self.gridLayout.addLayout(self.verticalLayout_5, 1, 0, 1, 1)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.line_2 = QtWidgets.QFrame(self.centralwidget)
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.horizontalLayout.addWidget(self.line_2)
        self.checkBox = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox.setObjectName("checkBox")
        self.horizontalLayout.addWidget(self.checkBox)
        self.checkBox.clicked.connect(self.action_for_checkbox)
        self.line = QtWidgets.QFrame(self.centralwidget)
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.horizontalLayout.addWidget(self.line)
        self.verticalLayout_4.addLayout(self.horizontalLayout)
        self.formLayout_3 = QtWidgets.QFormLayout()
        self.formLayout_3.setObjectName("formLayout_3")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.formLayout_2 = QtWidgets.QFormLayout()
        self.formLayout_2.setObjectName("formLayout_2")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setObjectName("label_3")
        self.verticalLayout.addWidget(self.label_3)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setObjectName("label_4")
        self.verticalLayout.addWidget(self.label_4)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setObjectName("label_5")
        self.verticalLayout.addWidget(self.label_5)
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setTextFormat(QtCore.Qt.AutoText)
        self.label_6.setWordWrap(False)
        self.label_6.setObjectName("label_6")
        self.verticalLayout.addWidget(self.label_6)
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setTextFormat(QtCore.Qt.AutoText)
        self.label_7.setWordWrap(False)
        self.label_7.setObjectName("label_7")
        self.verticalLayout.addWidget(self.label_7)
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setTextFormat(QtCore.Qt.AutoText)
        self.label_8.setWordWrap(False)
        self.label_8.setObjectName("label_8")
        self.verticalLayout.addWidget(self.label_8)
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setTextFormat(QtCore.Qt.AutoText)
        self.label_9.setWordWrap(False)
        self.label_9.setObjectName("label_9")
        self.verticalLayout.addWidget(self.label_9)
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setTextFormat(QtCore.Qt.AutoText)
        self.label_10.setWordWrap(False)
        self.label_10.setObjectName("label_10")
        self.verticalLayout.addWidget(self.label_10)
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setTextFormat(QtCore.Qt.AutoText)
        self.label_11.setWordWrap(False)
        self.label_11.setObjectName("label_11")
        self.verticalLayout.addWidget(self.label_11)
        self.label_12 = QtWidgets.QLabel(self.centralwidget)
        self.label_12.setTextFormat(QtCore.Qt.AutoText)
        self.label_12.setWordWrap(False)
        self.label_12.setObjectName("label_12")
        self.verticalLayout.addWidget(self.label_12)
        self.formLayout_2.setLayout(0, QtWidgets.QFormLayout.LabelRole, self.verticalLayout)
        self.formLayout = QtWidgets.QFormLayout()
        self.formLayout.setObjectName("formLayout")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_3.sizePolicy().hasHeightForWidth())
        self.lineEdit_3.setSizePolicy(sizePolicy)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.lineEdit_3)
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_4.sizePolicy().hasHeightForWidth())
        self.lineEdit_4.setSizePolicy(sizePolicy)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.lineEdit_4)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_5.sizePolicy().hasHeightForWidth())
        self.lineEdit_5.setSizePolicy(sizePolicy)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.lineEdit_5)
        self.lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_6.sizePolicy().hasHeightForWidth())
        self.lineEdit_6.setSizePolicy(sizePolicy)
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.lineEdit_6)
        self.lineEdit_7 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_7.sizePolicy().hasHeightForWidth())
        self.lineEdit_7.setSizePolicy(sizePolicy)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.lineEdit_7)
        self.lineEdit_8 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_8.sizePolicy().hasHeightForWidth())
        self.lineEdit_8.setSizePolicy(sizePolicy)
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.lineEdit_8)
        self.lineEdit_9 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_9.sizePolicy().hasHeightForWidth())
        self.lineEdit_9.setSizePolicy(sizePolicy)
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.LabelRole, self.lineEdit_9)
        self.lineEdit_10 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_10.sizePolicy().hasHeightForWidth())
        self.lineEdit_10.setSizePolicy(sizePolicy)
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.FieldRole, self.lineEdit_10)
        self.lineEdit_11 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_11.sizePolicy().hasHeightForWidth())
        self.lineEdit_11.setSizePolicy(sizePolicy)
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.LabelRole, self.lineEdit_11)
        self.lineEdit_12 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_12.sizePolicy().hasHeightForWidth())
        self.lineEdit_12.setSizePolicy(sizePolicy)
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.FieldRole, self.lineEdit_12)
        self.lineEdit_13 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_13.sizePolicy().hasHeightForWidth())
        self.lineEdit_13.setSizePolicy(sizePolicy)
        self.lineEdit_13.setObjectName("lineEdit_13")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.LabelRole, self.lineEdit_13)
        self.lineEdit_14 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_14.sizePolicy().hasHeightForWidth())
        self.lineEdit_14.setSizePolicy(sizePolicy)
        self.lineEdit_14.setObjectName("lineEdit_14")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.FieldRole, self.lineEdit_14)
        self.lineEdit_15 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_15.sizePolicy().hasHeightForWidth())
        self.lineEdit_15.setSizePolicy(sizePolicy)
        self.lineEdit_15.setObjectName("lineEdit_15")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.LabelRole, self.lineEdit_15)
        self.lineEdit_16 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_16.sizePolicy().hasHeightForWidth())
        self.lineEdit_16.setSizePolicy(sizePolicy)
        self.lineEdit_16.setObjectName("lineEdit_16")
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.FieldRole, self.lineEdit_16)
        self.lineEdit_17 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_17.sizePolicy().hasHeightForWidth())
        self.lineEdit_17.setSizePolicy(sizePolicy)
        self.lineEdit_17.setObjectName("lineEdit_17")
        self.formLayout.setWidget(7, QtWidgets.QFormLayout.LabelRole, self.lineEdit_17)
        self.lineEdit_18 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_18.sizePolicy().hasHeightForWidth())
        self.lineEdit_18.setSizePolicy(sizePolicy)
        self.lineEdit_18.setObjectName("lineEdit_18")
        self.formLayout.setWidget(8, QtWidgets.QFormLayout.LabelRole, self.lineEdit_18)
        self.lineEdit_19 = QtWidgets.QLineEdit(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_19.sizePolicy().hasHeightForWidth())
        self.lineEdit_19.setSizePolicy(sizePolicy)
        self.lineEdit_19.setObjectName("lineEdit_19")
        self.formLayout.setWidget(9, QtWidgets.QFormLayout.LabelRole, self.lineEdit_19)
        self.formLayout_2.setLayout(0, QtWidgets.QFormLayout.FieldRole, self.formLayout)
        self.verticalLayout_2.addLayout(self.formLayout_2)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem)
        self.formLayout_3.setLayout(0, QtWidgets.QFormLayout.LabelRole, self.verticalLayout_2)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        spacerItem1 = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_3.addItem(spacerItem1)
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_3.addWidget(self.pushButton_2)
        self.pushButton_2.clicked.connect(self.action_for_button_2)
        self.formLayout_3.setLayout(0, QtWidgets.QFormLayout.FieldRole, self.verticalLayout_3)
        self.verticalLayout_4.addLayout(self.formLayout_3)
        self.gridLayout.addLayout(self.verticalLayout_4, 2, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 720, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.disable_line(True)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "FinProgram"))
        self.label.setText(_translate("MainWindow", "Выберите файл:"))
        self.pushButton.setText(_translate("MainWindow", "Обзор"))
        self.label_2.setText(_translate("MainWindow", "Название листа с данными:"))
        self.checkBox.setText(_translate("MainWindow", "Указать диапозоны для данных"))
        self.label_3.setText(_translate("MainWindow", "Укажите диапозон для выручки:"))
        self.label_4.setText(_translate("MainWindow", "Укажите диапозон для Capex:"))
        self.label_5.setText(_translate("MainWindow", "Укажите диапозон для Opex:"))
        self.label_6.setText(_translate("MainWindow", "Укажите диапозон для процентов финансирования инвестором:"))
        self.label_7.setText(_translate("MainWindow", "Укажите диапозон для количества периодов кредитовая:"))
        self.label_8.setText(_translate("MainWindow", "Укажите диапозон для процентов по кредитам:"))
        self.label_9.setText(_translate("MainWindow", "Укажите диапозон для изменения оборотного капитала:"))
        self.label_10.setText(_translate("MainWindow", "Укажите ячейку для срока использования:"))
        self.label_11.setText(_translate("MainWindow", "Укажите ячейку для ставки дисконтирования FCFF:"))
        self.label_12.setText(_translate("MainWindow", "Укажите ячейку для ставки дисконтирования FCFE:"))
        self.pushButton_2.setText(_translate("MainWindow", "Расчитать"))

    def disable_line(self, a: bool):
        self.lineEdit_3.setDisabled(a)
        self.lineEdit_4.setDisabled(a)
        self.lineEdit_5.setDisabled(a)
        self.lineEdit_6.setDisabled(a)
        self.lineEdit_7.setDisabled(a)
        self.lineEdit_8.setDisabled(a)
        self.lineEdit_9.setDisabled(a)
        self.lineEdit_10.setDisabled(a)
        self.lineEdit_11.setDisabled(a)
        self.lineEdit_12.setDisabled(a)
        self.lineEdit_13.setDisabled(a)
        self.lineEdit_14.setDisabled(a)
        self.lineEdit_15.setDisabled(a)
        self.lineEdit_16.setDisabled(a)
        self.lineEdit_17.setDisabled(a)
        self.lineEdit_18.setDisabled(a)
        self.lineEdit_19.setDisabled(a)

    def action_for_checkbox(self):
        if self.checkBox.isChecked():
            self.disable_line(False)
        else:
            self.disable_line(True)

    def action_for_button_1(self):
        file_name = QtWidgets.QFileDialog.getOpenFileName(MainWindow, 'Open file', '/home')[0]
        self.lineEdit.setText(file_name)

    @staticmethod
    def activate_message_box(a: int):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowTitle('FinProgram')
        if a == 1:
            msg_box.setText('Неправильно указано имя файла')
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        elif a == 2:
            msg_box.setText('Неправильно указано название страницы')
            msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        elif a == 3:
            msg_box.setText('Данные не соответствующего формата')
            msg_box.setIcon(QtWidgets.QMessageBox.Critical)
        elif a == 4:
            msg_box.setText('Данные не могут быть записаны пока открыт файл')
            msg_box.setIcon(QtWidgets.QMessageBox.Critical)
        elif a == 5:
            msg_box.setText('Данные успешно записаны в файл')
            msg_box.setIcon(QtWidgets.QMessageBox.Information)
        msg_box.exec_()

    def action_for_button_2(self):
        def read_file(file_name: str, sheet_name: str, open_default=True):
            book = openpyxl.open(file_name, read_only=True)
            if open_default:
                try:
                    sheet = book[sheet_name]
                    b = []
                    for row in range(1, 11):
                        a = []
                        if row < 8:
                            for column in range(1, sheet.max_column):
                                a += [sheet[row][column].value]
                            a = tuple(a)
                            b += [a]
                        else:
                            b += [sheet[row][1].value]
                    b = tuple(b)
                    book.close()
                    return b
                except KeyError:
                    book.close()
                    raise KeyError

            else:
                try:
                    sheet = [sheet_name]
                    b = []
                    a = sheet[self.lineEdit_3.text():self.lineEdit_4.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_5.text():self.lineEdit_6.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_7.text():self.lineEdit_8.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_9.text():self.lineEdit_10.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_11.text():self.lineEdit_12.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_13.text():self.lineEdit_14.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_15.text():self.lineEdit_16.text()][0]
                    a = tuple(map(lambda x: x.value, a))
                    b += [a]
                    a = sheet[self.lineEdit_17.text()].value
                    b += [a]
                    a = sheet[self.lineEdit_18.text()].value
                    b += [a]
                    a = sheet[self.lineEdit_19.text()].value
                    b += [a]
                    b = tuple(b)
                    book.close()
                    return b
                except KeyError:
                    book.close()
                    raise KeyError

        def len_check(a: tuple):
            len_for_compare = len(a[0])
            for i in range(len(a)):
                if i < 7:
                    if len(a[i]) != len_for_compare:
                        return False
                else:
                    if type(a[i]) not in (int, float):
                        return False
            return True

        def delete_excess_list(file_name: str, sheet_name: str):
            book = openpyxl.open(file_name)
            if sheet_name in book.sheetnames:
                del book[sheet_name]
            book.save(file_name)
            book.close()

        try:
            if self.checkBox.isChecked():
                data_tuple = read_file(self.lineEdit.text(), self.lineEdit_2.text(), open_default=False)
            else:
                data_tuple = read_file(self.lineEdit.text(), self.lineEdit_2.text())
        except openpyxl.utils.exceptions.InvalidFileException:
            self.activate_message_box(1)
        except FileNotFoundError:
            self.activate_message_box(1)
        except KeyError:
            self.activate_message_box(2)
        if 'data_tuple' in locals():
            if len_check(data_tuple):
                main_object = FinIndicators(*data_tuple)
                try:
                    delete_excess_list(self.lineEdit.text(), 'Fin_Indicators')
                    with pd.ExcelWriter(self.lineEdit.text(), mode='a') as writer:
                        main_object.data_frame_properties.to_excel(writer, startrow=0, startcol=0,
                                                                   sheet_name='Fin_Indicators')
                        main_object.data_frame_indicators.to_excel(writer, startrow=29, startcol=0,
                                                                   sheet_name='Fin_Indicators')
                except PermissionError:
                    self.activate_message_box(4)
                else:
                    self.activate_message_box(5)
            else:
                self.activate_message_box(3)


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
