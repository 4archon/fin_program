from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMainWindow
import openpyxl
from financial_calculations import FinIndicators


class Ui_MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setupUi()

    def setupUi(self):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(570, 549)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(460, 30, 75, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.action_for_button1)
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(20, 30, 431, 21))
        self.lineEdit.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 10, 211, 16))
        self.label.setObjectName("label")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(170, 70, 281, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(30, 70, 141, 16))
        self.label_2.setObjectName("label_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_3.setGeometry(QtCore.QRect(210, 100, 51, 20))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(30, 100, 171, 16))
        self.label_3.setObjectName("label_3")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_4.setGeometry(QtCore.QRect(280, 100, 51, 20))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(30, 130, 171, 16))
        self.label_4.setObjectName("label_4")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_5.setGeometry(QtCore.QRect(210, 130, 51, 20))
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_6.setGeometry(QtCore.QRect(280, 130, 51, 20))
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.lineEdit_7 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_7.setGeometry(QtCore.QRect(210, 160, 51, 20))
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(30, 160, 171, 16))
        self.label_5.setObjectName("label_5")
        self.lineEdit_8 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_8.setGeometry(QtCore.QRect(280, 160, 51, 20))
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(30, 180, 171, 51))
        self.label_6.setTextFormat(QtCore.Qt.AutoText)
        self.label_6.setWordWrap(True)
        self.label_6.setObjectName("label_6")
        self.lineEdit_10 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_10.setGeometry(QtCore.QRect(280, 200, 51, 20))
        self.lineEdit_10.setObjectName("lineEdit_10")
        self.lineEdit_9 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_9.setGeometry(QtCore.QRect(210, 200, 51, 20))
        self.lineEdit_9.setObjectName("lineEdit_9")
        self.lineEdit_11 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_11.setGeometry(QtCore.QRect(210, 250, 51, 20))
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(30, 230, 171, 51))
        self.label_7.setTextFormat(QtCore.Qt.AutoText)
        self.label_7.setWordWrap(True)
        self.label_7.setObjectName("label_7")
        self.lineEdit_12 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_12.setGeometry(QtCore.QRect(280, 250, 51, 20))
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.lineEdit_13 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_13.setGeometry(QtCore.QRect(210, 290, 51, 20))
        self.lineEdit_13.setObjectName("lineEdit_13")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(30, 280, 171, 31))
        self.label_8.setTextFormat(QtCore.Qt.AutoText)
        self.label_8.setWordWrap(True)
        self.label_8.setObjectName("label_8")
        self.lineEdit_14 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_14.setGeometry(QtCore.QRect(280, 290, 51, 20))
        self.lineEdit_14.setObjectName("lineEdit_14")
        self.lineEdit_15 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_15.setGeometry(QtCore.QRect(210, 330, 51, 20))
        self.lineEdit_15.setObjectName("lineEdit_15")
        self.lineEdit_16 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_16.setGeometry(QtCore.QRect(280, 330, 51, 20))
        self.lineEdit_16.setObjectName("lineEdit_16")
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setGeometry(QtCore.QRect(30, 320, 171, 31))
        self.label_9.setTextFormat(QtCore.Qt.AutoText)
        self.label_9.setWordWrap(True)
        self.label_9.setObjectName("label_9")
        self.lineEdit_17 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_17.setGeometry(QtCore.QRect(210, 370, 51, 20))
        self.lineEdit_17.setObjectName("lineEdit_17")
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setGeometry(QtCore.QRect(30, 360, 171, 31))
        self.label_10.setTextFormat(QtCore.Qt.AutoText)
        self.label_10.setWordWrap(True)
        self.label_10.setObjectName("label_10")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(30, 400, 171, 31))
        self.label_11.setTextFormat(QtCore.Qt.AutoText)
        self.label_11.setWordWrap(True)
        self.label_11.setObjectName("label_11")
        self.lineEdit_18 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_18.setGeometry(QtCore.QRect(210, 410, 51, 20))
        self.lineEdit_18.setObjectName("lineEdit_18")
        self.label_12 = QtWidgets.QLabel(self.centralwidget)
        self.label_12.setGeometry(QtCore.QRect(30, 440, 171, 31))
        self.label_12.setTextFormat(QtCore.Qt.AutoText)
        self.label_12.setWordWrap(True)
        self.label_12.setObjectName("label_12")
        self.lineEdit_19 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_19.setGeometry(QtCore.QRect(210, 450, 51, 20))
        self.lineEdit_19.setObjectName("lineEdit_19")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(460, 480, 75, 23))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.action_for_button2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 570, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Fin program"))
        self.pushButton.setText(_translate("MainWindow", "Обзор"))
        self.label.setText(_translate("MainWindow", "Выберите файл"))
        self.label_2.setText(_translate("MainWindow", "Выберете название листа:"))
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

    def action_for_button1(self):
        file_name = QFileDialog.getOpenFileName(self, 'Open file', '/home')[0]
        self.file_name = file_name
        self.lineEdit.setText(file_name)

    def action_for_button2(self):
        def transform(cells):
            a = []
            cells = cells[0]
            for i in cells:
                a += [i.value]
            return tuple(a)

        def write_to_excel(proper_list: list, name_list: list, file):
            wb = openpyxl.load_workbook(file, read_only=False)
            wb.create_sheet('new_fin_indicators', 1)
            ws = wb['new_fin_indicators']
            for i in range(len(proper_list)):
                for j in range(-1, len(proper_list[i])):
                    if j == -1:
                        cell_column = 'A'
                        cell_row = str(i + 1)
                        ws[cell_column + cell_row] = name_list[i]
                    else:
                        cell_column = openpyxl.utils.get_column_letter(j + 2)
                        cell_row = str(i + 1)
                        ws[cell_column + cell_row] = proper_list[i][j]
            wb.save(file)
            wb.close()

        def write_to_excel_indicators(indicators_fcff: list, indicators_fcfe: list, name_indicators_list: list, file):
            wb = openpyxl.load_workbook(file, read_only=False)
            ws = wb['new_fin_indicators']
            cell_column_name = 'A'
            cell_column_fcff = 'B'
            cell_column_fcfe = 'C'
            for i in range(len(indicators_fcff)):
                cell_row = str(27 + i)
                ws[cell_column_name + cell_row] = name_indicators_list[i]
                ws[cell_column_fcff + cell_row] = indicators_fcff[i]
                ws[cell_column_fcfe + cell_row] = indicators_fcfe[i]
            wb.save(file)
            wb.close()

        book = openpyxl.open(self.file_name, read_only=True)
        sheet = book[self.lineEdit_2.text()]
        revenue = transform(sheet[self.lineEdit_3.text():self.lineEdit_4.text()])
        capex = transform(sheet[self.lineEdit_5.text():self.lineEdit_6.text()])
        opex = transform(sheet[self.lineEdit_7.text():self.lineEdit_8.text()])
        investments_percentage = transform(sheet[self.lineEdit_9.text():self.lineEdit_10.text()])
        number_of_credit_periods = transform(sheet[self.lineEdit_11.text():self.lineEdit_12.text()])
        credit_percentage = transform(sheet[self.lineEdit_13.text():self.lineEdit_14.text()])
        change_working_cap = transform(sheet[self.lineEdit_15.text():self.lineEdit_16.text()])
        useful_life = sheet[self.lineEdit_17.text()].value
        discount_rate_fcff = sheet[self.lineEdit_18.text()].value
        discount_rate_fcfe = sheet[self.lineEdit_19.text()].value
        book.close()
        self.main_object = FinIndicators(revenue, capex, opex, investments_percentage, number_of_credit_periods,
                                         credit_percentage, change_working_cap, useful_life, discount_rate_fcff,
                                         discount_rate_fcfe)
        name_list = ['Выручка', 'Расходы', 'CAPEX', 'Оборотный капитал', 'Финансирование', 'За счет акционера',
                     'За счет реинвестировани средств акционерного денежного потока', 'За счет заемных средств',
                     'EBITDA', 'Амортизация', 'EBIT', 'Проценты по кредитам', 'EBT', 'Налог на прибыль',
                     'Чистая прибыль', 'Изменение оборотного капитала', 'FCFF', 'Кумулятивный FCFF', 'DCF по FCFF',
                     'Кумулятивный DCF по FCFF', 'Изменение долга', 'FCFE', 'Кумулятивный FCFE', 'DCF по FCFE',
                     'Кумулятивный DCF по FCFE']
        proper_list = [self.main_object.revenue, self.main_object.costs, self.main_object.capex, self.main_object.opex,
                       self.main_object.costs, self.main_object.financing_equity,
                       self.main_object.financing_reinvestment, self.main_object.financing_credit,
                       self.main_object.ebitda, self.main_object.amortization, self.main_object.ebit,
                       self.main_object.interests, self.main_object.ebt, self.main_object.taxes,
                       self.main_object.net_profit, self.main_object.change_working_cap, self.main_object.fcff,
                       self.main_object.cumulative_fcff, self.main_object.dcf_fcff,
                       self.main_object.cumulative_dcf_fcff, self.main_object.debt_change, self.main_object.fcfe,
                       self.main_object.cumulative_fcfe, self.main_object.dcf_fcfe,
                       self.main_object.cumulative_dcf_fcfe]
        name_indicators_list = ['Инвестиционные показатели', 'Ставка дисконтирования', 'Срок окупаемости, лет',
                                'DPP, лет', 'NPV', 'irr, %', 'pi, %']
        indicators_fcff = ['по FCFF'] + list(self.main_object.indicators_fcff)
        indicators_fcfe = ['по FCFE'] + list(self.main_object.indicators_fcfe)
        write_to_excel(proper_list, name_list, self.file_name)
        write_to_excel_indicators(indicators_fcff, indicators_fcfe, name_indicators_list, self.file_name)
        # sys.exit()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi()
    MainWindow.show()
    sys.exit(app.exec_())
