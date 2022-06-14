from datetime import datetime
import sys
import os

from PyQt5 import QtWidgets, QtGui
from PyQt5 import QtCore
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QFileDialog, QLabel, QTableWidgetItem, QHeaderView, QMessageBox
from dohody import Ui_MainWindow
from read_excel import Read
from save import SaveExcel
from AboutForm import Ui_Dialog
import images_rc
import constants

class MyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon(':/img/roskazna.png'))
        self.ui.open_file_one.clicked.connect(self.open_file)
        self.ui.open_file_two.clicked.connect(self.open_file)
        self.ui.save.clicked.connect(self.save)
        self.aboutForm = None
        self.filename = ''
        self.filename_one = ''
        self.check_one = False
        self.filename_two = ''
        self.check_two = False
        self.headers_vert = ['Всего поступило по сч 40101/03100',
                             'Возврат  излишне уплаченных сумм',
                             'Всего перечислено  в бюджет',
                             'Консолидированный бюджет (ст.I  + ст. II)',
                             'Статья I. федеральный бюджет, в т.ч:',
                             'НДС на товары, реализуемые на территории РФ',
                             'НДС на товары, ввозимые на территории РФ',
                             'Налог на прибыль',
                             'Статья II. консолидированный бюджет области',
                             'в том числе:',
                             'областной бюджет, в т.ч:',
                             'НДФЛ',
                             'Налог на прибыль организаций',
                             'местные бюджеты, в т.ч:',
                             'НДФЛ',
                             'Земельный налог с организаций',
                             'Налоги на совокупный доход',
                             'Статья III. государственные внебюджетные фонды',
                             'в том числе:',
                             'Пенсионный фонд',
                             'Фонд социального страхования',
                             'Федеральный фонд медицинского страхования',
                             'Территориальный фонд медицинского страхования',
                             'Статья IY.Иные получатели ( МОУ ФК )',
                             'Остаток на  счете 40101',
                             'НВС глава 100',
                             'Всего по разделу III',
                             'федеральный бюджет',
                             'областной бюджет',
                             'местные бюджеты',
                             'ГВФ']
        self.headers_horiz = []
        stylesheet = "::section{background-color:rgb(68, 68, 68); color: white;}"
        self.ui.table.horizontalHeader().setStyleSheet(stylesheet)
        self.ui.table.verticalHeader().setStyleSheet(stylesheet)
        self.ui.table.verticalHeader(
        ).setSectionResizeMode(QHeaderView.Stretch)
        self.ui.menu.actions()[0].triggered.connect(self.open_about)
        self.ui.save.setEnabled(False)
        self.result_one = {}
        self.result_two = {}
        self.thread_one_ended = False
        self.thread_two_ended = False

    def open_file(self):
        self.filename = QFileDialog.getOpenFileName(
            None, 'Открыть', os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'),
            'All Files(*.xlsx)')
        sender = self.sender()
        if str(self.filename) in "('', '')":
            self.ui.statusbar.showMessage('Файл не выбран')
        else:
            self.ui.save.setEnabled(False)
            if sender.text() == 'Загрузить таблицу 1':
                self.ui.status_one.setPixmap(QPixmap(":/img/good.png"))
                self.filename_one = self.filename
                self.check_one = True
                self.thread_one_ended = False
                self.check_two = False
                self.new_thread()
            else:
                self.ui.status_two.setPixmap(QPixmap(":/img/good.png"))
                self.filename_two = self.filename
                self.check_one = False
                self.check_two = True
                self.thread_two_ended = False
                self.new_thread()
        #     self.new_thread()

    def new_thread(self):
        self.my_thread = Read(my_window=self)
        self.my_thread.start()
        self.my_thread.result.connect(self.get_result)

    @QtCore.pyqtSlot(dict)
    def get_result(self, dict):
        self.ui.table.setColumnCount(2)
        self.ui.table.setRowCount(len(dict))
        self.ui.table.setVerticalHeaderLabels(self.headers_vert)
        current_year = datetime.now().year
        if self.check_one:
            self.result_one = dict
            schet = 0
            for key, value in self.result_one.items():
                self.ui.table.setItem(
                    schet, 0, QTableWidgetItem(str(value)))
                schet = schet + 1
                self.headers_horiz = [f'в {current_year-1} году за соответс-твующий период',
                                      f'в {current_year} году за соответс-твующий период']
                self.ui.table.setHorizontalHeaderLabels(self.headers_horiz)
                self.ui.table.resizeColumnsToContents()
        else:
            self.result_two = dict
            schet = 0
            for key, value in self.result_two.items():
                self.ui.table.setItem(
                    schet, 1, QTableWidgetItem(str(value)))
                schet = schet + 1
                self.headers_horiz = [f'в {current_year-1} году за соответс-твующий период',
                                      f'в {current_year} году за соответс-твующий период']
                self.ui.table.setHorizontalHeaderLabels(self.headers_horiz)
                self.ui.table.resizeColumnsToContents()

        if self.thread_one_ended and self.thread_two_ended:
            self.ui.save.setEnabled(True)

    def save(self):
        saves = SaveExcel()
        wb = saves.save_excel(self.result_one, self.result_two)

        file_save, _ = QFileDialog.getSaveFileName(
            self, 'Сохранить', 'Поступления на', 'All Files(*.xlsx)')
        try:
            if str(file_save) != "":
                wb.save(file_save)
                self.ui.statusbar.showMessage('Таблица сохранена')
        except PermissionError as err:
            messagebox = QMessageBox(
                parent=self, text='Ошибка доступа. Необходимо закрыть файл', detailedText=str(err))
            messagebox.setWindowTitle('Внимание!')
            messagebox.setStyleSheet(
                '.QPushButton{background-color: #444444;color: white;}')
            messagebox.show()
        except Exception as ex:
            messagebox = QMessageBox(
                parent=self, text='Ошибка', detailedText=str(ex))
            messagebox.setWindowTitle('Внимание!')
            messagebox.setStyleSheet(
                '.QPushButton{background-color: #444444;color: white;}')
            messagebox.show()

    def open_about(self):
        if (self.aboutForm != None):
            self.aboutForm.close()
        self.aboutForm = AboutWindows()
        self.aboutForm.show()


class AboutWindows(QtWidgets.QDialog):

    def __init__(self):
        super(AboutWindows, self).__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.lbl_name_version.setText(self.ui.lbl_name_version.text() + f' {constants.app_version}')
        self.setWindowTitle('О программе')



app = QtWidgets.QApplication([])
application = MyWindow()
application.setWindowTitle("Конвертер excel Доходы")
application.show()

sys.exit(app.exec())
