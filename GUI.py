import sys
import PullData
from datetime import datetime
import calendar
from PyQt5.QtWidgets import QApplication, QWidget, QCalendarWidget, QPushButton, QMessageBox
from PyQt5.QtCore import QDate


class CalendarDemo(QWidget):
    global currentYear, currentMonth, final_date
    currentMonth = datetime.now().month
    currentYear = datetime.now().year
    final_date = '{0}-{1}-{2}'.format(datetime.now().day,datetime.now().month,datetime.now().year)

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Calendario')
        self.setGeometry(300, 300, 680, 460)
        self.initUI()

    def initUI(self):
        self.calendar = QCalendarWidget(self)
        self.calendar.setGeometry(20,20,630,410)
        self.calendar.setGridVisible(True)
        self.button = QPushButton('Confirmar')
        self.calendar.layout().addWidget(self.button)

        self.calendar.setMinimumDate(QDate(currentYear-10, currentMonth, 1))
        self.calendar.setMaximumDate(QDate(currentYear, currentMonth + 1, calendar.monthrange(currentYear, currentMonth)[1]))

        self.calendar.setSelectedDate(QDate(currentYear, currentMonth, 1))

        self.calendar.clicked.connect(self.printDateInfo)
        self.button.clicked.connect(self.on_button_clicked)

    def printDateInfo(self, qDate):
        self.final_date = '{0}-{1}-{2}'.format(qDate.day(), qDate.month(), qDate.year())
        # print(self.final_date)
        # print(f'Day Number of the year: {qDate.dayOfYear()}')
        # print(f'Day Number of the week: {qDate.dayOfWeek()}')

    def on_button_clicked(self):
        # print(self.final_date)
        PullData.main(self.final_date)
        alert = QMessageBox()
        alert.setText('Baixando processos')
        alert.exec_()

def main():
    app = QApplication(sys.argv)
    window = CalendarDemo()
    window.show()
    sys.exit(app.exec_())


main()
