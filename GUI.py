import sys
import PullData
from datetime import datetime
import calendar
from PyQt5.QtWidgets import QApplication, QWidget, QCalendarWidget, QPushButton, QMessageBox
from PyQt5.QtCore import QDate


class CalendarDemo(QWidget):
    global currentYear, currentMonth

    currentMonth = datetime.now().month
    currentYear = datetime.now().year

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Calendar Demo')
        self.setGeometry(300, 300, 450, 300)
        self.initUI()

    def initUI(self):
        self.calendar = QCalendarWidget(self)
        self.calendar.move(20, 20)
        self.calendar.setGridVisible(True)
        self.button = QPushButton('Confirmar')
        self.calendar.layout().addWidget(self.button)



        self.calendar.setMinimumDate(QDate(currentYear-10, currentMonth, 1))
        self.calendar.setMaximumDate(
            QDate(currentYear, currentMonth + 1, calendar.monthrange(currentYear, currentMonth)[1]))

        qDate = self.calendar.setSelectedDate(QDate(currentYear, currentMonth, 1))

        self.calendar.clicked.connect(self.printDateInfo)
        self.button.clicked.connect(self.on_button_clicked)

    def printDateInfo(self, qDate):
        print('{0}/{1}/{2}'.format(qDate.month(), qDate.day(), qDate.year()))
        print(f'Day Number of the year: {qDate.dayOfYear()}')
        print(f'Day Number of the week: {qDate.dayOfWeek()}')

    def on_button_clicked(self):
        PullData.main('10-03-2020')
        alert = QMessageBox()
        alert.setText('Baixando processos')
        alert.exec_()

def main():
    app = QApplication(sys.argv)
    window = CalendarDemo()
    window.show()
    sys.exit(app.exec_())


main()
