from PyQt5 import QtWidgets, QtCore
from UI import Ui_Calendar

class Calendar(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.Calen = QtWidgets.QWidget()
        self.ui = Ui_Calendar()
        self.ui.setupUi(self.Calen)
        self.MainWin = MainWin
        self.Show_Hide()
        self.ui.Confirm.clicked.connect(self.Confirm)
        self.i = 0

    def Show_Hide(self):
        if self.MainWin.ui.Calendar.isChecked():
            self.Calen.show()
        else:
            self.MainWin.ui.DD_MM_YY.setText('')
            self.Calen.close()

    def Confirm(self):
        self.i = 1
        self.Calen.close()
        self.MainWin.ui.Calendar.setEnabled(1)
        self.MainWin.ui.DD_MM_YY.setText(str(self.ui.calendar.selectedDate().toString("yyyy-MM-dd")))