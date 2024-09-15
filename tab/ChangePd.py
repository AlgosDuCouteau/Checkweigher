from PyQt5 import QtWidgets, QtCore
from UI import Ui_ChangePD

class ChangePd(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.ChangePD = QtWidgets.QWidget()
        self.ui = Ui_ChangePD()
        self.ui.setupUi(self.ChangePD)
        self.MainWin = MainWin
        self.ui.Confirm.clicked.connect(self.Confirm)
        self.Show_Hide()

    def Show_Hide(self):
        if not self.MainWin.shownChangePd:
            self.ChangePD.show()
            self.MainWin.shownChangePd = True


    def Confirm(self):
        self.ChangePD.close()