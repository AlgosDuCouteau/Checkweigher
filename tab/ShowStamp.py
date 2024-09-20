from PyQt5 import QtWidgets, QtCore, QtGui
from UI import Ui_ShowStamp

class ShowStamp(QtWidgets.QWidget):
    def __init__(self, MainWin, imgFile, parent = None):
        super().__init__(parent)
        self.ShowStamp = QtWidgets.QWidget()
        self.ShowStamp.setWindowFlags(self.ShowStamp.windowFlags() & ~(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinMaxButtonsHint))
        self.ui = Ui_ShowStamp()
        self.ui.setupUi(self.ShowStamp)
        self.MainWin = MainWin
        self.imgFile = imgFile
        self.ui.Confirm.clicked.connect(self.Confirm)
        self.loadImage()
        self.ShowStamp.show()
        self.MainWin.hide()
        
    def loadImage(self):
        pixmap = QtGui.QPixmap(self.imgFile)
        self.ui.Stamp.setPixmap(pixmap)
        self.ui.Stamp.setScaledContents(True)
        
    def Confirm(self):
        self.MainWin.show()
        self.MainWin.qTimer2.start()
        self.MainWin.qTimer3.start()
        self.MainWin.qTimer4.start()
        self.ShowStamp.close()