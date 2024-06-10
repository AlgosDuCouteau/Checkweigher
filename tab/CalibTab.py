from PyQt5 import QtWidgets, QtCore
from UI import Ui_CalibTab

class CalibTab(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.CalibTab = QtWidgets.QWidget()
        self.ui = Ui_CalibTab()
        self.ui.setupUi(self.CalibTab)
        self.ui.Confirm.setText('Xác nhận')
        self.ui.Confirm2.setDisabled(1)
        self.CalibTab.show()
        self.MainWin = MainWin
        self.MainWin.hide()
        self.ui.Arrow2.hide()
        self.ui.Arrow3.hide()
        self.ui.ConfirmAll.setText('Quay lại')
        self.ui.Confirm.clicked.connect(self.FirstConfirm)
        self.ui.Confirm2.clicked.connect(self.SecondConfirm)
        self.ui.ConfirmAll.clicked.connect(self.ConfirmAll)
        self.ui.ZeroRet.clicked.connect(lambda: self.MainWin.ZeroRet())
        self.ui.MaxWe.setText(str(self.MainWin.maxwe))
        self.ui.MinWe.setText(str(self.MainWin.minwe))
        self.maxw = float(self.MainWin.maxwe)
        self.minw = float(self.MainWin.minwe)
        self.ui.ProductName.setText(str(self.MainWin.ui.ProductName.text()))
        self.step = 0
        self.qTimer1 = QtCore.QTimer()
        self.qTimer1.setInterval(100)
        self.qTimer1.timeout.connect(self.ShowDataR)
        self.qTimer1.start()

    def ShowDataR(self):
        if self.step == 0 or self.step == 2:
            self.ui.PackageWe.display(self.MainWin.scale.dataR)
        elif self.step == 1:
            self.ui.BunchWe.display(abs(self.packwe - self.MainWin.scale.dataR))

    @QtCore.pyqtSlot()
    def FirstConfirm(self):
        self.ui.Confirm.setDisabled(1)
        self.ui.Confirm2.setEnabled(1)
        self.ui.ConfirmAll.setText('Quay lại')
        self.packwe = float(self.ui.PackageWe.value())
        self.ui.Arrow1.hide()
        self.ui.Arrow2.show()
        self.ui.Arrow3.hide()
        self.step = 1

    def SecondConfirm(self):
        self.ui.Confirm2.setDisabled(1)
        self.ui.Confirm.setEnabled(1)
        self.bunchwe = float(self.ui.BunchWe.value())
        if abs(self.packwe - self.bunchwe*12) >= 0.35 and abs(self.packwe - self.bunchwe*12) <= 1.55:
            if self.packwe <= self.bunchwe*12.25 + self.packwe - self.bunchwe*12 and \
            self.packwe >= self.bunchwe*11.75 + self.packwe - self.bunchwe*12 and \
            self.packwe > 0:
                self.maxw = round(self.packwe + abs(self.bunchwe/2), 2)
                self.minw = round(self.packwe - abs(self.bunchwe/2), 2)
                self.ui.ConfirmAll.setText('Cập nhật')
                self.ui.Arrow3.show()
                self.step = 2
                # self.UpdateData()
            else:
                self.step = 0
                self.MainWin.in4(infor = "Cân khối lượng sai")
        else:
            self.step = 0
            self.MainWin.in4(infor = "Cân khối lượng sai")
        self.ui.Confirm.setText('Cân lại')
        self.ui.MaxWe.setText(str(self.maxw))
        self.ui.MinWe.setText(str(self.minw))
        self.ui.Arrow1.show()
        self.ui.Arrow2.hide()

    def ConfirmAll(self):
        self.CalibTab.destroy()
        self.qTimer1.stop()
        self.MainWin.maxwe = round(float(self.maxw), 2)
        self.MainWin.minwe = round(float(self.minw), 2)
        self.MainWin.ui.MinWe.setText(str(self.MainWin.minwe))
        self.MainWin.ui.MaxWe.setText(str(self.MainWin.maxwe))
        self.MainWin.qTimer1.start()
        self.MainWin.qTimer2.start()
        self.MainWin.qTimer3.start()
        self.MainWin.show()
        self.MainWin.ui.ProductID.setFocus()
    
    # def UpdateData(self):
    #     os.system('taskkill /F /IM EXCEL.exe')
    #     wb = Workbook()
    #     self.wb = wb.open(categorization_file)
    #     self.ws = self.wb['Database']
    #     cell = ['Q%d'  % (self.MainWin.n + 2), 'P%d'  % (self.MainWin.n + 2)]
    #     self.ws.cell(cell[0]).value = self.maxw*1000
    #     self.ws.cell(cell[1]).value = self.minw*1000
    #     self.wb.save(categorization_file)
    #     self.wb.close()

    