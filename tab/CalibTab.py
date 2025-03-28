from PyQt5 import QtWidgets, QtCore
from UI import Ui_CalibTab

class CalibTab(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.CalibTab = QtWidgets.QWidget()
        self.CalibTab.setWindowFlags(self.CalibTab.windowFlags() & ~(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinMaxButtonsHint))
        self.ui = Ui_CalibTab()
        self.ui.setupUi(self.CalibTab)
        self.ui.Confirm.setText('Xác nhận')
        self.ui.Confirm2.setDisabled(1)
        self.CalibTab.show()
        self.MainWin = MainWin
        self.MainWin.hide()
        if self.MainWin.data['Dinh_dang_Tem_Thung'].item() == 'OEM':
            self.quantity_batch = int(self.MainWin.data['So_luong_cay_bo_trong_thung'].item())
        else:
            self.quantity_batch = int(int(self.MainWin.data['Quantity_thung'].item())/int(self.MainWin.data['So_luong_liner_trong_bo'].item()))
        self.ui.label.setText(f"Khối lượng\n1 thùng {self.quantity_batch} gói (kg)")
        self.ui.label_2.setText(f"Khối lượng\n1 gói {int(self.MainWin.data['So_luong_liner_trong_bo'].item())} đôi (kg)")
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
        if abs(self.packwe - self.bunchwe*self.quantity_batch) >= 0.35 and abs(self.packwe - self.bunchwe*self.quantity_batch) <= 1.55:
            if self.packwe <= self.bunchwe*(self.quantity_batch+0.25) + self.packwe - self.bunchwe*self.quantity_batch and \
            self.packwe >= self.bunchwe*(self.quantity_batch-0.25) + self.packwe - self.bunchwe*self.quantity_batch and \
            self.packwe > 0:
                self.maxw = round(self.packwe + abs(self.bunchwe*self.MainWin.range), 2)
                self.minw = round(self.packwe - abs(self.bunchwe*self.MainWin.range), 2)
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
        self.CalibTab.close()
        self.qTimer1.stop()
        self.MainWin.maxwe = round(float(self.maxw), 2)
        self.MainWin.minwe = round(float(self.minw), 2)
        self.MainWin.ui.MinWe.setText(str(self.MainWin.minwe))
        self.MainWin.ui.MaxWe.setText(str(self.MainWin.maxwe))
        self.MainWin.show()
        self.MainWin.qTimer2.start()
        self.MainWin.qTimer3.start()
        self.MainWin.qTimer4.start()
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

    