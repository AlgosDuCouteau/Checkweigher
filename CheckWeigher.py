import sys, os, polars as pl, logging
from PyQt5 import QtWidgets, QtCore
from UI import Ui_CheckWeigher
from tab import *
from Function import *

logging.basicConfig(filename="errors.txt",
                    format='\n%(asctime)s %(message)s',
                    filemode='a')

class MainGUI(QtWidgets.QMainWindow):
    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        self.ui = Ui_CheckWeigher()
        self.ui.setupUi(self)
        self.show()
        fileData = data['fileData']
        filePrint = data['filePrint']
        fileUpdateData = data['GdriveData']
        fileUpdatePrint = data['GdrivePrint']
        portScale = data['portScale']
        portArduino = data['portArduino']
        Ard2Convey = data['Ard2Convey']
        self.Delay2Print = data['Delay2Print']
        self.Quan2Print = data['Quan2Print']
        self.loaddata = LoadData(self, fileData=fileData, filePrint=filePrint, fileUpdateData=fileUpdateData, fileUpdatePrint=fileUpdatePrint)
        self.print = Print(self, filePrint=filePrint)
        self.scale = GetScale(portScale=portScale, portArduino=portArduino, Ard2Convey=Ard2Convey)
        self.scale.start()
        self.ui.UpdateData.clicked.connect(self.UpdateData)
        self.ui.CheckStamp.clicked.connect(self.CheckStamp)
        self.ui.ArduinoCon.clicked.connect(self.ArduinoCon)
        self.ui.ZeroRet.clicked.connect(self.ZeroRet)
        self.ui.Calib.clicked.connect(self.Calib)
        self.ui.Calendar.clicked.connect(self.Calendar)
        self.ui.ProductID.returnPressed.connect(self.ProductID)
        self.ui.ProductID.setFocus()
        self.qTimer1 = QtCore.QTimer()
        self.qTimer1.setInterval(100)
        self.qTimer1.timeout.connect(self.ShowDataR)
        self.qTimer1.start()
        self.qTimer2 = QtCore.QTimer()
        self.qTimer2.setInterval(100)
        self.qTimer2.timeout.connect(self.ShowStt)
        self.qTimer2.start()
        self.qTimer3 = QtCore.QTimer()
        self.qTimer3.setInterval(100)
        self.qTimer3.timeout.connect(self.Compare)
        self.qTimer3.start()
        self.qTimer4 = QtCore.QTimer()
        self.qTimer4.setInterval(200)
        self.qTimer4.timeout.connect(self.dataRfilter)
        self.qTimer4.start()
        self.count, self.index, self.k, self.t1, self.t2, self.maxwe, self.minwe,  self.dataR = 0, 0, 0, 0, 0, 0, 0, 0
        self.printed = True, False
        self.seq = ['rgb(255, 255, 255)', 'rgb(255, 0, 0)', 'rgb(255, 255, 0)', 'rgb(0, 255, 0)']
        self.filterlist = []

    def in4(self, infor:str, title = 'Lỗi'):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(infor)
        if title == 'Lỗi':
            msg.setIcon(QtWidgets.QMessageBox.Warning)
        elif title == 'Thông tin':
            msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.exec_()
        return

    def resetdata(self):
        self.ui.ProductID.clear()
        self.ui.PdID.clear()
        self.ui.CATiD.clear()
        self.ui.INTiD.clear()
        self.ui.ProductName.clear()
        self.ui.MaxWe.clear()
        self.ui.MinWe.clear()
        self.ui.PO.clear()
        self.ui.Remain.clear()
        self.ui.Quan.clear()
        self.ui.Calendar.setChecked(0)

    def dataRfilter(self):
        if len(self.ui.PdID.text()) <= 0 or self.maxwe == 0 or self.minwe == 0:
            return
        if self.dataR != 0 and self.maxwe != 0 and self.dataR < self.maxwe:
            self.filterlist.append(self.dataR)
            if len(self.filterlist) > 1:
                if self.filterlist[0] - self.filterlist[1] > self.filterlist[1]*4 and self.filterlist[0] - self.filterlist[1] > 0:
                    self.scale.zeroret = True
                    self.filterlist = []
                else:
                    self.filterlist = []

    def ShowDataR(self):
        if not self.scale._run_flag:
            return
        self.dataR = self.scale.dataR
        if self.isHidden() == 0:
            self.dataR = self.scale.dataR
            self.ui.CurrentWe.display(float(self.dataR))

    def ShowPO(self):
        self.ui.Quan.setText(str(int(self.count)))

    def Compare(self):
        if len(self.ui.PdID.text()) <= 0 or self.maxwe == 0 or self.minwe == 0:
            return
        if self.dataR < self.minwe/11.5*3 and not self.scale.full:
            self.printed = False
            self.t1 = 0
            self.t2 = 0
            return
        if not self.printed and \
            ((self.dataR+abs(12-self.Quan2Print)*(self.maxwe-self.minwe) < self.maxwe and self.dataR+abs(12-self.Quan2Print)*(self.maxwe-self.minwe) > self.minwe) or
            (self.dataR <= self.maxwe and self.dataR >= self.minwe)):
            if self.t1 < self.Delay2Print:
                self.t1 += 1
            if self.t1 >= self.Delay2Print:
                self.CheckStamp()
                self.count += 1
                self.printed = True
                self.ShowPO()
                self.t1 = 0
        else:
            self.t1 = 0
        if self.t2 > 0:
            self.ui.Status.setText(str(float(self.t2/10)))
        else:
            self.ui.Status.setText("")
        if self.dataR <= self.maxwe and self.dataR >= self.minwe and not self.scale.full:
            if self.t2 < self.Delay2Print:
                self.t2 += 1
            if self.t2 >= self.Delay2Print:
                self.scale.full = True
                self.t2 = 0
        elif self.dataR < self.minwe/11.5:
            if self.scale.full:
                self.scale.full = False
            self.printed = False
            self.t2 = 0
        else:
            if self.t2 != 0:
                self.t2 = 0
            return
    
    def ShowStt(self):
        if len(self.ui.PdID.text()) <= 0 or self.maxwe == 0 or self.minwe == 0:
            return
        if self.dataR <= self.minwe/4 or self.dataR >= self.maxwe+self.minwe*3/4:
            self.index = 0
        elif (self.dataR <= self.minwe/2 and self.dataR >= self.minwe/4) or (self.dataR <= self.maxwe+self.minwe*3/4 and self.dataR >= self.maxwe+self.minwe*1/2):
            self.index = 1
        elif (self.dataR < self.minwe and self.dataR >= self.minwe/2) or (self.dataR <= self.maxwe+self.minwe*1/2 and self.dataR > self.maxwe):
            self.index = 2
        elif self.dataR >= self.minwe and self.dataR <= self.maxwe:
            if self.scale.full:
                if self.k == 0:
                    self.index = 0
                    self.k = 1
                else:
                    self.index = 3
                    self.k = 0
            else:
                self.index = 3
        color = str("background-color: " + self.seq[self.index])
        self.ui.Status.setStyleSheet(color)

    @QtCore.pyqtSlot()
    def UpdateData(self):
        self.loaddata.updatedata()

    @QtCore.pyqtSlot()
    def CheckStamp(self):
        if len(self.ui.PdID.text()) <= 0:
            self.in4(infor = 'Quét mã sản phẩm')
            self.ui.ProductID.setFocus()
            return
        self.data_print = self.loaddata.data_print
        self.print.printdata()

    @QtCore.pyqtSlot()
    def ArduinoCon(self):
        if not self.scale._run_flag:
            self.in4('Kiểm tra kết nối')
            return
        self.scale.scale.serial.close()
        self.scale.board.shutdown()
        self.scale.openPort()

    @QtCore.pyqtSlot()
    def ZeroRet(self):
        if not self.scale._run_flag:
            self.in4('Kiểm tra kết nối')
            return
        self.scale.zeroret = True

    @QtCore.pyqtSlot()
    def Calib(self):
        if len(self.ui.PdID.text()) <= 0:
            self.in4(infor = 'Quét mã sản phẩm')
            self.ui.ProductID.setFocus()
            return
        self.qTimer1.stop()
        self.qTimer2.stop()
        self.qTimer3.stop()
        self.calib = CalibTab(self)

    @QtCore.pyqtSlot()
    def Calendar(self):
        if len(self.ui.PdID.text()) <= 0:
            self.ui.Calendar.setChecked(0)
            self.in4(infor = 'Quét mã sản phẩm')
            self.ui.ProductID.setFocus()
            return
        self.calendar = Calendar(self)

    @QtCore.pyqtSlot()
    def ProductID(self):
        pdid = self.ui.ProductID.text().strip().lower()
        self.resetdata()
        if not pdid:
            return self.in4(infor = 'Quét mã sản phẩm')
        self.data = self.loaddata.data.filter(pl.col('Code Item') == pdid)
        if self.data.is_empty():
            return self.in4(infor = 'Không có mã sản phẩm này')
        self.minwe, self.maxwe = 20.0, 24.0
        self.count = 0
        self.ShowPO()
        self.ui.PdID.setText(pdid)
        self.ui.INTiD.setText(str(self.data['INT'].item()))
        self.ui.CATiD.setText(str(self.data['CAT'].item()))
        self.ui.ProductName.setText(str(self.data['Name'].item()))
        self.ui.MinWe.setText(str(self.minwe))
        self.ui.MaxWe.setText(str(self.maxwe))

    def closeEvent(self, event = None):
        try:
            self.scale.stop()
        except:
            pass
        os.system('taskkill /F /IM EXCEL.exe')
        file_name =  os.path.basename(sys.argv[0])
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            os.system("taskkill /F /IM " + file_name)
        else:
            os.system("taskkill /F /IM python.exe")
        sys.exit()
    
if __name__ == "__main__":
    try:
        data = GetConfig()
        app = QtWidgets.QApplication(sys.argv)
        gui = MainGUI(data.config)
    except:
        logging.exception('Error')
    sys.exit(app.exec_())