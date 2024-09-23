import sys, os, polars as pl, logging, inspect, traceback, time
from PyQt5 import QtWidgets, QtCore
from UI import Ui_CheckWeigher
from tab import *
from Function import *

# Helper function to get current file name and line number
def get_file_and_line():
    frame = inspect.currentframe().f_back
    return f"{os.path.basename(frame.f_code.co_filename)}:{frame.f_lineno}"

# Configure logging to write errors to errors.txt in the same directory as CheckWeigher.py
logging.basicConfig(filename="errors.txt",
                    format='\n%(asctime)s %(message)s',
                    filemode='a')

class WorkerSignals(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    error = QtCore.pyqtSignal(tuple)
    result = QtCore.pyqtSignal(object)

class Worker(QtCore.QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @QtCore.pyqtSlot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)
        finally:
            self.signals.finished.emit()

class MainGUI(QtWidgets.QMainWindow):
    def __init__(self, data: dict, parent=None):
        super().__init__(parent)
        self.ui = Ui_CheckWeigher()
        self.ui.setupUi(self)
        self.show()
        self.threadpool = QtCore.QThreadPool()
        print(f"Multithreading with maximum {self.threadpool.maxThreadCount()} threads")
        
        # Initialize data and configurations
        fileData = data['fileData']
        self.filePort = data['filePort']
        self.filePrint = data['filePrint']
        fileUpdateData = data['GdriveData']
        fileUpdatePrint = data['GdrivePrint']
        self.portScale = data['portScale']
        self.portArduino = data['portArduino']
        self.Ard2Convey = data['Ard2Convey']
        self.Ard2Light = data['Ard2Light']
        self.Delay2Print = data['Delay2Print']
        self.Quan2Print = data['Quan2Print']
        
        # Initialize components
        self.loadData = LoadData(self, fileData=fileData, filePrint=self.filePrint, fileUpdateData=fileUpdateData, fileUpdatePrint=fileUpdatePrint)
        self.fileProcess = FileProcess(self, filePrint=self.filePrint)
        self.scale = GetScale(portScale=self.portScale, portArduino=self.portArduino, Ard2Convey=self.Ard2Convey, Ard2Light=self.Ard2Light)
        t = time.time()
        self.scale.openPort()
        print(time.time() - t)
        if self.scale._run_flag:
            self.in4(infor = 'Kết nối thành công!', title = 'Thông tin')
            self.start_scale_worker()
        else:
            self.in4(infor = 'Không thể kết nối đến cân!')
        
        # Connect UI elements to their respective functions
        self.ui.MachineID.setCurrentIndex(-1)
        self.ui.MachineID.currentTextChanged.connect(lambda _: setattr(self.fileProcess, 'changeEvent', True))
        self.ui.actionConfig.triggered.connect(self.setConfig)
        self.ui.UpdateData.clicked.connect(self.UpdateData)
        self.ui.ShowStamp.clicked.connect(self.ShowStamp)
        self.ui.ChangePd.clicked.connect(self.ChangePd)
        self.ui.ArduinoCon.clicked.connect(self.ArduinoCon)
        self.ui.ZeroRet.clicked.connect(self.ZeroRet)
        self.ui.Calib.clicked.connect(self.Calib)
        self.ui.Calendar.clicked.connect(self.Calendar)
        self.ui.ProductID.returnPressed.connect(self.ProductID)
        self.ui.ProductID.setFocus()
        
        # Set up timers for various periodic tasks
        self.setupTimers()
        
        # Initialize variables
        self.count, self.index, self.k, self.t1, self.t2, self.maxwe, self.minwe,  self.dataR = 0, 0, 0, 0, 0, 0, 0, 0
        self.printed = False
        self.t = 0
        self.seq = ['rgb(255, 255, 255)', 'rgb(255, 0, 0)', 'rgb(255, 255, 0)', 'rgb(0, 255, 0)']
        self.filterlist = []

    def start_scale_worker(self):
        self.scale.data_ready.connect(self.ShowDataR)  # Connect the signal to ShowDataR
        worker = Worker(self.scale.read_control_continuously)
        self.threadpool.start(worker)

    def setupTimers(self):
        # Timer for updating status display
        self.qTimer2 = QtCore.QTimer()
        self.qTimer2.setInterval(100)
        self.qTimer2.timeout.connect(self.ShowStt)
        self.qTimer2.start()
        
        # Timer for comparing weight and triggering actions
        self.qTimer3 = QtCore.QTimer()
        self.qTimer3.setInterval(100)
        self.qTimer3.timeout.connect(self.Compare)
        self.qTimer3.start()
        
        # Timer for filtering weight data
        self.qTimer4 = QtCore.QTimer()
        self.qTimer4.setInterval(200)
        self.qTimer4.timeout.connect(self.dataRfilter)
        self.qTimer4.start()

    def in4(self, infor:str, title = 'Lỗi'):
        # Display an information or error message box
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
        # Clear all input fields and reset display
        self.ui.ProductID.clear()
        self.ui.PdID.clear()
        self.ui.CATiD.clear()
        self.ui.INTiD.clear()
        self.ui.ProductName.clear()
        self.ui.MaxWe.clear()
        self.ui.MinWe.clear()
        self.ui.QuanBox.clear()
        self.ui.QuanStampAuto.clear()
        self.ui.Calendar.setChecked(0)

    def dataRfilter(self):
        # Filter weight data to detect sudden changes
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

    @QtCore.pyqtSlot(float)
    def ShowDataR(self, value):
        self.dataR = value
        self.ui.CurrentWe.display(float(self.dataR))

    def ShowPO(self):
        # Update quantity displays
        self.ui.QuanBox.setText(str(int(self.count)))
        self.ui.QuanStampAuto.setText(str(int(self.count)))

    def Compare(self):
        # Compare weight and trigger printing if within range
        if len(self.ui.PdID.text()) <= 0 or self.maxwe == 0 or self.minwe == 0:
            return
        if self.t == 0:
            self.ui.Status.setText("")
        if self.dataR <= self.maxwe and self.dataR >= self.minwe and not self.printed and not self.scale.full:
            if self.t < self.Delay2Print:
                self.t += 1
                self.ui.Status.setText(str(float((self.Delay2Print-self.t)/10)))
            if self.t >= self.Delay2Print:
                self.t = 0
                self.fileProcess.printSheet()
                self.printed = True
                self.scale.full = True
                self.count += 1
                self.ShowPO()
        elif self.dataR < self.minwe/11.5:
            self.t = 0
            if self.scale.full:
                self.scale.full = False
            if self.printed:
                self.printed = False
        else:
            if self.t != 0:
                self.t = 0
    
    def ShowStt(self):
        # Update status display based on weight
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
    def setConfig(self):
        # Open configuration window
        self.setConfig = SetConfig(self)

    @QtCore.pyqtSlot()
    def UpdateData(self):
        # Update data from source
        try:
            self.loadData.updatedata()
            self.in4(infor = 'Hoàn thành cập nhật', title = 'Thông tin')
        except Exception as e:
            self.in4(infor = 'Lỗi dữ liệu, không thể cập nhật!')
            return

    @QtCore.pyqtSlot()
    def ShowStamp(self):
        if len(self.ui.PdID.text()) <= 0 or self.ui.MachineID.currentIndex() == -1:
            self.in4(infor = 'Quét mã sản phẩm và chọn số máy!')
            self.ui.ProductID.setFocus()
            return
        if self.fileProcess.is_processing:
            self.in4(infor = 'Đang xử lý dữ liệu. Vui lòng đợi.', title = 'Thông tin')
            return
        # Use a worker to generate an image
        worker = Worker(self.fileProcess.generateImage, self.data)
        worker.signals.result.connect(self.on_image_generated)
        worker.signals.error.connect(self.on_image_generation_error)
        self.threadpool.start(worker)

    def on_image_generated(self, imgFile):
        if imgFile:
            self.showStamp = ShowStamp(self, imgFile)
        else:
            self.in4(infor = 'Không thể tạo hình ảnh', title = 'Lỗi')

    def on_image_generation_error(self, error):
        error_type, error_instance, traceback = error
        self.in4(infor = f'Lỗi khi tạo hình ảnh: {error_instance}')

    @QtCore.pyqtSlot()
    def ChangePd(self):
        # Change product
        if len(self.ui.PdID.text()) <= 0 or self.ui.MachineID.currentIndex() == -1:
            self.in4(infor = 'Quét mã sản phẩm và chọn số máy!')
            self.ui.ProductID.setFocus()
            return
        if not self.scale._run_flag:
            self.in4('Kiểm tra kết nối!')
            return
        self.scale.light = True
        self.qTimer2.stop()
        self.qTimer3.stop()
        self.qTimer4.stop()
        self.changepd = ChangePd(self)

    @QtCore.pyqtSlot()
    def ArduinoCon(self):
        # Reconnect to Arduino
        if not self.scale._run_flag:
            self.in4('Kiểm tra kết nối!')
            return
        self.scale.scale.serial.close()
        self.scale.board.shutdown()
        self.scale.openPort()
        if self.scale._run_flag:
            self.in4(infor = 'Kết nối thành công!', title = 'Thông báo')
            self.start_scale_worker()
        else:
            self.in4(infor = 'Không thể kết nối đến cân!')

    @QtCore.pyqtSlot()
    def ZeroRet(self):
        # Reset scale to zero
        if not self.scale._run_flag:
            self.in4('Kiểm tra kết nối!')
            return
        self.scale.zeroret = True

    @QtCore.pyqtSlot()
    def Calib(self):
        # Open calibration window
        if len(self.ui.PdID.text()) <= 0 or self.ui.MachineID.currentIndex() == -1:
            self.in4(infor = 'Quét mã sản phẩm và chọn số máy!')
            self.ui.ProductID.setFocus()
            return
        if self.fileProcess.changeEvent:
            self.in4(infor = 'Nhấn nút "Kiểm tra tem" trước!', title = 'Lỗi')
            return
        self.qTimer2.stop()
        self.qTimer3.stop()
        self.qTimer4.stop()
        self.calib = CalibTab(self)

    @QtCore.pyqtSlot()
    def Calendar(self):
        # Open calendar window
        if len(self.ui.PdID.text()) <= 0 or self.ui.MachineID.currentIndex() == -1:
            self.in4(infor = 'Quét mã sản phẩm và chọn số máy!')
            self.ui.Calendar.setChecked(0)
            self.ui.ProductID.setFocus()
            return
        self.calendar = Calendar(self)

    @QtCore.pyqtSlot()
    def ProductID(self):
        # Process product ID input
        def add_newline_every_7_spaces(s):
            blank_count = 0
            result = ""
            for char in s:
                if char == " ":
                    blank_count += 1
                    if blank_count % 6 == 0:
                        result += "\n"
                result += char
            return result
        pdid = self.ui.ProductID.text().strip().lower()
        self.resetdata()
        if not pdid:
            return self.in4(infor = 'Quét mã sản phẩm!')
        self.data = self.loadData.data.filter(pl.col('Code_Item') == pdid)
        if self.data.is_empty():
            return self.in4(infor = 'Không có mã sản phẩm này')
        try:
            self.ui.INTiD.setText(str(self.data['INT'].item()))
            self.ui.CATiD.setText(str(self.data['CAT'].item()))
            self.minwe, self.maxwe = 20.0, 24.0
            self.count = 0
            self.countManual = 0
            self.ShowPO()
            self.ui.PdID.setText(pdid)
            self.ui.ProductName.setText(add_newline_every_7_spaces(str(self.data['Name'].item())))
            self.ui.MinWe.setText(str(self.minwe))
            self.ui.MaxWe.setText(str(self.maxwe))
            self.fileProcess.changeEvent = True
        except Exception as e:
            self.in4(infor = 'Lỗi dữ liệu, xem lại mã hàng này')
            return

    def closeEvent(self, event = None):
        # Handle application close event
        if self.scale._run_flag:
            self.scale.stop()
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
    except Exception as e:
        # Check if this error has already been logged
        logging.exception(f'Error initializing application at {get_file_and_line()}: {e}')
    sys.exit(app.exec_())