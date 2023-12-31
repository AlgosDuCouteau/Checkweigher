from PyQt5 import QtCore, QtWidgets
from telemetrix import telemetrix
from minimalmodbus import Instrument, serial, NoResponseError
import time, logging

logging.basicConfig(filename="errors.txt",
                    format='\n%(asctime)s %(message)s',
                    filemode='a')

class GetScale(QtCore.QThread):
    def __init__(self, portScale: str, portArduino: str, Ard2Convey: int, parent = None):
        super().__init__(parent)
        self._run_flag = True
        self.cps = False
        self.full = False
        self.zeroret = False
        self.portScale = portScale
        self.portArduino = portArduino
        self.Ard2Convey = Ard2Convey
        self.dataR = 0
        self.tt = 0
        self.openPort()

    def in4(self, infor:str, title = 'Lỗi'):
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(infor)
        if title == 'Lỗi':
            msg.setIcon(QtWidgets.QMessageBox.Warning)
            logging.exception("Failed")
        elif title == 'Thông tin':
            msg.setIcon(QtWidgets.QMessageBox.Information)
        msg.exec_()
        return
    
    def openPort(self):
        time.sleep(0.5)
        try:
            self.scale = Instrument(self.portScale, 1)
            self.scale.serial.parity = serial.PARITY_EVEN
        except serial.serialutil.SerialException:
            self.in4('Không có kết nối tới cân')
            self._run_flag = False
            return
        try:
            self.board = telemetrix.Telemetrix(com_port = self.portArduino, arduino_wait = 6)
            self.board.set_pin_mode_digital_output(self.Ard2Convey)
            self.board.digital_write(self.Ard2Convey, 0)
        except serial.serialutil.SerialException:
            self.in4('Không có kết nối tới băng tải')
            self._run_flag = False
            return
        self.in4(infor = 'Kết nối thành công', title = 'Thông tin')
        self._run_flag = True

    def readScale(self):
        if self.scale.serial.is_open:
            try:
                value = self.scale.read_registers(0, 2, 3)
            except NoResponseError:
                logging.exception("Cant read from scale")
                return
            if value[0] != 0:
                dataR = -(value[0] - value[1]+1)/100
            else:
                dataR = value[1]/100
            if dataR <= 0.25 and dataR != 0:
                if dataR < 0:
                    dataR = 0
                    try:
                        self.scale.write_register(7, 42254, functioncode = 6)
                    except:
                        logging.exception("Cant write to scale")
                        return
                    time.sleep(0.1)
                self.tt += 1
                if self.tt == 60:
                    dataR = 0
                    try:
                        self.scale.write_register(7, 42254, functioncode = 6)
                    except:
                        logging.exception("Cant write to scale")
                        return
                    time.sleep(0.1)
                    self.tt = 0
            else:
                self.tt = 0
            self.dataR = dataR
            
    def calibScale(self):
        if self.zeroret:
            try:
                self.scale.write_register(7, 42254, functioncode = 6)
            except:
                logging.exception("Cant write to scale")
                return
            self.zeroret = False

    def runConvey(self):
        if self.full:
            if not self.cps:
                try:
                    self.board.digital_write(self.Ard2Convey, 1)
                    self.cps = True
                except:
                    logging.exception("Cant write to arduino")
                    self.board.shutdown()
                    self.scale.serial.close()
                    self.openPort()
                    return
            self.t = time.time()
        elif not self.full and self.cps:
            if time.time()-self.t>15:
                try:
                    self.board.digital_write(self.Ard2Convey, 0)
                    self.cps = False
                except:
                    logging.exception("Cant write to arduino")
                    self.board.shutdown()
                    self.scale.serial.close()
                    self.openPort()
                    return

    def run(self):
        while self._run_flag:
            self.readScale()
            self.calibScale()
            self.runConvey()
            time.sleep(0.1)
    
    def stop(self):
        self._run_flag = False
        self.board.shutdown()
        self.scale.serial.close()