from PyQt5 import QtCore
from telemetrix import telemetrix
from minimalmodbus import Instrument, serial
import time, logging, os, inspect, sys

# Helper function to get current file name and line number
def get_file_and_line():
    frame = inspect.currentframe().f_back
    return f"{os.path.basename(frame.f_code.co_filename)}:{frame.f_lineno}"

def resource_path():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return base_path

# Configure logging
logging.basicConfig(filename=os.path.join(resource_path(), "errors.txt"),
                    format='\n%(asctime)s %(message)s',
                    filemode='a')

class GetScale(QtCore.QObject):
    data_ready = QtCore.pyqtSignal(float)  # New signal to emit scale readings

    def __init__(self, portScale: str, portArduino: str, Ard2Convey: int, Ard2Light: int, parent=None):
        super().__init__(parent)
        # Initialize flags and parameters for the scale and Arduino
        self._run_flag = True
        self.cps = False
        self.full = False
        self.zeroret = False
        self.light = False
        self.light_on = False
        self.portScale = portScale
        self.portArduino = portArduino
        self.Ard2Convey = Ard2Convey
        self.Ard2Light = Ard2Light
        self.dataR = 0
        self.tt = 0

    def openPort(self):
        # Attempt to open the scale and Arduino ports
        time.sleep(0.5)
        try:
            self.scale = Instrument(self.portScale, 1)  # Initialize the scale instrument
            self.scale.serial.parity = serial.PARITY_EVEN
            self.scale.serial.timeout  = 0.5
            self.scale.clear_buffers_before_each_transaction = True
        except serial.serialutil.SerialException as e:  # Capture the exception
            logging.error(f"Could not open port '{self.portScale}' at {get_file_and_line()}: {e}")  # Log the error with line number
            self._run_flag = False
            return
        try:
            self.board = telemetrix.Telemetrix(com_port = self.portArduino, arduino_wait = 6)  # Initialize Arduino
            self.board.set_pin_mode_digital_output(self.Ard2Convey)  # Set pin modes
            self.board.set_pin_mode_digital_output(self.Ard2Light)
            self.board.digital_write(self.Ard2Convey, 0)  # Initialize outputs to low
            self.board.digital_write(self.Ard2Light, 0)
        except serial.serialutil.SerialException as e:
            logging.error(f"Could not open port '{self.portArduino}' at {get_file_and_line()}: {e}")  # Log the error with line number
            self._run_flag = False
            return
        self._run_flag = True

    def readScale(self):
        # Read data from the scale if the serial port is open
        if self.scale.serial.is_open:
            try:
                value = self.scale.read_registers(0, 2, 3)  # Read registers from the scale
            except Exception as e:
                logging.error(f"Cant read from scale at {get_file_and_line()}: {e}")
                return
            # Process the read values to calculate dataR
            if value[0] != 0:
                dataR = -(value[0] - value[1]+1)/100
            else:
                dataR = value[1]/100
            # Handle specific conditions for dataR
            if dataR <= 0.25 and dataR != 0:
                if dataR < 0:
                    dataR = 0
                    try:
                        self.scale.write_register(7, 42254, functioncode = 6)  # Write to scale to reset
                    except Exception as e:
                        logging.error(f"Cant write to scale at {get_file_and_line()}: {e}")
                        return
                    time.sleep(0.5)
                self.tt += 1
                if self.tt == 60:
                    dataR = 0
                    try:
                        self.scale.write_register(7, 42254, functioncode = 6)  # Write to scale to reset
                    except Exception as e:
                        logging.error(f"Cant write to scale at {get_file_and_line()}: {e}")
                        return
                    time.sleep(0.5)
                    self.tt = 0
            else:
                self.tt = 0
            self.dataR = dataR  # Store the read data
            
    def calibScale(self):
        # Calibrate the scale if zeroret flag is set
        if self.zeroret:
            try:
                self.scale.write_register(7, 42254, functioncode = 6)  # Write to scale to reset
            except Exception as e:
                logging.error(f"Cant write to scale at {get_file_and_line()}: {e}")
                return
            self.zeroret = False

    def runConvey(self):
        # Control the conveyor based on the full flag
        if self.full:
            if not self.cps:
                try:
                    self.board.digital_write(self.Ard2Convey, 1)  # Turn on conveyor
                    self.cps = True
                except Exception as e:
                    logging.error(f"Cant write to arduino at {get_file_and_line()}: {e}")
                    self.board.shutdown()  # Shutdown Arduino on error
                    self.scale.serial.close()  # Close scale serial
                    self.openPort()  # Attempt to reopen ports
                    return
            self.t = time.time()  # Record the time the conveyor was turned on
        elif not self.full and self.cps:
            if time.time()-self.t>15:  # Turn off conveyor after 15 seconds
                try:
                    self.board.digital_write(self.Ard2Convey, 0)  # Turn off conveyor
                    self.cps = False
                except Exception as e:
                    logging.error(f"Cant write to arduino at {get_file_and_line()}: {e}")
                    self.board.shutdown()  # Shutdown Arduino on error
                    self.scale.serial.close()  # Close scale serial
                    self.openPort()  # Attempt to reopen ports
                    return

    def runLight(self):
        # Control the light based on the light flag
        if self.light:
            if not self.light_on:
                try:
                    self.board.digital_write(self.Ard2Light, 1)  # Turn on light
                    self.light_on = True
                except Exception as e:
                    logging.error(f"Cant write to arduino at {get_file_and_line()}: {e}")
                    self.board.shutdown()  # Shutdown Arduino on error
                    self.scale.serial.close()  # Close scale serial
                    self.openPort()  # Attempt to reopen ports
                    return
        elif not self.light and self.light_on:
            try:
                self.board.digital_write(self.Ard2Light, 0)  # Turn off light
                self.light_on = False
            except Exception as e:
                logging.error(f"Cant write to arduino at {get_file_and_line()}: {e}")
                self.board.shutdown()  # Shutdown Arduino on error
                self.scale.serial.close()  # Close scale serial
                self.openPort()  # Attempt to reopen ports
                return

    def read_control_continuously(self):
        while self._run_flag:
            self.readScale()
            self.calibScale()  # Include calibration in the continuous reading loop
            self.runConvey()
            self.runLight()
            self.data_ready.emit(self.dataR)  # Emit the current reading
            time.sleep(0.1)

    def stop(self):
        self._run_flag = False
        self.board.shutdown()
        self.scale.serial.close()
