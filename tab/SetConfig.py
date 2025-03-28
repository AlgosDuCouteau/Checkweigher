from PyQt5 import QtWidgets, QtCore
from UI import Ui_SetConfig

class SetConfig(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.SetConfig = QtWidgets.QWidget()
        self.SetConfig.setWindowFlags(self.SetConfig.windowFlags() & ~(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinMaxButtonsHint))
        self.ui = Ui_SetConfig()
        self.ui.setupUi(self.SetConfig)
        self.MainWin = MainWin
        self.SetConfig.show()
        self.setDefault()
        self.change = False
        self.ui.Confirm.clicked.connect(self.Confirm)
        self.ui.scale_Confirm.clicked.connect(lambda: self.changeHandle(self.ui.scale.currentText(), sender=self.ui.scale_Confirm))
        self.ui.arduino_Confirm.clicked.connect(lambda: self.changeHandle(self.ui.arduino.currentText(), sender=self.ui.arduino_Confirm))
        self.ui.conveyor_Confirm.clicked.connect(lambda: self.changeHandle(self.ui.conveyor.currentText(), sender=self.ui.conveyor_Confirm))
        self.ui.light_Confirm.clicked.connect(lambda: self.changeHandle(self.ui.light.currentText(), sender=self.ui.light_Confirm))
        self.ui.range_Confirm.clicked.connect(self.setRange)

    def setRange(self):
        if self.ui.range.currentText() == '1/2':
            self.MainWin.range = 0.5
        elif self.ui.range.currentText() == '1/4':
            self.MainWin.range = 0.25

    def setDefault(self):
        self.ui.scale_def.setText(str(self.MainWin.portScale))
        self.ui.arduino_def.setText(str(self.MainWin.portArduino))
        self.ui.conveyor_def.setText(str(self.MainWin.Ard2Convey))
        self.ui.light_def.setText(str(self.MainWin.Ard2Light))
        self.ui.scale.setCurrentText(str(self.MainWin.portScale))
        self.ui.arduino.setCurrentText(str(self.MainWin.portArduino))
        self.ui.conveyor.setCurrentText(str(self.MainWin.Ard2Convey))
        self.ui.light.setCurrentText(str(self.MainWin.Ard2Light))
        self.ui.range.setCurrentIndex(1) if self.MainWin.range == 0.5 else self.ui.range.setCurrentIndex(0)

    def changeHandle(self, text, sender):
        # Set change to True if any value is different from the default
        if sender == self.ui.scale_Confirm:
            if text != self.MainWin.portScale:
                self.change = True
        elif sender == self.ui.arduino_Confirm:
            if text != self.MainWin.portArduino:
                self.change = True
        elif sender == self.ui.conveyor_Confirm:
            if text != str(self.MainWin.Ard2Convey):
                self.change = True
        elif sender == self.ui.light_Confirm:
            if text != str(self.MainWin.Ard2Light):
                self.change = True
                
        # Check if all values are back to default
        if (self.ui.scale.currentText() == str(self.MainWin.portScale) and
            self.ui.arduino.currentText() == str(self.MainWin.portArduino) and
            self.ui.conveyor.currentText() == str(self.MainWin.Ard2Convey) and
            self.ui.light.currentText() == str(self.MainWin.Ard2Light)):
            self.change = False  # Set change to False if all are default

    def Confirm(self):
        if self.change:
            # Prepare the new port settings to write to ports.txt
            new_ports = [
                f"Weight scale to PC: {self.ui.scale.currentText()}",
                f"Arduino to PC: {self.ui.arduino.currentText()}",
                f"Arduino to Conveyor: {self.ui.conveyor.currentText()}",
                f"Arduino to Light: {self.ui.light.currentText()}"
            ]
            # Read the last three lines from ports.txt
            with open(self.MainWin.filePort, 'r') as file:
                lines = file.readlines()
            last_three_lines = lines[-3:]  # Keep the last three lines

            # Write the new settings and the last three lines back to ports.txt
            with open(self.MainWin.filePort, 'w') as file:
                file.write('\n'.join(new_ports) + '\n')
                file.writelines(last_three_lines)

            self.MainWin.in4('Khởi động lại chương trình!', 'Thông tin')
            self.MainWin.close()
        self.SetConfig.close()
