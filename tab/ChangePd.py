from PyQt5 import QtWidgets, QtCore
from UI import Ui_ChangePD

class ChangePd(QtWidgets.QWidget):
    def __init__(self, MainWin, parent = None):
        super().__init__(parent)
        self.ChangePD = QtWidgets.QWidget()
        self.ChangePD.setWindowFlags(self.ChangePD.windowFlags() & ~(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinMaxButtonsHint))
        self.ui = Ui_ChangePD()
        self.ui.setupUi(self.ChangePD)
        self.MainWin = MainWin
        self.ui.Confirm.clicked.connect(self.Confirm)
        self.ChangePD.show()
        self.MainWin.hide()
        self.ui.Product.setText(self.add_newline(str(self.MainWin.data['Name'].item())))
        self.elapsed_time = 0  # Initialize elapsed time
        minutes = self.elapsed_time // 60  # Calculate minutes
        seconds = self.elapsed_time % 60  # Calculate seconds
        self.ui.Timer.setText(f"{minutes:02}:{seconds:02}")  # Format as MM:SS
        self.timer = QtCore.QTimer(self)  # Create a timer
        self.timer.timeout.connect(self.Timer)  # Connect the timeout signal to the Timer method
        self.timer.start(1000)  # Start the timer with a 1-second interval

    def Timer(self):
        self.elapsed_time += 1
        minutes = self.elapsed_time // 60  # Calculate minutes
        seconds = self.elapsed_time % 60  # Calculate seconds
        self.ui.Timer.setText(f"{minutes:02}:{seconds:02}")  # Format as MM:SS
        

    def add_newline(self,s):
        blank_count = 0
        result = ""
        for char in s:
            if char == " ":
                blank_count += 1
                if blank_count % 5 == 0:
                    result += "\n"
            result += char
        return result
        
    def Confirm(self):
        self.timer.stop()
        self.MainWin.show()
        self.MainWin.scale.light = False
        self.MainWin.qTimer1.start()
        self.MainWin.qTimer2.start()
        self.MainWin.qTimer3.start()
        self.ChangePD.close()