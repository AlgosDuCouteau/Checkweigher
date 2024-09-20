import os, polars as pl, xlwings as xw, logging, inspect, sys, traceback, pythoncom, win32com.client, win32clipboard
from pathlib import Path
from datetime import datetime, timedelta, date
from SubFunction import IDAutomation_Uni_C128, IDAutomation_Uni_C39
from PyQt5 import QtCore
from PIL import ImageGrab, Image

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

class FileProcess(QtCore.QObject):
    finished = QtCore.pyqtSignal()
    error = QtCore.pyqtSignal(tuple)

    def __init__(self, MainWin, filePrint: Path):
        super().__init__()
        self.MainWin = MainWin
        self.filePrint = filePrint
        self.data_print = self.MainWin.loadData.data_print
        self.moa = ['A', 'B', 'C', 'D',
                    'E', 'F', 'G', 'H',
                    'J', 'K', 'L', 'M', 
                    'N', 'P', 'Q', 'R',
                    'S', 'T', 'U']
        self.is_processing, self.printable = False, False
    
    def removeSpc(self, text):
        text = str(text)
        spc = '#$%()^*&'
        for char in spc:
            text = text.replace(char, '')
        return str(text)
    
    @QtCore.pyqtSlot(object)
    def appendData(self, data: pl.DataFrame):
        if self.is_processing:
            self.error.emit(("ProcessingError", "Excel file is currently being processed. Please wait."))
            return

        self.is_processing = True
        
        try:
            if data['Dinh_dang_Tem_Thung'].item() == 'OEM':
                self.printable = False
                return
            self.ws = data['Dinh_dang_Tem_Thung'].item()
            curr_time = datetime.now()
            if curr_time.hour >= 6 and curr_time.hour < 14:
                shift = 1
            elif curr_time.hour >= 14 and curr_time.hour < 22:
                shift = 2
            else:
                shift = 3
            if self.MainWin.ui.Calendar.isChecked() == 0:
                if curr_time.hour >= 6 and curr_time.hour < 24:
                    ymdn = str(date.today())
                else:
                    ymdn = str(date.today() - timedelta(days = 1))
            else:
                ymdn = self.MainWin.ui.DD_MM_YY.text()
            dte = ymdn[2]+ymdn[3]+ymdn[5]+ymdn[6]+ymdn[8]+ymdn[9]
            if self.MainWin.ui.MachineID.currentText() == 'HT1':
                mo = self.moa[16]
            elif self.MainWin.ui.MachineID.currentText() == 'HT2':
                mo = self.moa[17]
            elif self.MainWin.ui.MachineID.currentText() == 'OEM':
                mo = self.moa[18]
            else:
                mo = self.moa[int(self.MainWin.ui.MachineID.currentText())-1]
            info = dte+mo+str(shift)+data['Ma_1'].item()
            self.info = info
            year_month = ymdn[:-3].rstrip()
            lot = ymdn[2]+ymdn[3]+ymdn[5]+ymdn[6]+data['Ma_1'].item()
            layouts = self.data_print.filter(pl.col('Types')==self.ws)
            layout = layouts[[s.name for s in layouts if not (s.null_count() > 0)]]
            os.system('taskkill /f /IM EXCEL.exe')
            pythoncom.CoInitialize()
            with xw.App(visible=False) as app:
                wb = xw.Book(self.filePrint)
                ws = wb.sheets[self.ws]
                ws[layout['info'].item()].value = info
                ws[layout['int'].item()].value = data['INT'].item()
                ws[layout['lot'].item()].value = lot
                ws[layout['name'].item()].value = data['Name'].item()
                ws[layout['year_month'].item()].value = year_month
                ws[layout['quantity'].item()].value = str(int(data['Quantity_thung'].item()))+' '+data['Don_vi_thung'].item()
                ws[layout['cat'].item()].value = data['CAT'].item()
                ws[layout['size'].item()].value = data['Size'].item()
                ws[layout['product_of'].item()].value = data['ProductOf'].item()
                barcode = IDAutomation_Uni_C128(self.removeSpc(data['Ma_vach_thung_dau'].item()+lot+data['Ma_vach_thung_duoi'].item()))
                ws[layout['barcode'].item()].value = barcode
                barcode = ''
                ws[layout['code'].item()].value = data['Ma_vach_thung_dau'].item()+lot+data['Ma_vach_thung_duoi'].item()
                if layout.width > 12:
                    if layout['Types'].item() == 'TemThung2MaVachKoSo':
                        ws[layout['cat_code'].item()].value = '('+str(int(data['Quantity_thung'].item()))+') '+data['CAT'].item()
                        barcode = IDAutomation_Uni_C39(self.removeSpc(data['CAT'].item()))
                        ws[layout['cat_barcode'].item()].value = barcode
                        barcode = ''
                    elif layout['Types'].item() == 'TemThung2MaVachCoSo':
                        ws[layout['quan_cat_code'].item()].value = '('+str(int(data['Quantity_thung'].item()))+') '+data['CAT'].item()
                        barcode = IDAutomation_Uni_C39(self.removeSpc(str(int(data['Quantity_thung'].item()))+' '+data['CAT'].item()))
                        ws[layout['quan_cat_barcode'].item()].value = barcode
                        barcode = ''
                    elif layout['Types'].item() == 'TemThung2MaVachMoi':
                        ws[layout['quan_cat_code'].item()].value = '('+str(int(data['Quantity_thung'].item()))+') '+data['Ma_vach_nho'].item()
                        barcode = IDAutomation_Uni_C39(self.removeSpc(str(int(data['Quantity_thung'].item()))+' '+data['Ma_vach_nho'].item()))
                        ws[layout['quan_cat_barcode'].item()].value = barcode
                        barcode = ''
                        ws[layout['int1'].item()].value = data['INT'].item()
                wb.save()
                wb.close()
        except Exception as e:
            logging.error(f"{get_file_and_line()}: {traceback.format_exc()}")
            self.error.emit((type(e), str(e), traceback.format_exc()))
        finally:
            pythoncom.CoUninitialize()
            self.is_processing = False
            self.finished.emit()
            self.printable = True
            
    @QtCore.pyqtSlot()
    def generateImage(self):
        if self.is_processing:
            self.error.emit(("ProcessingError", "Excel file is currently being processed. Please wait."))
            return None

        self.is_processing = True
        pythoncom.CoInitialize()
        try:
            if not self.printable:
                return None

            # Check if the image already exists
            pdid = self.MainWin.ui.PdID.text()
            output_folder = os.path.join(resource_path(), 'cell_range_images')
            imgFile = os.path.join(output_folder, f'{pdid}_{self.info}.jpg')
            
            if os.path.exists(imgFile):
                return imgFile

            # Get the range from data_print
            cell_range = self.data_print.filter(pl.col('Types')==self.ws)['range'].item()

            # Split the cell range
            start_cell, end_cell = cell_range.split(':')
            
            # Convert column letters to numbers
            def col_to_num(col):
                num = 0
                for c in col:
                    if c.isalpha():
                        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
                return num

            # Extract start and end columns and rows
            start_col = col_to_num(''.join(filter(str.isalpha, start_cell)))
            start_row = int(''.join(filter(str.isdigit, start_cell)))
            end_col = col_to_num(''.join(filter(str.isalpha, end_cell)))
            end_row = int(''.join(filter(str.isdigit, end_cell)))

            o = win32com.client.Dispatch('Excel.Application')
            o.DisplayAlerts = False

            try:
                # Open the workbook
                wb = o.Workbooks.Open(self.filePrint)
                ws = wb.Worksheets(self.ws)

                ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col)).Copy()
                img = ImageGrab.grabclipboard()

                # Resize the image to 600x360
                img = img.resize((920, 520), Image.LANCZOS)

                # Create the output folder if it doesn't exist
                os.makedirs(output_folder, exist_ok=True)

                img.save(imgFile, quality=95)  # Save with high quality

                return imgFile

            finally:
                # Clear clipboard before closing Excel
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.CloseClipboard()

                # Close Excel
                if 'wb' in locals():
                    wb.Close(SaveChanges=False)
                o.Quit()

                os.system('taskkill /f /IM EXCEL.exe')

        except Exception as e:
            self.error.emit((type(e), str(e), traceback.format_exc()))
            return None
        finally:
            pythoncom.CoUninitialize()
            self.is_processing = False
            self.finished.emit()

    @QtCore.pyqtSlot()
    def printSheet(self):
        if self.is_processing:
            self.error.emit(("ProcessingError", "Excel file is currently being processed. Please wait."))
            return

        self.is_processing = True
        try:
            if not self.printable:
                return
            with xw.App(visible=False) as app:
                wb = xw.Book(self.filePrint)
                ws = wb.sheets[self.ws]
                ws.api.PrintOut()
        except Exception as e:
            self.error.emit((type(e), str(e), traceback.format_exc()))
        finally:
            self.is_processing = False
            self.finished.emit()

