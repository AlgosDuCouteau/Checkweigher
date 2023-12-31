import os, polars as pl, xlwings as xw
from pathlib import Path
from datetime import datetime, timedelta, date
from SubFunction import IDAutomation_Uni_C128, IDAutomation_Uni_C39

class Print():
    def __init__(self, MainWin, filePrint: Path):
        self.MainWin = MainWin
        self.filePrint = filePrint
        self.moa = ['A', 'B', 'C', 'D',
                    'E', 'F', 'G', 'H',
                    'J', 'K', 'L', 'M', 
                    'N', 'P', 'Q', 'R',
                    'S', 'T', 'U']

    def loaddata(self) -> pl.DataFrame:
        data = self.MainWin.data
        data_print = self.MainWin.data_print
        return data, data_print

    def essentialdata(self, data: pl.DataFrame, data_print: pl.DataFrame):
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
        info = dte+mo+str(shift)+data['Mã 1'].item()
        year_month = ymdn[:-3].rstrip()
        lot = ymdn[2]+ymdn[3]+ymdn[5]+ymdn[6]+data['Mã 1'].item()
        layouts = data_print.filter(pl.col('Types')==data['Type'].item())
        layouts = layouts[[s.name for s in layouts if not (s.null_count() > 0)]]
        return data, layouts, info, year_month, lot
    
    def removespc(self, text):
        text = str(text)
        spc = '#$%()^*&'
        for char in spc:
            text = text.replace(char, '')
        return str(text)
    
    def appenddata(self, data: pl.DataFrame, layout: pl.DataFrame, info: str, year_month: str, lot: str):
        os.system('taskkill /f /IM EXCEL.exe')
        with xw.App(visible=False) as app:
            wb = xw.Book(self.filePrint)
            ws = wb.sheets[data['Type'].item()]
            ws[layout['info'].item()].value = info
            ws[layout['int'].item()].value = data['INT'].item()
            ws[layout['lot'].item()].value = lot
            ws[layout['name'].item()].value = data['Name'].item()
            ws[layout['year_month'].item()].value = year_month
            ws[layout['quantity'].item()].value = str(int(data['Quantity'].item()))+' PR'
            ws[layout['cat'].item()].value = data['CAT'].item()
            ws[layout['size'].item()].value = data['Size'].item()
            ws[layout['product_of'].item()].value = data['ProductOf'].item()
            barcode = IDAutomation_Uni_C128(self.removespc(data['Mã vạch thùng đầu'].item()+lot+data['Mã vạch thùng đuôi'].item()))
            ws[layout['barcode'].item()].value = barcode
            barcode = ''
            ws[layout['code'].item()].value = data['Mã vạch thùng đầu'].item()+lot+data['Mã vạch thùng đuôi'].item()
            if layout.width > 12:
                if layout['Types'].item() == 'TemThung2MaVachKoSo':
                    ws[layout['cat_code'].item()].value = '('+str(int(data['Quantity'].item()))+') '+data['CAT'].item()
                    barcode = IDAutomation_Uni_C39(self.removespc(data['CAT'].item()))
                    ws[layout['cat_barcode'].item()].value = barcode
                    barcode = ''
                elif layout['Types'].item() == 'TemThung2MaVachCoSo':
                    ws[layout['quan_cat_code'].item()].value = '('+str(int(data['Quantity'].item()))+') '+data['CAT'].item()
                    barcode = IDAutomation_Uni_C39(self.removespc(str(int(data['Quantity'].item()))+' '+data['CAT'].item()))
                    ws[layout['quan_cat_barcode'].item()].value = barcode
                    barcode = ''
                elif layout['Types'].item() == 'TemThung2MaVachMoi':
                    ws[layout['quan_cat_code'].item()].value = '('+str(int(data['Quantity'].item()))+') '+data['Mã vạch nhỏ'].item()
                    barcode = IDAutomation_Uni_C39(self.removespc(str(int(data['Quantity'].item()))+' '+data['Mã vạch nhỏ'].item()))
                    ws[layout['quan_cat_barcode'].item()].value = barcode
                    barcode = ''
                    ws[layout['int1'].item()].value = data['INT'].item()
            ws.api.PrintOut()
            wb.save()
            wb.close()
        os.system('taskkill /F /IM EXCEL.exe')

    def printdata(self):
        data, data_print = self.loaddata()
        if data['Type'].item() == 'OEM':
            return
        data, layouts, info, year_month, lot = self.essentialdata(data, data_print)
        if layouts.height==0:
            return
        self.appenddata(data, layouts, info, year_month, lot)