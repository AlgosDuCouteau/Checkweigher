import os, polars as pl, numpy as np
from pathlib import Path
from shutil import copyfile
from SubFunction import append_df_to_excel

class LoadData():
    def __init__(self, MainWin, fileData: Path, filePrint: Path, fileUpdateData: Path, fileUpdatePrint: Path):
        self.data = []
        self.data_update = None
        self.MainWin = MainWin
        self.fileData = fileData
        self.fileUpdateData = fileUpdateData
        self.filePrint = filePrint
        self.fileUpdatePrint = fileUpdatePrint
        self.data, self.data_print = self.loaddata()

    def loaddata(self) -> pl.DataFrame:
        self.MainWin.resetdata()
        data = pl.read_excel(source = self.fileData, sheet_name = 'Database', schema_overrides = {'Code Item': str},
                             read_csv_options = {'has_header': True, 'infer_schema_length': None, 'columns': ['Code Item', 'Name', 'ProductOf', 'Quantity thùng', 'CAT', 'INT', 'Size', 'Type', 'Mã 1', 'Mã vạch thùng đầu', 'Mã vạch thùng đuôi', 'Mã vạch nhỏ', 'Đơn vị thùng']}, 
                             xlsx2csv_options = {'skip_empty_lines': True, 'skip_hidden_rows': False})
        data_print = pl.read_excel(source = self.fileData, sheet_name = 'Layouts',
                             read_csv_options = {'has_header': True, 'infer_schema_length': None}, xlsx2csv_options = {'skip_empty_lines': True, 'skip_hidden_rows': False})
        return data, data_print

    def updatedata(self) -> pl.DataFrame:
        os.system('taskkill /f /IM EXCEL.exe')
        copyfile(self.fileUpdatePrint, self.filePrint)
        data_update = pl.read_excel(source = self.fileUpdateData, sheet_name = 'Sheet1', schema_overrides = {'Code Item': str},
                             read_csv_options = {'has_header': True, 'infer_schema_length': None, 'columns': ['Code Item', 'Name', 'ProductOf', 'Quantity thùng', 'CAT', 'INT', 'Size', 'Type', 'Mã 1', 'Mã vạch thùng đầu', 'Mã vạch thùng đuôi', 'Mã vạch nhỏ', 'Đơn vị thùng']}, 
                             xlsx2csv_options = {'skip_empty_lines': True, 'skip_hidden_rows': False})
        data_update = data_update.filter(pl.all_horizontal(pl.col('Code Item').is_not_null()))
        data_update.unique(subset = ['Code Item'])
        # minwe = np.empty(len(df1), dtype=np.int)
        # minwe.fill(20)
        # maxwe = np.empty(len(df1), dtype=np.int)
        # maxwe.fill(24)
        # dfwe = pl.DataFrame([minwe, maxwe], schema=[('Min', pl.Float32), ('Max', pl.Float32)])
        # data_update = data_update.hstack(dfwe)
        data_update_pd = data_update.to_pandas()
        append_df_to_excel(self.fileData, data_update_pd, sheet_name = 'Database', index = False, startrow = 0, header = True, truncate_sheet = True)
        data_print_update = pl.read_excel(source = self.fileUpdateData, sheet_name = 'Layouts',
                             read_csv_options = {'has_header': True, 'infer_schema_length': None}, xlsx2csv_options = {'skip_empty_lines': True, 'skip_hidden_rows': False})
        data_print_update_pd = data_print_update.to_pandas()
        append_df_to_excel(self.fileData, data_print_update_pd, sheet_name = 'Layouts', index = False, startrow = 0, header = True, truncate_sheet = True)
        self.data, self.data_print = self.loaddata()
        self.MainWin.in4(infor = 'Hoàn thành cập nhật', title = 'Thông tin')