import os, polars as pl, sys, logging, inspect
from pathlib import Path
from shutil import copyfile
from SubFunction import append_df_to_excel

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

class LoadData():
    def __init__(self, MainWin, fileData: Path, filePrint: Path, fileUpdateData: Path, fileUpdatePrint: Path):
        self.data = []
        self.data_update = None
        self.MainWin = MainWin
        self.fileData = fileData
        self.fileUpdateData = fileUpdateData
        self.filePrint = filePrint
        self.fileUpdatePrint = fileUpdatePrint
        try:
            self.data, self.data_print = self.loaddata()
        except Exception as e:
            logging.error(f"Error initializing LoadData at {get_file_and_line()}: {e}")

    def loaddata(self) -> pl.DataFrame:
        try:
            self.MainWin.resetdata()
            data = pl.read_excel(source = self.fileData, sheet_name = 'Database', schema_overrides = {'Code_Item': str, 'Quantity_thung': int},
                                 engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False},
                                 columns=['Code_Item', 'Name', 'ProductOf', 'Quantity_thung',
                                          'CAT', 'INT', 'Size', 'Dinh_dang_Tem_Thung',
                                          'Ma_1', 'Ma_vach_thung_dau', 'Ma_vach_thung_duoi', 'Ma_vach_nho', 'Don_vi_thung'])
            data_print = pl.read_excel(source = self.fileData, sheet_name = 'Layouts',
                                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False})
            return data, data_print
        except Exception as e:
            logging.error(f"Error in loaddata method at {get_file_and_line()}: {e}")
            raise

    def updatedata(self) -> pl.DataFrame:
        try:
            os.system('taskkill /f /IM EXCEL.exe')
            copyfile(self.fileUpdatePrint, self.filePrint)
            data_update = pl.read_excel(source = self.fileUpdateData, sheet_name = 'Sheet1', schema_overrides = {'Code Item': str, 'Quantity thùng': int, 'Quantity': int},
                                 engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False},
                                 columns=['Code Item', 'IND', 'Item', 'ProductOf',
                                 'Quantity thùng', 'Name', 'CAT', 'INT',
                                 'Size', 'Định dạng Tem Thùng', 'Quantity', 'Mã 1', 
                                 'Mã vạch thùng đầu', 'Mã vạch thùng đuôi', 'Mã vạch túi đầu', 'Mã vạch túi cuối',
                                 'Định dạng Tem Túi', 'Mã vạch nhỏ', 'Đơn vị thùng', 'Đơn vị túi'])
            data_update = data_update.filter(pl.col('Code Item').is_not_null() & (pl.col('Code Item') != ''))
            data_update = data_update.rename({
                        'Code Item': 'Code_Item',
                        'Quantity thùng': 'Quantity_thung',
                        'Định dạng Tem Thùng': 'Dinh_dang_Tem_Thung',
                        'Quantity': 'Quantity_tui',
                        'Mã 1': 'Ma_1',
                        'Mã vạch thùng đầu': 'Ma_vach_thung_dau',
                        'Mã vạch thùng đuôi': 'Ma_vach_thung_duoi',
                        'Mã vạch túi đầu': 'Ma_vach_tui_dau',
                        'Mã vạch túi cuối': 'Ma_vach_tui_cuoi',
                        'Định dạng Tem Túi': 'Dinh_dang_Tem_Tui',
                        'Mã vạch nhỏ': 'Ma_vach_nho',
                        'Đơn vị thùng': 'Don_vi_thung',
                        'Đơn vị túi': 'Don_vi_tui'
                    })
            data_update_pd = data_update.to_pandas()
            append_df_to_excel(self.fileData, data_update_pd, sheet_name = 'Database', index = False, startrow = 0, header = True, truncate_sheet = True)
            data_print_update = pl.read_excel(source = self.fileUpdateData, sheet_name = 'Layouts',
                                 engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False})
            data_print_update_pd = data_print_update.to_pandas()
            append_df_to_excel(self.fileData, data_print_update_pd, sheet_name = 'Layouts', index = False, startrow = 0, header = True, truncate_sheet = True)
            self.data, self.data_print = self.loaddata()
        except Exception as e:
            logging.error(f"Error in updatedata method at {get_file_and_line()}: {e}")
            raise