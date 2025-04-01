import os, polars as pl, sys, logging, inspect, json
from pathlib import Path
from shutil import copyfile
from SubFunction import append_df_to_excel
from PyQt5 import QtCore

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
    # Add signal for progress updates
    progress_update = QtCore.pyqtSignal(int, str)

    def __init__(self, MainWin, fileData: Path, filePrint: Path, fileUpdateData: Path, fileUpdatePrint: Path, fileHeadersConfig: Path):
        self.data = []
        self.data_update = None
        self.MainWin = MainWin
        self.fileData = fileData
        self.fileUpdateData = fileUpdateData
        self.filePrint = filePrint
        self.fileUpdatePrint = fileUpdatePrint
        
        # Read headers configuration from JSON file
        try:
            with open(fileHeadersConfig, 'r', encoding='utf-8') as f:
                headers_config = json.load(f)
                self.headers_mapping = headers_config['headers']
                self.headersToRead = list(self.headers_mapping.values())
                self.headersToUpdate = list(self.headers_mapping.keys())
        except Exception as e:
            logging.error(f"Error loading headers config at {get_file_and_line()}: {e}")
            self.headers_mapping = {}
            self.headersToRead = []
            self.headersToUpdate = []
            
        try:
            self.data, self.data_print = self.load_data()
        except Exception as e:
            logging.error(f"Error initializing LoadData at {get_file_and_line()}: {e}")

    def load_data(self) -> pl.DataFrame:
        try:
            self.MainWin.resetdata()
            data = pl.read_excel(source=self.fileData, sheet_name='Database', 
                                 schema_overrides={'Code_Item': str, 'Quantity_thung': int},
                                 engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False},
                                 columns=self.headersToRead)
            data_print = pl.read_excel(source=self.fileData, sheet_name='Layouts',
                                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False})
            return data, data_print
        except Exception as e:
            logging.error(f"Error in load_data method at {get_file_and_line()}: {e}")
            self.update_data()
            return

    def update_data(self) -> pl.DataFrame:
        try:
            # Update to show "Closing Excel..."
            self.safe_update_progress(20, "Đang đóng Excel...")
            os.system('taskkill /f /IM EXCEL.exe')
            
            # Update to show "Copying template file..."
            self.safe_update_progress(30, "Đang sao chép tệp mẫu...")
            copyfile(self.fileUpdatePrint, self.filePrint)
            
            # Update to show "Reading data..."
            self.safe_update_progress(40, "Đang đọc dữ liệu...")
            
            # Only use headers defined in headers.json file
            data_update = pl.read_excel(source=self.fileUpdateData, sheet_name='Sheet1', 
                                       schema_overrides={'Code Item': str, 'Quantity thùng': int},
                                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False},
                                       columns=self.headersToUpdate)
            
            self.safe_update_progress(50)
            
            data_update = data_update.filter(pl.col('Code Item').is_not_null() & (pl.col('Code Item') != ''))
            
            # Rename columns using only the headers from headers.json
            data_update = data_update.rename(self.headers_mapping)
            
            # Update Database sheet
            self.safe_update_progress(60, "Đang cập nhật cơ sở dữ liệu...")
            data_update_pd = data_update.to_pandas()
            append_df_to_excel(self.fileData, data_update_pd, sheet_name='Database', index=False, startrow=0, header=True, truncate_sheet=True)
            
            # Update Layouts sheet
            self.safe_update_progress(80, "Đang cập nhật bố cục...")
            data_print_update = pl.read_excel(source=self.fileUpdateData, sheet_name='Layouts',
                                 engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False})
            data_print_update_pd = data_print_update.to_pandas()
            append_df_to_excel(self.fileData, data_print_update_pd, sheet_name='Layouts', index=False, startrow=0, header=True, truncate_sheet=True)

            # Update data
            self.safe_update_progress(90, "Đang tải dữ liệu mới...")
            self.MainWin.resetdata()
            try:
                data = pl.read_excel(source=self.fileData, sheet_name='Database', 
                                    schema_overrides={'Code_Item': str, 'Quantity_thung': int},
                                    engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False},
                                    columns=self.headersToRead)
                data_print = pl.read_excel(source=self.fileData, sheet_name='Layouts',
                                        engine='xlsx2csv', engine_options={"skip_empty_lines": True, 'skip_hidden_rows': False})
                self.data = data
                self.data_print = data_print
            except Exception as e:
                logging.error(f"Error reloading data in update_data method at {get_file_and_line()}: {e}")
                
        except Exception as e:
            logging.error(f"Error in update_data method at {get_file_and_line()}: {e}")
            raise
            
    # Add a safe helper method to update progress
    def safe_update_progress(self, value, text=None):
        try:
            # Try the Qt queued connection method first
            if text is not None:
                QtCore.QMetaObject.invokeMethod(
                    self.MainWin, "update_progress", 
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(int, value), 
                    QtCore.Q_ARG(str, text)
                )
            else:
                QtCore.QMetaObject.invokeMethod(
                    self.MainWin, "update_progress", 
                    QtCore.Qt.QueuedConnection,
                    QtCore.Q_ARG(int, value), 
                    QtCore.Q_ARG(str, "")
                )
        except Exception:
            # Fall back to direct method call
            try:
                self.MainWin.update_progress(value, text)
            except Exception as e:
                logging.error(f"Failed to update progress: {e}")