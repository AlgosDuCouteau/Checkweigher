import os, sys, logging, inspect

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

class GetConfig(object):
    def __init__(self):
        try:
            self.config = self.data()
        except Exception as e:
            logging.error(f"Error initializing GetConfig at {get_file_and_line()}: {e}")
    
    def data(self):
        try:
            file = self.getfilepath()
            port = self.readfileport(file['filePort'])
            file_final = file.copy()
            file_final.pop('filePort')
            file_final.update(port)
            file_final.update({'GdriveData': os.path.join(port['GdriveData'], 'My Drive/IN&PACKING.xlsm')})
            file_final.update({'GdrivePrint': os.path.join(port['GdrivePrint'], 'My Drive/in.xlsx')})
            return file_final
        except Exception as e:
            logging.error(f"Error in data method at {get_file_and_line()}: {e}")
            raise

    def getfilepath(self):
        try:
            basedir = resource_path()
            categorization_fileData = os.path.join(basedir + '/data/', 'Database.xlsx')
            categorization_filePort = os.path.join(basedir + '/data/', 'ports.txt')
            categorization_fileIN = os.path.join(basedir + '/data/', 'in.xlsx')
            return {'fileData': categorization_fileData, 'filePrint': categorization_fileIN, 'filePort': categorization_filePort}
        except Exception as e:
            logging.error(f"Error in getfilepath method at {get_file_and_line()}: {e}")
            raise
    
    def readfileport(self, filePort):
        try:
            portall = []
            with open(filePort) as f:
                contents = f.readlines()
                for i in range(len(contents)):
                    contents[i] = contents[i].strip().replace(" ", "")
                    portall.append(contents[i].rsplit(':', 1)[1])
            return {'portScale': portall[0], 'portArduino': portall[1], 'Ard2Convey': int(portall[2]), 'Ard2Light': int(portall[3]),
                    'GdriveData': portall[4]+ ":/", 'GdrivePrint': portall[4]+ ":/", 'Delay2Print': float(portall[5]), 'Quan2Print': int(portall[6])}
        except Exception as e:
            logging.error(f"Error in readfileport method at {get_file_and_line()}: {e}")
            raise
    
if __name__ == "__main__":
    try:
        data = GetConfig()
    except Exception as e:
        logging.error(f"Error in main execution of GetConfig at {get_file_and_line()}: {e}")