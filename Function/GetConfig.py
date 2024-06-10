import os, sys

class GetConfig(object):
    def __init__(self):
        self.config = self.data()
    
    def data(self):
        file = self.getfilepath()
        port = self.readfileport(file['filePort'])
        file_final = file.copy()
        file_final.pop('filePort')
        file_final.update(port)
        file_final.update({'GdriveData': os.path.join(port['GdriveData'], 'My Drive/IN&PACKING.xlsm')})
        file_final.update({'GdrivePrint': os.path.join(port['GdrivePrint'], 'My Drive/in.xlsx')})
        return file_final

    def getfilepath(self):
        def resource_path():
            """ Get absolute path to resource, works for dev and for PyInstaller """
            if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  
            return base_path

        basedir = resource_path()
        categorization_fileData = os.path.join(basedir + '/data/', 'Database.xlsx')
        categorization_filePort = os.path.join(basedir + '/data/', 'ports.txt')
        categorization_fileIN = os.path.join(basedir + '/data/', 'in.xlsx')
        return {'fileData': categorization_fileData, 'filePrint': categorization_fileIN, 'filePort': categorization_filePort}
    
    def readfileport(self, filePort):
        portall = []
        with open(filePort) as f:
            contents = f.readlines()
            for i in range(len(contents)):
                contents[i] = contents[i].strip().replace(" ", "")
                portall.append(contents[i].rsplit(':', 1)[1])
        return {'portScale': portall[0], 'portArduino': portall[1], 'GdriveData': portall[2]+ ":/", 'GdrivePrint': portall[2]+ ":/", 'Ard2Convey': int(portall[3]),
                'Delay2Print': float(portall[6]), 'Quan2Print': int(portall[8])}
    
if __name__ == "__main__":
    data = GetConfig()