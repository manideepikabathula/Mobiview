'''

@FileName : GrepInfo.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for Grepping infor and Generate Individual Specs Json File
'''


import os
import logging
import cProfile
from Base import LogConf
from datetime import datetime
from Base.FrameworkConstants import HostFilePaths
from Lib.FileSystemUtils import FileSystemUtils
from SpecsInfo.HwSpecsInfo import HWSpecsInfo
from SpecsInfo.SwSpecsInfo import SwSpecsInfo
from SpecsInfo.CpuSpecsInfo import CpuSpecsInfo
from SpecsInfo.DisplaySpecsInfo import DisplaySpecsInfo
from Lib.XlsxGenerator import XlsxGenerator

log = logging.getLogger(__name__)

class GrepInfo():

    '''
    Class for grepInfo Object Object

    :param: None
    :return: None
    '''

    def __init__(self):
        '''
        @function:
            Initializes GrepInfo Object

        @param: None
        @return: None
        '''

        self.HwSpecsObj = None
        self.SwSpecsObj = None
        self.CpuSpecsObj = None
        #self.BatterySpecsObj = None
        self.DisplaySpecsObj = None

        self.fpObj = None
        self.xlObj = None

        self.wb = None
        self.activesheet = None
        self.resultsDir = None
        self.jsonFileName = None
        self.jsonFilePath = None
        self.newLogFileName = None

        self.specsDict = {}

    def setup(self):
        self.fpObj = HostFilePaths()
        self.xlObj = XlsxGenerator()

        self.HwSpecsObj = HWSpecsInfo()
        self.SwSpecsObj = SwSpecsInfo()
        self.CpuSpecsObj = CpuSpecsInfo()
#        self.BatterySpecsObj = BatterySpecsInfo()
        self.DisplaySpecsObj = DisplaySpecsInfo()

        self.HwSpecsObj.setup()
        self.SwSpecsObj.setup()
        self.CpuSpecsObj.setup()
       # self.BatterySpecsObj.setup()
        self.DisplaySpecsObj.setup()

        self.setResultsDir()
        LogConf.configureLogging(self.getResultsFolder())
        workbookname = self.HwSpecsObj.getDeviceSerialNo()+'_Device_Specs_Info.xlsx'
        #workbookname = self.CpuSpecsObj.getClusterCount()+'_Device_Specs_Info.xlsx'

        self.xlObj.createOrLoadWorkBook(folder=self.getResultsFolder(),
                                        bookname= workbookname)
        self.wb = self.xlObj.getWorkBook()

        self.xlObj.AddLastRowStyle()
        self.xlObj.AddHeaderRowStyle()
        self.xlObj.AddNormalRowStyle()


    def execute(self):

        self.generateSpecInfoJson(self.HwSpecsObj)
        self.generateSpecInfoJson(self.SwSpecsObj)
        self.generateSpecInfoJson(self.CpuSpecsObj)
        #self.generateSpecInfoJson(self.BatterySpecsObj)
        self.generateSpecInfoJson(self.DisplaySpecsObj)



    def generateSpecInfoJson(self,specsInfoObj=None):
        self.newLogFileName = 'console_output_'+specsInfoObj.__class__.__name__+'.log'
        self.jsonFileName = specsInfoObj.__class__.__name__ + '.json'
        self.jsonFilePath = os.path.join(self.getResultsFolder(), self.jsonFileName)
        self.wb.create_sheet(specsInfoObj.__class__.__name__)
        self.ws = self.xlObj.getWorkSheet(self.wb, sheet_name=specsInfoObj.__class__.__name__)
        self.specsDict = specsInfoObj.grepInfo()
        FileSystemUtils.createJsonFromDict(self.specsDict, self.jsonFilePath)

        FileSystemUtils.renameLogFile(os.path.join(self.getResultsFolder(), 'console_output.log'),
                                      os.path.join(self.getResultsFolder(), self.newLogFileName))

        specsInfoObj.generateXLSXReport(self.xlObj,self.wb,self.ws,self.specsDict)

    '''def generateSpecInfoJson(self, specsInfoObj=None):
        self.newLogFileName = 'console_output_' + specsInfoObj.__class__.__name__ + '.log'
        self.jsonFileName = specsInfoObj.__class__.__name__ + '.json'
        self.jsonFilePath = os.path.join(self.getResultsFolder(), self.jsonFileName)

        self.specsDict = specsInfoObj.grepInfo()
        print("specsDict:", self.specsDict)  # Add this line to print the dictionary

        df = pd.DataFrame.from_dict(self.specsDict, orient='index').transpose()

        df.to_json(self.jsonFilePath, orient='records')

        df.to_excel(os.path.join(self.getResultsFolder(), specsInfoObj.__class__.__name__ + '.xlsx'), index=False)

        FileSystemUtils.renameLogFile(os.path.join(self.getResultsFolder(), 'console_output.log'),
                                      os.path.join(self.getResultsFolder(), self.newLogFileName))'''

    def cleanup(self):
        del self.wb['Sheet']
        self.wb.save(self.xlObj.getXlsxFilePath())

    def setResultsDir(self):
        logFolderName = "{}_{}".format(self.__class__.__name__, datetime.now().strftime('%Y_%m_%d_%H_%M_%S'))

        self.resultsDir = os.path.join(self.fpObj.getLogDir(),
                                       logFolderName,
                                       'Results',
                                       self.HwSpecsObj.getDeviceSerialNo())
        if not os.path.exists(self.resultsDir):
            os.makedirs(self.resultsDir)

    '''def setResultsDir(self):
        logFolderName = "{}_{}".format(self.__class__.__name__, datetime.now().strftime('%Y_%m_%d_%H_%M_%S'))

        self.resultsDir = os.path.join(self.fpObj.getLogDir(),
                                       logFolderName,
                                       'Results',
                                       self.CpuSpecsObj.getClusterCount())
        if not os.path.exists(self.resultsDir):
            os.makedirs(self.resultsDir)'''


    def getResultsFolder(self):
        return self.resultsDir



