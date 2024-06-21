'''

@FileName : MobileSpecsInfo.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for retrieving Mobile Specifications from Mobile
'''

import logging
from Base.ADB import ADBInfo
from Base.Commands import Commands
from Base.FrameworkConstants import HostFilePaths
from Base.FrameworkConstants import NumericConstants

log = logging.getLogger(__name__)


class MobileSpecsInfo(object):
    '''
    Abstract Class for Mobile Specs Info  Object

    :param: None
    :return: None
    '''

    def __init__(self):
        '''
        @function:
            Initializes Mobile Specs Info  Object

        @param: None
        @return: None
        '''

        self.cmdObj = None
        self.ADBObj = None
        self.fpObj = None

        self.output = None
        self.workbook = None
        self.worksheet = None
        self.command = None

        self.HwSpecsInfoDict = {}
        self.SwSpecsInfoDict  = {}
        self.CpuSpecsInfoDict = {}
        self.MemSpecsInfoDict = {}
        self.DisplaySpecsInfoDict ={}
        #self.batterySpecsInfoDict = {}
        self.cameraSpecsInfoDict = {}

        self.SwUpdatesSpecsInfoDict = {}
        self.MobileSpecsInfoDict = {} # Cumulative Dictionary
        self.setup()

    def setup(self):

        self.cmdObj = Commands()
        self.ADBObj = ADBInfo()
        self.fpObj = HostFilePaths()

        self.HwSpecsInfoDict = {}
        self.SwSpecsInfoDict = {}
        self.CpuSpecsInfoDict = {}
        self.MemSpecsInfoDict = {}
        self.DisplaySpecsInfoDict = {}
        self.cameraSpecsInfoDict = {}
        # self.BuildSpecsInfo = {}
        self.SwUpdatesSpecsInfoDict = {}
        self.MobileSpecsInfoDict = {}  # Cumulative Dictionary

    def grepInfo(self):
        pass

    def cleanup(self):
        pass

    def generateXLSXReport(self, xlsxObj=None, wb=None,ws=None, dataDict=None):
        pass

    def updateDictionary(self, dictName=None, key=None, value=None):
        dictName[key] = value

    def printDictionary(self, dictName=None):
        if dictName:
            log.info("Printing Items of : {} Dictionary".format(dictName))
            for key, value in dictName.items():
                log.info("{} : {}".format(key,value))

    def mergeDictionaries(self, srcDict=None, destDict=None):
        pass

    def printLogInfo(self, cmdInfo=None, message=None):
        log.info("{:>30}: {:>30}".format("Command Executed ", cmdInfo))
        if message is not None:
            log.info("{:>30}: {:>30}".format("Command Status ", "Success"))
        else:
            log.info("{:>30}: {:>30}".format("Command Status ", "Failed"))
        log.info("{:>30}: {:>30}".format("Command Output ", message))

    def executeCommandOnDevice(self, command=None, timeout=NumericConstants.DEFAULT_TIMEOUT.value):
        '''
        @function: executeCommandOnDevice
            executes Command On Device and Obtains output

        @param: Command to be executed
        @return: returns output of executed command
        '''
        self.output = self.cmdObj.executeCommand(executeCmd=command, timeout=timeout)
        self.printLogInfo(cmdInfo=command, message=self.output)
        return self.output




