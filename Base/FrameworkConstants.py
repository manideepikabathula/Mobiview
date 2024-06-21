'''

@FileName : FrameworkConstants.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for FrameworkConstants
'''


import os
import logging
from enum import Enum

log = logging.getLogger(__name__)

class NumericConstants(Enum):
    '''
        Class Containing Numeric Constants to be used accordingly

    :param: None
    :return: None
    '''
    DEFAULT_ITERATION = int(1)
    DEFAULT_TIMEOUT = int(240)

class HostFilePaths():

    '''
        Class Containing FilePaths to be used accordingly

    :param: None
    :return: None
    '''
    def __init__(self):
        '''
        @function:
            Initializes FilePaths Object

        @param: None
        @return: None
        '''
        self.rootdir = None
        self.basedir = None
        self.specsinfodir = None
        self.libdir = None
        self.logdir = None
        self.runnerdir = None
        self.testdir = None

    def setup(self):
        self.rootdir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.basedir = os.path.join(self.rootdir, "Base")
        self.specsinfodir = os.path.join(self.rootdir, "SpecsInfo")
        self.libdir = os.path.join(self.rootdir, "Lib")
        self.logdir = os.path.join(self.rootdir, "Logs")
        self.runnerdir = os.path.join(self.rootdir, "Runner")
        self.testdir = os.path.join(self.rootdir, "Test")

        if not os.path.exists(self.logdir):
            os.makedirs(self.logdir)
    def getRootDir(self):
        '''
        @function:
            Returns rootDir filePath

        @param: None
        @return: returns rootDir filePath
        '''
        self.setup()
        return self.rootdir

    def getBaseDir(self):
        '''
        @function:
            Returns basedir filePath

        @param: None
        @return: returns basedir filePath
        '''
        self.setup()
        return self.basedir

    def getDeviceInfoDir(self):
        '''
        @function:
            Returns deviceinfo filePath

        @param: None
        @return: returns deviceinfo filePath
        '''
        self.setup()
        return self.specsinfodir

    def getLogDir(self):
        '''
        @function:
            Returns logdir filePath

        @param: None
        @return: returns logdir filePath
        '''
        self.setup()
        return self.logdir

    def getRunnerDir(self):
        '''
        @function:
            Returns runnerdir filePath

        @param: None
        @return: returns runnerdir filePath
        '''
        self.setup()
        return self.runnerdir

    def getTestDir(self):
        '''
        @function:
            Returns runnerdir filePath

        @param: None
        @return: returns runnerdir filePath
        '''
        self.setup()
        return self.testdir

class TestResult(Enum):
    '''
        Class Test Results Options

    :param: None
    :return: None
    '''

    PASS = 'Pass'
    FAIL = 'Fail'
    CRASHED = 'Crashed'
    TIMEDOUT = 'Timedout'
    NOTTESTED = 'NotTested'
    INFRAERROR = 'InFraError'
    COMPLETED = 'Completed'
    UNSUPPORTED = 'Unsupported'
    NOTAPPLICABLE = 'NotApplicable'
    ANALYSISPENDING = 'AnalysisPending'

    def __str__(self):
        return str(self.value)

class DeviceFilePaths():
    '''
        Class Containing FilePaths to be used accordingly

    :param: None
    :return: None
    '''
    def __init__(self):
        '''
        @function:
            Initializes FilePaths Object

        @param: None
        @return: None
        '''
        self.rootDir = None
        self.sysDir = None
        self.systemDir = None
        self.mntDir = None
        self.vendorDir = None
        self.dataDir = None
        self.procDir = None
        self.etcDir = None
        self.debugDir = None
        self.sdcardDir = None
        self.storageDir = None
        self.devDir = None
        self.odmDir = None

    def setup(self):
        self.rootDir = '/'
        self.sysDir = os.path.join(self.rootDir, 'sys')
        self.systemDir = os.path.join(self.rootDir, 'system')
        self.mntDir = os.path.join(self.rootDir, 'mnt')
        self.vendorDir = os.path.join(self.rootDir, 'vendor')
        self.dataDir = os.path.join(self.rootDir, 'data')
        self.procDir = os.path.join(self.rootDir, 'proc')
        self.etcDir = os.path.join(self.rootDir, 'etc')
        self.debugDir = os.path.join(self.rootDir, 'd')
        self.sdcardDir = os.path.join(self.rootDir, 'sdcard')
        self.storageDir = os.path.join(self.rootDir, 'storage')
        self.devDir = os.path.join(self.rootDir, 'dev')
        self.odmDir = os.path.join(self.rootDir, 'odm')

    def getRootDir(self):
        return  self.rootDir

    def getSysDir(self):
        return  self.sysDir

    def getSystemDir(self):
        return self.systemDir

    def getMntDir(self):
        return self.mntDir

    def getVendorDir(self):
        return  self.vendorDir

    def getDataDir(self):
        return self.dataDir

    def getEtcDir(self):
        return  self.etcDir

    def getDebugDir(self):
        return self.debugDir

    def getSDCardDir(self):
        return self.sdcardDir

    def getStorageDir(self):
        return self.storageDir

    def getDevDir(self):
        return self.devDir

    def getODMDir(self):
        return self.odmDir

    def setup(self):
        self.rootdir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.basedir = os.path.join(self.rootdir, "Base")
        self.specsinfodir = os.path.join(self.rootdir, "SpecsInfo")
        self.libdir = os.path.join(self.rootdir, "Lib")
        self.logdir = os.path.join(self.rootdir, "Logs")
        self.runnerdir = os.path.join(self.rootdir, "Runner")
        self.testdir = os.path.join(self.rootdir, "Test")

        if not os.path.exists(self.logdir):
            os.makedirs(self.logdir)

