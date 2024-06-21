
'''
@FileName : Commands.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for Commands execution
'''

import logging
import subprocess
from Base.FrameworkConstants import NumericConstants

log = logging.getLogger(__name__)

class Commands(object):

    def __init__(self):
        '''
        @function:
            Initializes Commands Object

        @param: None
        @return: None
        '''
        self._shellValue = False
        self._inputhandler = subprocess.DEVNULL


    def setShellValue(self,value):
        '''
        @function:
            Sets Shell Value

        @param: Value :  Value to be set
        @return: None
        '''
        self._shellValue = value

    def getShellValue(self):
        '''
        @function:
            gets Shell Value

        @param: None
        @return: shellValue : Retrieves Shell Value
        '''
        return self._shellValue

    def setInputHandler(self,value):
        '''
        @function:
            Sets Input Handler Value

        @param: Value :  Value to be set
        @return: None
        '''
        self._inputhandler = value

    def getInputHandler(self):
        '''
        @function:
            Gets Input Handler Value

        @param: None
        @return: Retrieves Input Handler Value
        '''
        return self._inputhandler


    def executeCommand(self, executeCmd=None, timeout = NumericConstants.DEFAULT_TIMEOUT.value):
        '''
        @function:
            Executes the Command

        @param: executeCmd :  Command to be executed
        @param: timeout :  Timeout or Default Timeout Value
        @return: None
        '''
        try:

            output = subprocess.check_output(executeCmd,
                                             timeout=timeout,
                                             stderr=subprocess.DEVNULL,
                                             shell=self.getShellValue(),
                                             stdin=self.getInputHandler())
        except subprocess.CalledProcessError as cpe:
            output = cpe.output
        except Exception as exc:
            result = exc.__str__()
            log.error('Error executing command ' + result)
            output = None
            raise exc
        if output is not None:
            outputStr = output.decode('utf-8', 'backslashreplace').strip()
            log.debug('Command output - ' + outputStr)
            return outputStr
        else:
            return None
