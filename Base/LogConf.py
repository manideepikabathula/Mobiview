
'''
@FileName : LogConf.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for Logging Configuration
'''


import os
import stat
import logging


LOG_FILE = 'console_output.log'
LOG_FORMAT =  logging.Formatter(
    '%(asctime)s %(levelname)-8s %(name)-12s  [%(filename)s:%(lineno)d] [%(threadName)s]  %(message)s')



class LogHandler(object):

    def __init__(self, logFolder, handler) -> None:
        '''
        :function:
            Initializes LogHandler Object

        :param: logFolder :  Log Folder  configure Logging
        :param: handler :  Log Handler Type
        :return: None
        '''
        super().__init__()
        self.consoleLogName= None
        os.chmod(self.getConsoleLog(logFolder), stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)
        handler.setLevel(logging.DEBUG)
        handler.setFormatter(LOG_FORMAT)
        self._handler = handler


    def getLogHandle(self):
        '''
        :function:
            Retrieves Log Handler

        :param: None
        :return: Handler: Retrieves Handler
        '''
        return self._handler

    def withLogLevel(self, logLevel=None):
        '''
        :function:
            Set Log Level

        :param: logLevel : Sets LogLevel
        :return: Retrieves Class Object
        '''
        self._handler.setLevel(logLevel)
        return self

    def withFormatter(self, logFormatter=None):
        '''
        :function:
            Set Log Format

        :param: logFormatter : Sets Log Format
        :return: Retrieves Class Object
        '''
        self._handler.setFormatter(logFormatter)
        return self

    def getConsoleLog(self, logLocation=None):
        '''
        :function:
            Set Console Log

        :param: logLocation : Sets Console Log per LogLocation
        :return: Retrieves LogLocation
        '''

        return os.path.join(logLocation, LOG_FILE)

class FileLoggingHandler(LogHandler):
        def __init__(self, logFolder) -> None:
            super().__init__(logFolder, logging.FileHandler(self.getConsoleLog(logFolder), 'a'))


def configureLogging(logLocation, handlers=None):
    logger = logging.getLogger()
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG)

    configureRootConsoleLogger(logger)

    actualHandlers = handlers if handlers else [FileLoggingHandler(logLocation)]

    for logHandler in actualHandlers:
        logger.addHandler(logHandler.getLogHandle())


def configureRootConsoleLogger(logger=None):
    if logger is None:
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)

    handler = logging.StreamHandler()
    handler.setLevel(logging.INFO)
    handler.setFormatter(LOG_FORMAT)
    logger.addHandler(handler)
