'''

@FileName : ADB.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for ADB
'''

from enum import Enum
import os
import re
import json
import logging
from Base.Commands import Commands


log = logging.getLogger(__name__)

class ADBInfo():
    '''
        Class for ADB Object

    :param: None
    :return: None
    '''

    def __init__(self):
        '''
        @function:
            Initializes Device Object

        @param: None
        @return: None
        '''
        self.baseADBCmd = None
        self.command = None

    def getADBCommand(self, waitfordevice=True):
        '''
        @function:
            gets Basic ADB Command

        @param: waitfordevice : True/False
        @return: returns adb basic command 'adb wait-for-device'
        '''
        self.baseADBCmd = 'adb wait-for-device '
        return self.baseADBCmd

    def getADBShellCommand(self):
        '''
        @function:
            gets  ADB Shell Command

        @param: None
        @return: returns adb shell command 'adb wait-for-device shell'
        '''
        self.command = self.getADBCommand() + ' shell '
        return self.command

    def getADBPullCommand(self):
        '''
        @function:
            gets  ADB Pull Command

        @param: None
        @return: returns adb pull command 'adb wait-for-device pull'
        '''
        self.command = self.getADBCommand() + ' pull '
        return self.command

    def getADBPushCommand(self):
        '''
        @function:
            gets  ADB Push Command

        @param: None
        @return: returns adb pull command 'adb wait-for-device push'
        '''
        self.command = self.getADBCommand() + ' push '
        return self.command

    def getADBRebootCommand(self):
        '''
        @function:
            gets  ADB reboot Command

        @param: None
        @return: returns adb reboot command 'adb wait-for-device reboot'
        '''
        self.command = self.getADBCommand() + ' reboot '
        return self.command

    def getADBRebootBootLoaderCommand(self):
        '''
        @function:
            gets  ADB reboot bootloader Command

        @param: None
        @return: returns adb pull command 'adb wait-for-device reboot bootloader'
        '''
        self.command = self.getADBRebootCommand() + ' bootloader '
        return self.command

    def getADBRebootRecoveryCommand(self):
        '''
        @function:
            gets  ADB reboot recovery Command

        @param: None
        @return: returns adb pull command 'adb wait-for-device reboot recovery'
        '''
        self.command = self.getADBRebootCommand() + ' recovery '
        return self.command

    def getADBInstallCommand(self, option = ' r '):
        '''
        @function:
            gets  ADB install Command

        @param: None
        @return: returns adb install command 'adb wait-for-device install'
        '''
        self.command = self.getADBCommand() + ' install ' + option
        return self.command

    def getADBUninstallCommand(self, option = ' -k '):
        '''
        @function:
            gets  ADB uninstall Command

        @param: None
        @return: returns adb uninstall command 'adb wait-for-device uninstall'
        '''
        self.command = self.getADBCommand() + ' uninstall ' + option
        return self.command

    def getADBVersionCommand(self):
        '''
        @function:
            gets  ADB version Command

        @param: None
        @return: returns adb version command 'adb wait-for-device version'
        '''
        self.command = self.getADBCommand() + ' version '
        return self.command

    def getADBStateCommand(self):
        '''
        @function:
            gets  ADB state Command

        @param: None
        @return: returns adb state command 'adb wait-for-device get-state'
        '''
        self.command = self.getADBShellCommand() + ' get-state '
        return self.command

    def getADBSerialNoCommand(self):
        '''
        @function:
            gets  ADB SerialNo Command

        @param: None
        @return: returns adb SerialNo command 'adb wait-for-device get-serialno'
        '''
        self.command = self.getADBCommand() + ' get-serialno '
        return self.command

    def getADBStartServerCommand(self):
        '''
        @function:
            gets  ADB start-server Command

        @param: None
        @return: returns adb start-server command 'adb wait-for-device start-server'
        '''
        self.command = self.getADBCommand() + ' start-server '
        return self.command

    def getADBKillServerCommand(self):
        '''
        @function:
            gets  ADB Kill-Server Command

        @param: None
        @return: returns adb kill-server command 'adb wait-for-device kill-server'
        '''
        self.command = self.getADBCommand() + ' kill-server '
        return self.command

    def getADBLogcatCommand(self):
        '''
        @function:
            gets  ADB logcat Command

        @param: None
        @return: returns adb logcat command 'adb wait-for-device logcat'
        '''
        self.command = self.getADBCommand() + ' logcat '
        return self.command

    def getADBGetPropCommand(self):
        '''
        @function:
            gets  ADB logcat Command

        @param: None
        @return: returns adb logcat command 'adb wait-for-device logcat'
        '''
        self.command = self.getADBShellCommand() + ' getprop '
        return self.command

    def getADBDumpsysCommand(self):
        '''
        @function:
            gets  ADB DumpSys Command

        @param: None
        @return: returns adb DumpSys command 'adb wait-for-device logcat'
        '''
        self.command = self.getADBShellCommand() + ' dumpsys | grep -i '
        return self.command

    def getADBWindowsManagerCommand(self):
        '''
        @function:
            gets  ADB WM Command

        @param: None
        @return: returns adb DumpSys command 'adb wait-for-device logcat'
        '''
        self.command = self.getADBShellCommand() + ' wm '
        return self.command

    def getADBListCommand(self):
        '''
        @function:
            gets  ADB List Command

        @param: None
        @return: returns adb List command 'adb wait-for-device ls'
        '''
        self.command = self.getADBShellCommand() + ' ls '
        return self.command

    def getADBReadContentCommand(self):
        '''
        @function:
            gets  ADB cat Command

        @param: None
        @return: returns adb cat command 'adb wait-for-device cat'
        '''
        self.command = self.getADBShellCommand() + ' cat '
        return self.command












