'''

@FileName : FileSystemUtils.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for FileSystem Utilities
'''

import os
import re
import sys
import glob
import argparse
import logging
import traceback
from pydoc import locate

from Base.FrameworkConstants import HostFilePaths

fpObj = HostFilePaths()

class BaseRunner():
    def run(self):

        try:
            sys.path.append(fpObj.getRootDir())
            sys.path.append(fpObj.getBaseDir())
            sys.path.append(fpObj.getDeviceInfoDir())
            sys.path.append(fpObj.getRunnerDir())
            sys.path.append(fpObj.getLogDir())
            sys.path.append(fpObj.getTestDir())

            parser = argparse.ArgumentParser()
            parser.add_argument('--name', '-n', dest='classname', default='',
                                help='Name of the class')

            cmdLineArgs, unknownArgs = parser.parse_known_args()

            fullQualifiedClassName = self.getFullClassName(cmdLineArgs.classname)
            print("Fully Qualified Class Name : {} ".format(fullQualifiedClassName))
            testClass = locate(fullQualifiedClassName)

            test=testClass()
        except Exception as e:
            traceback.print_exc()
            logging.error("Unexpected error while test class initialization:" + e.__str__())
            sys.exit()

        try:
            test.setup()
            test.execute()
        except Exception as e:
            traceback.print_exc()
            logging.error("Unexpected error in setup() or execute() :" + e.__str__())
        finally:
            try:
                test.cleanup()
            except Exception as e:
                traceback.print_exc()
                logging.error("Unexpected error while cleanup:" + e.__str__())

    def getFullClassName(self,clsname=None):
        fullyQualifiedClassName = None

        #rootDir = FilePaths.ROOT_FOLDER
        filesList = glob.glob(os.path.join(fpObj.getRootDir(), '**', '*.py'), recursive=True)
        for filePath in filesList:
            with open(filePath, 'r', encoding='utf8', errors='ignore') as fileHandler:
                for line in fileHandler:
                    pattern = re.search(
                        re.compile(r'''.*?class\s+({})\s*\('''.format(clsname.rstrip()), re.IGNORECASE),
                        line.rstrip())
                    if pattern:
                        fullyQualifiedClassName = "{}.{}.{}".format(os.path.basename(os.path.dirname(filePath)),
                                                                    re.sub('\\.py$', '', os.path.basename(filePath)),
                                                                    pattern.group(1))
                        break
        return fullyQualifiedClassName

