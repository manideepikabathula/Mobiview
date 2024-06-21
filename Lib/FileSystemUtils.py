'''

@FileName : FileSystemUtils.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for FileSystem Utilities
'''
import os
import json
import glob
import shutil
import logging

log = logging.getLogger(__name__)

class FileSystemUtils():



    @staticmethod
    def createJsonFromDict(dict=None, jsonFilePath=None):
        json_string = json.dumps(dict,ensure_ascii=True, indent=4)
        with open(jsonFilePath, 'w') as file:
            file.write(json_string)

    @staticmethod
    def getDictFromJson(jsonFilePath=None):
        with open(jsonFilePath, 'r') as file:
            dataDict = json.load(file)
        return dataDict


    @staticmethod
    def searchForFiles(fileExt=None, folderPath=None):
        if fileExt:
            wildcarddata = '*'+fileExt
        else:
            wildcarddata = '*.*'
        filesList = glob.glob(os.path.join(folderPath, '**', wildcarddata), recursive=True)
        return filesList

    @staticmethod
    def renameLogFile(srcPath=None,dstPath=None):
        shutil.copyfile(src=srcPath,dst=dstPath)

    @staticmethod
    def replaceChars(data,charlist):
        for char in charlist:
            data = str(data).replace(char, '')
        return data






