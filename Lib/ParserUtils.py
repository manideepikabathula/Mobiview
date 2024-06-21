'''

@FileName : ParserUtils.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for Parser Utilities
'''

import re
import logging

log = logging.getLogger(__name__)
class ParserUtils():

    @staticmethod
    def parseDataViaRegex(regexpattern=None,text=None):
        match = re.search(pattern=regexpattern, string=text)

        return match.groupdict()







