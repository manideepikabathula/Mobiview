
''''
@FileName : Runner.py
@Author : Srinivas Ganti
@place : Hyderabad, 07 Jan 2024

@purpose : Class Contain Definition of functions
           for Runner execution
'''


#from Runner.Runner import BaseRunner
from Runner import BaseRunner

class  Current_Runner(BaseRunner):

    def __init__(self):
        self.runner = None

    def run(self):
        self.runner= BaseRunner()
        #super().run()
        self.runner.run()


def main():
    crObj = Current_Runner()
    crObj.run()

if __name__ == '__main__':
    main()
