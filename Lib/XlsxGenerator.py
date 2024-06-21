'''

@FileName : XLSXGenerator.py
@Author : Srinivas Ganti
@place : Hyderabad, 22 Feb 2024

@purpose : Class Contain Definition of functions
           for Xlsx  Utilities
'''
import os
import logging
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment,NamedStyle

log = logging.getLogger(__name__)

class XlsxGenerator():

    def __init__(self):
        self.headers = None
        self.log_folder = None
        self.xlsx_File_Path = None
        self.work_Book = None
        self.work_Sheet = None
        self.cell_Style = None
        self.work_Book_Name = None
        self.cell_border = None
        self.cell_font = None
        self.cell_alignment =  None
        self.cell_pattern_fill = None
        self.stylesList = []

    def createOrLoadWorkBook(self, folder = None, bookname = None):
        if os.path.exists(folder):
            if bookname and bookname.endswith('.xlsx'):
                self.xlsx_File_Path = os.path.join(folder,bookname)
                if not os.path.exists(self.xlsx_File_Path):
                    self.work_Book = Workbook()
                else:
                    self.work_Book = load_workbook(self.xlsx_File_Path)
            else:
                self.xlsx_File_Path = os.path.join(folder, "Device_SpecsInfo.xlsx")
                self.work_Book = Workbook()

    def getWorkBook(self):
        return self.work_Book

    def getWorkSheet(self, wb=None, sheet_name = None):
        if wb is None:
            wb = self.getWorkBook()
        ws_sheet_names = wb.sheetnames
        for item in ws_sheet_names:
            if sheet_name == item:
                self.work_Sheet = wb[sheet_name]
            else:
                self.work_Sheet= wb.active

        return self.work_Sheet

    def AddHeaderRowStyle(self):
        self.setCellFont(bold=True)
        self.setPatternFill(fgcolor="00C0C0C0")
        headerStyle =  NamedStyle(name="headerRow",
                            font=self.getCellFont(),
                            border=self.getCellWithThickBorder(),
                            alignment=self.getCellAlignment(),
                            fill= self.getPatternFill())
        self.work_Book.add_named_style(headerStyle)

    def AddNormalRowStyle(self):
        self.setCellFont(bold=False)
        self.setPatternFill(fgcolor="00FFFFFF")
        normalStyle = NamedStyle(name="normalRow",
                                 font=self.getCellFont(),
                                 border=self.getCellWithThickThinBorder(),
                                 alignment=self.getCellAlignment(),
                                 fill=self.getPatternFill())
        self.work_Book.add_named_style(normalStyle)

    def AddLastRowStyle(self):
        self.setCellFont(bold=False)
        self.setPatternFill(fgcolor="00FFFFFF")
        normalStyle = NamedStyle(name="lastRow",
                                 font=self.getCellFont(),
                                 border=self.getCellWithThickBottomBorder(),
                                 alignment=self.getCellAlignment(),
                                 fill=self.getPatternFill())
        self.work_Book.add_named_style(normalStyle)

    def getNamedStyle(self, stylename = None):
        try:
            if stylename in  self.work_Book.named_styles:
                return stylename
        except ValueError:
            log.info("{} is not in the list.".format(stylename))
            return self.work_Book.named_styles()[0]

    def getCellWithThickBorder(self):
        thick_border = Side(border_style="thick", color="000000")
        self.cell_border = Border(left=thick_border,
                             right=thick_border,
                             top=thick_border,
                             bottom=thick_border)
        return self.cell_border

    def getCellWithThickThinBorder(self):
        thick_border = Side(border_style="thick", color="000000")
        thin_border = Side(border_style="thin", color="000000")
        self.cell_border = Border(left=thick_border,
                             right=thick_border,
                             top=thin_border,
                             bottom=thin_border)
        return self.cell_border

    def getCellWithThinBorder(self):
        thin_border = Side(border_style="thin", color="000000")
        self.cell_border = Border(left=thin_border,
                             right=thin_border,
                             top=thin_border,
                             bottom=thin_border)
        return self.cell_border

    def getCellWithThickBottomBorder(self):
        thick_border = Side(border_style="thick", color="000000")
        thin_border = Side(border_style="thin", color="000000")
        self.cell_border = Border(left=thick_border,
                             right=thick_border,
                             top=thin_border,
                             bottom=thick_border)
        return self.cell_border

    def getCellFont(self):
        if self.cell_font is None:
            self.setCellFont()
        return self.cell_font

    def setCellFont(self,
                    font_name = 'Calibri',
                    Size = 14,
                    bold = False,
                    italic = False,
                    strike = False,
                    underline = 'none',
                    color = '000000'):
        self.cell_font = Font(name= font_name,
                     size=Size,
                     bold=bold,
                     italic=italic,
                     strike=strike,
                     underline=underline,
                     color=color
                     )

    def getCellAlignment(self):
        if self.cell_alignment is None:
            self.setCellAlignment()
        return  self.cell_alignment

    def setCellAlignment(self):
        self.cell_alignment = Alignment(horizontal="center",
                                   vertical="center",
                                   wrapText=True,
                                   shrinkToFit=True)

    def getPatternFill(self):
        if self.cell_pattern_fill is None:
            self.setPatternFill()
        return self.cell_pattern_fill

    def setPatternFill(self,type='solid',fgcolor = '00FFFFFF', bgcolor = '00FFFFFF' ):
        self.cell_pattern_fill = PatternFill(fill_type=type,fgColor=fgcolor,bgColor=bgcolor)

    def getXlsxFilePath(self):
        return self.xlsx_File_Path

