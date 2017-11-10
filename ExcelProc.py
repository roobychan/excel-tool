import glob
import xlrd
# import csv
import xlwt


class ExcelProc(object):
    """docstring for ExcelProc"""
    fileList = []
    dataTab = [[], [], [], [], []]
    path = 'c:/temp/test/'
    langDict = {'THA': 'TH', 'RUS': 'RU',
                'KOR': 'KO', 'VIT': 'VI', 'JPN': 'JA',
                'DEU': 'DE', 'CHS': 'ZH', 'CHT': 'ZF',
                'CSY': 'CS', 'DAN': 'DA', 'ESP': 'ES',
                'FIN': 'FI', 'FRA': 'FR', 'HUN': 'HU',
                'ITA': 'IT', 'NLD': 'NL', 'NOR': 'NO',
                'PLK': 'PL', 'PTG': 'PT', 'ROM': 'RO',
                'SKY': 'SK', 'SWE': 'SV', 'TRK': 'TR'}

    sheetDict = ['ZXADP_M99_TRANS', 'Message', 'Email', 'Navigation', 'Misc']

    def __init__(self, arg):
        super(ExcelProc, self).__init__()
        self.creatFile()

    def openFolder(self):
        self.path = input('input:')
        listpath = self.path + '*.xlsx'
        self.fileList = glob.glob(listpath)

    def saveFile(self):
        # with open("c:/Temp/test/test.csv", 'w', newline = '', encoding='utf-8') as csvFile:
        #     writer = csv.writer(csvFile, delimiter = '|')
        #     # writer.writerows(self.dataTab)
        #     for row in self.dataTab:
        #         trow = [row[0], row[1], row[2], row[3]]
        #         writer.writerow(trow)

        for y in range(0, len(self.sheetDict)):
            wbk = xlwt.Workbook()
            sheet = wbk.add_sheet(self.sheetDict[y])
            for x in range(0, len(self.dataTab[y])):
                    row = self.dataTab[y][x]
                    for i in range(0, len(row)):
                        sheet.write(x, i, row[i])
            fpath = self.path + self.sheetDict[y] + '.xls'
            wbk.save(fpath)

    def creatFile(self):
        pass

    def process(self):
        for file in self.fileList:
            lang = file[43:46]
            if lang in self.langDict:
                xls = xlrd.open_workbook(file)
                sheet = xls.sheet_by_name(self.sheetDict[0])
                i = 0
                for row in sheet.get_rows():
                    if i < 3:
                        i = i + 1
                        continue
                    dRow = [row[0].value, self.langDict[lang], row[1].value, row[3].value]
                    self.dataTab[0].append(dRow)
                sheet = xls.sheet_by_name(self.sheetDict[1])
                i = 0
                for row in sheet.get_rows():
                    if i < 2:
                        i = i + 1
                        continue
                    dRow = [self.langDict[lang], row[0].value, row[1].value, row[3].value]
                    self.dataTab[1].append(dRow)
                sheet = xls.sheet_by_name(self.sheetDict[2])
                i = 0
                for row in sheet.get_rows():
                    if i < 1:
                        i = i + 1
                        continue
                    dRow = [row[0].value, self.langDict[lang], row[2].value]
                    self.dataTab[2].append(dRow)
                sheet = xls.sheet_by_name(self.sheetDict[3])
                i = 0
                for row in sheet.get_rows():
                    if i < 1:
                        i = i + 1
                        continue
                    dRow = [row[0].value, row[1].value, self.langDict[lang], row[3].value]
                    self.dataTab[3].append(dRow)
                sheet = xls.sheet_by_name(self.sheetDict[4])
                i = 0
                for row in sheet.get_rows():
                    if i < 1:
                        i = i + 1
                        continue
                    dRow = [row[0].value, row[1].value, row[2].value,
                            self.langDict[lang], row[5].value]
                    if row[0].value != "":
                        self.dataTab[4].append(dRow)

a = ExcelProc("test")
a.openFolder()
a.process()
a.saveFile()
