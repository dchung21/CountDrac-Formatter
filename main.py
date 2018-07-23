import SheetUtil, configparser, ast, os, shutil, calendar


CONFIG = configparser.ConfigParser()
CONFIG.read('CONFIG.ini')

DIRECTORY_PATH = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH")
DIRECTORY_PATH_TURNS = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_TURNS")
DIRECTORY_PATH_MAINLINE = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_MAINLINE")
DIRECTORY_PATH_DUPICATES = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_DUPICATES")
MAINLINE_TEMPLATE = CONFIG.get("BASE_FILES", "MAINLINE_TEMPLATE")
MAINLINE_KEYWORDS = ast.literal_eval(CONFIG.get("FILTER_KEYWORDS", "MAINLINE_KEYWORDS"))
TURNING_KEYWORDS = ast.literal_eval(CONFIG.get("FILTER_KEYWORDS", "TURNING_KEYWORDS"))
GOOGLEMAP_APIKEY = CONFIG.get("API_KEY", "GOOGLEMAP_APIKEY")
DIRECTORY_PATH_FORMATTED_MAINLINE = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_FORMATTED_MAINLINE")
DIRECTORY_PATH_ORIGINAL = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_ORIGINAL" )
DIRECTORY_PATH_REVIEW = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_REVIEW")
fileDir = os.listdir(DIRECTORY_PATH)
mainlineDir = os.listdir(DIRECTORY_PATH_MAINLINE)
MONTHS = {v: k for k, v in enumerate(calendar.month_name)}

class MainlineFormatShell():
    def __init__(self, filename, filterKeyword, fileDir, isCountsUnlimited, dateKeyword, isADT):
        self.filename = filename
        self.workbook = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH_MAINLINE, filename))
        self.sheet = self.workbook.getSheet(0)
        self.filterKeyword = filterKeyword
        self.fileDir = fileDir
        self.isCountsUnlimited = isCountsUnlimited
        self.isADT = isADT
        self.dateKeyword = dateKeyword

    def filterShell(self):
                if self.workbook.findCell([self.filterKeyword])[0] == True:
                    print 'Entering ' + self.filename

                    #Get

                    self.location = self.workbook.getLocation(self.isCountsUnlimited)
                    self.DATE = self.workbook.getDate(self.dateKeyword)
                    self.mapUtil = SheetUtil.mapUtil(self.location, GOOGLEMAP_APIKEY)
                    self.newFilename = self.mapUtil.mainlineNaming()
                    self.inSoMa = self.mapUtil.SoMaCheck()
                    self.DIRECTIONS = self.workbook.findDirectionCell()
                    self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE)
                    self.excelWrite = SheetUtil.excelWrite(self.newBook, self.DATE)

                    self.filledCols = 0

                    for key in self.DIRECTIONS:
                        if self.DIRECTIONS.get(key)[0]:
                            dir = key


                            if self.isADT:
                                if self.workbook.checkEmptyCell(self.DIRECTIONS.get(dir)[1], self.DIRECTIONS.get(dir)[2] - 2):
                                    rowMin = self.DIRECTIONS.get(dir)[1] + 3
                                    rowMax = self.DIRECTIONS.get(dir)[1] + 3 + 24
                                    colMin = self.DIRECTIONS.get(dir)[2] - 2
                                    colMax = self.DIRECTIONS.get(dir)[2] + 2

                                elif self.workbook.checkEmptyCell(self.DIRECTIONS.get(dir)[1], self.DIRECTIONS.get(dir)[2] - 1):
                                    rowMin = self.DIRECTIONS.get(dir)[1] + 3
                                    rowMax = self.DIRECTIONS.get(dir)[1] + 3 + 24
                                    colMin = self.DIRECTIONS.get(dir)[2] - 1
                                    colMax = self.DIRECTIONS.get(dir)[2] + 3

                            else:
                                rowMin = self.DIRECTIONS.get(dir)[1] + 3
                                rowMax = self.DIRECTIONS.get(dir)[1] + 3 + 24
                                colMin = self.DIRECTIONS.get(dir)[2] - 1
                                colMax = self.DIRECTIONS.get(dir)[2] + 3

                            if self.inSoMa:
                                dir = self.mapUtil.directionFix(dir)

                            self.DATA = self.workbook.getData(rowMin, rowMax, colMin, colMax)

                            self.excelWrite.inputData(0, self.DATA[0], 1, 1, self.filledCols )

                            if self.DATA[1] == False:
                                self.excelWrite.write(0, 1, 1 + self.filledCols, dir)
                                self.filledCols = self.filledCols + 1

                    self.excelWrite.write(0, 1, 5, '')
                    self.excelWrite.write(1, 0, 0, self.filename)
                    self.newWorkbookSave = self.excelWrite.getWorkbook()

                    self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_FORMATTED_MAINLINE, self.newFilename))

                    # TODO: Add these file paths to config folder
                    if 'NAME_ERROR' not in self.newFilename:
                        shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename),
                                    DIRECTORY_PATH_ORIGINAL)
                    else:
                        shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_ORIGINAL)


def excelSort():

    #Loops through all files in directory
    for filename in fileDir:

        #Check if the file is an excel file
        if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith('XLSX'):

            excelTool = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH, filename))
            sheet = excelTool.getSheet(0)

            #Checks for turning keywords to confirm it is a turning file and moves file to turns
            if (excelTool.findCell(TURNING_KEYWORDS)[0]):
                print("Moved " + filename + " to turns folder")
                shutil.move(os.path.join(DIRECTORY_PATH, filename), DIRECTORY_PATH_TURNS)

            #Checks for mainine keywords to confirm it is a mainline file and moves file to mainline
            elif (excelTool.findCell(MAINLINE_KEYWORDS)[0]):
                print("Moved " + filename + " to mainline folder")
                shutil.move(os.path.join(DIRECTORY_PATH, filename), DIRECTORY_PATH_MAINLINE)

def pdfDuplicates():
    # Loops through all files in directory and looks for PDF files
    for filename in fileDir:
        if filename.endswith(".pdf"):
            baseUnderscore = SheetUtil.findCharInString(filename, '_')

            for matchFile in fileDir:
                if filename != matchFile and filename.endswith(".pdf"):

                    matchFound = 0  # Need 3 for a 'match'

                    for endIndex in baseUnderscore:
                        if filename[: endIndex] == matchFile[: endIndex]:
                            matchFound = matchFound + 1
                        else:
                            break

                        if matchFound == 3:
                            try:
                                if filename.__len__() > matchFile.__len__():
                                    print("Moved " + matchFile + " to duplicates folder")
                                    shutil.move(os.path.join(DIRECTORY_PATH, matchFile), DIRECTORY_PATH_DUPICATES)
                                elif filename.__len__() < matchFile.__len__():
                                    print("Moved " + filename + " to duplicates folder")
                                    shutil.move(os.path.join(DIRECTORY_PATH, filename), DIRECTORY_PATH_DUPICATES)
                            except Exception as err:
                                print("Unable to move file " + str(err))


def CountsUnlimitedFormat():
    for filename in mainlineDir:
        CU = MainlineFormatShell(filename, MAINLINE_KEYWORDS[1], mainlineDir, True, 'Date:', False)
        CU.filterShell()


def ADTFormat():
    for filename in mainlineDir:
        adt = MainlineFormatShell(filename, MAINLINE_KEYWORDS[0], mainlineDir, False, 'DATE:', True)
        adt.filterShell()

def IDAXFormat():
    idax = MainlineFormatShell(MAINLINE_KEYWORDS[2], mainlineDir, False, 'DATE:', False)

ADTFormat()