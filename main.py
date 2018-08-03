import SheetUtil, configparser, ast, os, shutil, calendar


def enum(**enums):
    return type('Enum', (), enums)

#Load config file
CONFIG = configparser.ConfigParser()
CONFIG.read('CONFIG.ini')

#Load variables from config file
DIRECTORY_PATH = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH")
DIRECTORY_PATH_TURNS = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_TURNS")
DIRECTORY_PATH_MAINLINE = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_MAINLINE")
DIRECTORY_PATH_DUPICATES = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_DUPICATES")
MAINLINE_TEMPLATE = CONFIG.get("BASE_FILES", "MAINLINE_TEMPLATE")
MAINLINE_TEMPLATE2 = CONFIG.get("BASE_FILES", "MAINLINE_TEMPLATE2")
MAINLINE_TEMPLATE3 = CONFIG.get("BASE_FILES", "MAINLINE_TEMPLATE3")
MAINLINE_TEMPLATE4 = CONFIG.get("BASE_FILES", "MAINLINE_TEMPLATE4")
MAINLINE_KEYWORDS = ast.literal_eval(CONFIG.get("FILTER_KEYWORDS", "MAINLINE_KEYWORDS"))
TURNING_KEYWORDS = ast.literal_eval(CONFIG.get("FILTER_KEYWORDS", "TURNING_KEYWORDS"))
GOOGLEMAP_APIKEY = CONFIG.get("API_KEY", "GOOGLEMAP_APIKEY")
DIRECTORY_PATH_FORMATTED_MAINLINE = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_FORMATTED_MAINLINE")
DIRECTORY_PATH_ORIGINAL = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_ORIGINAL")
DIRECTORY_PATH_NAME_ERROR = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_NAME_ERROR")
DIRECTORY_PATH_NAME_ERROR_ORIG = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_NAME_ERROR_ORIG")
DIRECTORY_PATH_MULTIDAY = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_MULTIDAY")
DIRECTORY_PATH_MULTIDAY_ORIG = CONFIG.get("FILE_PATHS", "DIRECTORY_PATH_MULTIDAY_ORIG")
fileDir = os.listdir(DIRECTORY_PATH)
mainlineDir = os.listdir(DIRECTORY_PATH_MAINLINE)
formattedMainline = os.listdir(DIRECTORY_PATH_FORMATTED_MAINLINE)


MONTHS = {v: k for k, v in enumerate(calendar.month_name)}
fileType = enum(IDAX='IDAX', ADT='ADT', COUNTSUNLIMITED='CountsUnlimited')

#Basic template to format files
class MainlineFormatShell():
    def __init__(self, filename, filterKeyword, fileDir, fileType, dateKeyword):
        self.filename = filename
        self.workbook = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH_MAINLINE, filename))
        self.sheet = self.workbook.getSheet(0)
        self.filterKeyword = filterKeyword
        self.fileType = fileType
        self.dateKeyword = dateKeyword
        self.numberOfdays = self.workbook.checkNumberInstances(self.dateKeyword)

    #Function to execute everything else, call this function. do not directly invoke the other ones
    #This picks whether or not the file is a "multi-day"
    def execute(self):
        if self.numberOfdays == 1:
            self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE)
            self.excelWrite = SheetUtil.excelWrite(self.newBook)
            self.filterShell()

        elif self.numberOfdays > 1:
            self.workbook = SheetUtil.multiExcelUtil(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename))
            self.multiFilterShell()

    #Returns bounds in array -> [rowMin, rowMax, colMin, colMax]
    #Parameters to_someVal = the coordinate of the location cell, add_someVal = what to add to get the correct bounds of ata
    def getBounds(self, to_minRow, to_maxRow, to_minCol, to_maxCol, add_minRow, add_maxRow, add_minCol, add_maxCol):
        rowMin = to_minRow + add_minRow
        rowMax = to_maxRow + add_maxRow
        colMin = to_minCol + add_minCol
        colMax = to_maxCol + add_maxCol

        return [rowMin, rowMax, colMin, colMax]


    #Basic template to format files with single dates
    def filterShell(self):
        if self.workbook.findCell([self.filterKeyword])[0] == True:
            print 'Entering ' + self.filename
            self.location = self.workbook.getLocation(self.fileType)
            self.DATE = self.workbook.getDate(self.dateKeyword)
            self.mapUtil = SheetUtil.mapUtil(self.location, GOOGLEMAP_APIKEY)
            self.newFilename = self.mapUtil.mainlineNaming()
            self.inSoMa = self.mapUtil.SoMaCheck()
            self.DIRECTIONS = self.workbook.findDirectionCell()
            print self.DIRECTIONS

            filledCols = 0

            for key in self.DIRECTIONS:
                if self.DIRECTIONS.get(key)[0]:

                    dir = key

                    #Get bounds for the data extraction
                    to_rowMin = self.DIRECTIONS.get(dir)[1]
                    to_rowMax = self.DIRECTIONS.get(dir)[1]
                    to_colMin = self.DIRECTIONS.get(dir)[2]
                    to_colMax = self.DIRECTIONS.get(dir)[2]

                    if self.fileType == 'ADT':
                        if self.workbook.checkEmptyCell(self.DIRECTIONS.get(dir)[1], self.DIRECTIONS.get(dir)[2] - 2):
                            #Shouldn't hardcode the addition/subtraction values
                            #Put this data in config file later
                            bounds = self.getBounds(to_rowMin, to_rowMax, to_colMin, to_colMax, 3, 27, -2, 2)

                        elif self.workbook.checkEmptyCell(self.DIRECTIONS.get(dir)[1], self.DIRECTIONS.get(dir)[2] - 1):
                            bounds = self.getBounds(to_rowMin, to_rowMax, to_colMin, to_colMax, 3, 27, -1, 3)

                    else:
                        bounds = self.getBounds(to_rowMin, to_rowMax, to_colMin, to_colMax, 3, 27, -1, 3)

                    self.DATA = self.workbook.getData(bounds[0], bounds[1], bounds[2], bounds[3])

                    self.excelWrite.inputData(0, self.DATA[0], 1, 1, filledCols)

                    if self.inSoMa:
                        dir = self.mapUtil.directionFix(key)
                    else:
                        dir = key

                    if self.DATA[1] == False:
                        self.excelWrite.write(0, 1, 1 + filledCols, dir)
                        filledCols = filledCols + 1


            self.excelWrite.write(0, 1, 5, '')
            self.excelWrite.write(1, 0, 0, self.filename)
            self.excelWrite.changeSheetName(0, self.DATE )
            self.newWorkbookSave = self.excelWrite.getWorkbook()



            if 'NAME_ERROR' in str(self.newFilename):
                self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_NAME_ERROR, self.newFilename))
                shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_NAME_ERROR_ORIG)

            elif 'MULTI_DAY' in str(self.newFilename):
                self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_MULTIDAY, self.newFilename))
                shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_MULTIDAY_ORIG)
            else:
                self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_FORMATTED_MAINLINE, self.newFilename))
                shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_ORIGINAL)


    #Basic template to format files with multiple days
    def multiFilterShell(self):
            print('Entering ' + self.filename)
            self.DIRECTIONS = self.workbook.findDirectionCell()
            print(self.DIRECTIONS)

            dates = self.workbook.getAllInstances([self.dateKeyword])
            formattedDates = []
            indexesToRemove = []

            for index, dateRow in enumerate(dates):
                f = dateRow[0]
                tempDate = self.workbook.getRightCell(dateRow[0]).value

                if tempDate != '-':
                    formattedDates.append(SheetUtil.dateFormat(tempDate))
                elif tempDate == '-':
                    dates.remove(dateRow)
                    indexesToRemove.append(index)

            # Create new workbook
            if dates.__len__() == 1:
                self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE)
                self.excelWrite = SheetUtil.excelWrite(self.newBook)

            if dates.__len__() == 2:
                self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE2)
                self.excelWrite = SheetUtil.excelWrite(self.newBook)

            if dates.__len__() == 3:
                self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE3)
                self.excelWrite = SheetUtil.excelWrite(self.newBook)

            if dates.__len__() == 4:
                self.newBook = self.workbook.createNewWorkbook(MAINLINE_TEMPLATE4)
                self.excelWrite = SheetUtil.excelWrite(self.newBook)

            for key in self.DIRECTIONS:
                if self.DIRECTIONS.get(key)[0]:
                    for index in indexesToRemove:
                        del self.DIRECTIONS.get(key)[1][index]

            filledCols = 0

            for key in self.DIRECTIONS:
                if self.DIRECTIONS.get(key)[0]:

                    dir = key
                    for index, coord in enumerate(self.DIRECTIONS.get(key)[1]):
                        to_rowMin = coord[0]
                        to_rowMax = coord[0]
                        to_colMin = coord[1]
                        to_colMax = coord[1]

                        if self.fileType == 'IDAX':
                            bound = self.getBounds(to_rowMin, to_rowMax, to_colMin, to_colMax, 3, 27, -2, 2)
                        elif self.fileType == 'ADT':
                            bound = self.getBounds(to_rowMin, to_rowMax, to_colMin, to_colMax, 3, 27, -1, 3)

                        self.DATA = self.workbook.getData(bound[0], bound[1], bound[2], bound[3])
                        validData = SheetUtil.dayValidation(self.DATA[0])
                        print(validData)

                        if validData:
                            sheet = self.excelWrite.getSheet(index)
                            self.excelWrite.changeSheetName(index, formattedDates[index])
                            self.excelWrite.write(index, 1, 1 + filledCols, dir)
                            self.excelWrite.inputData(index, self.DATA[0], 1, 1, filledCols)
                            self.excelWrite.write(index, 1, 5, '')

            if self.fileType == 'ADT':
                self.location = self.workbook.getLocation(self.fileType)
                self.mapUtil = SheetUtil.mapUtil(self.location, GOOGLEMAP_APIKEY)
                self.newFilename = self.mapUtil.mainlineNaming()

            else:
                self.newFilename = 'formatted ' + self.filename + '.xls'
            self.excelWrite.write(dates.__len__(), 0, 0, self.filename)
            self.newWorkbookSave = self.excelWrite.getWorkbook()

            if 'formatted' in str(self.newFilename):
                self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_MULTIDAY, self.newFilename))
                shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_MULTIDAY_ORIG)
            else:
                self.newWorkbookSave.save(os.path.join(DIRECTORY_PATH_FORMATTED_MAINLINE, self.newFilename))
                shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, self.filename), DIRECTORY_PATH_ORIGINAL)


#To validate files
class ValidationCheck():
    def __init__(self, filename):
        self.filename = filename
        self.workbook = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH_FORMATTED_MAINLINE, filename))
        print('Checking ' + filename)

    def checkEmptyWorkbook(self):
        numberSheets = self.workbook.getNumberSheets()

        for sheet in range(numberSheets - 1):
            self.workbook.getSheet(sheet)

            try:
                if self.workbook.checkEmptyCell(1, 1):
                    print('Empty!')
            except(IndexError):
                print('Error: Empty')







def excelSort():
    # Loops through all files in directory
    for filename in fileDir:

        # Check if the file is an excel file
        if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith('XLSX'):

            excelTool = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH, filename))
            sheet = excelTool.getSheet(0)

            # Checks for turning keywords to confirm it is a turning file and moves file to turns
            if (excelTool.findCell(TURNING_KEYWORDS)[0]):
                print("Moved " + filename + " to turns folder")
                shutil.move(os.path.join(DIRECTORY_PATH, filename), DIRECTORY_PATH_TURNS)

            # Checks for mainine keywords to confirm it is a mainline file and moves file to mainline
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
        CU = MainlineFormatShell(filename, MAINLINE_KEYWORDS[1], mainlineDir, fileType.COUNTSUNLIMITED, 'Date:')
        CU.execute()


def ADTFormat():
    for filename in mainlineDir:
        adt = MainlineFormatShell(filename, MAINLINE_KEYWORDS[0], mainlineDir, fileType.ADT, 'DATE:')
        adt.execute()


def IDAXFormat():
    for filename in mainlineDir:
        idax = MainlineFormatShell(filename, MAINLINE_KEYWORDS[2], mainlineDir, fileType.IDAX, 'DATE:')
        idax.execute()

def checkEmptyFiles():
    for filename in formattedMainline:
        valid = ValidationCheck(filename)
        valid.checkEmptyWorkbook()

checkEmptyFiles()