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
fileDir = os.listdir(DIRECTORY_PATH)
mainlineDir = os.listdir(DIRECTORY_PATH_MAINLINE)
MONTHS = {v: k for k, v in enumerate(calendar.month_name)}

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

def ADTformat():
    #Loop throuh all excel workbooks
    for filename in mainlineDir:
        if filename.endswith(".xls"):
            workbook = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH_MAINLINE, filename))
            sheet = workbook.getSheet(0)


            if workbook.findCell(MAINLINE_KEYWORDS[0])[0] == True:
                print 'Entering ' + filename
                #Get data
                location = workbook.getLocation()
                DATE = workbook.getDate('DATE:')

                #Initilaze instance of mapUtils
                mapUtil = SheetUtil.mapUtil(location, GOOGLEMAP_APIKEY)
                mainline = mapUtil.getMainline()
                street1 = mapUtil.getStreet1()
                street2 = mapUtil.getStreet2()
                newFilename = mapUtil.mainlineNaming()


                inSoMa = mapUtil.SoMaCheck()
                DIRECTIONS = workbook.findDirectionCell()

                newBook = workbook.createNewWorkbook(MAINLINE_TEMPLATE)
                excelWrite = SheetUtil.excelWrite(newBook, DATE)

                #Data sheet index 0, source sheet index 1


                filledCols = 0

                for key in DIRECTIONS:
                    if DIRECTIONS.get(key)[0] == True:

                        dir = key

                        if workbook.checkEmptyCell(DIRECTIONS.get(dir)[1], DIRECTIONS.get(dir)[2] - 2):
                            rowMin = DIRECTIONS.get(dir)[1] + 3
                            rowMax = DIRECTIONS.get(dir)[1] + 3 + 24
                            colMin = DIRECTIONS.get(dir)[2] - 2
                            colMax = DIRECTIONS.get(dir)[2] + 2

                        elif workbook.checkEmptyCell(DIRECTIONS.get(dir)[1], DIRECTIONS.get(dir)[2] - 1):
                            rowMin = DIRECTIONS.get(dir)[1] + 3
                            rowMax = DIRECTIONS.get(dir)[1] + 3 + 24
                            colMin = DIRECTIONS.get(dir)[2] - 1
                            colMax = DIRECTIONS.get(dir)[2] + 3

                        #Maybe remove SoMa check
                        if inSoMa:
                            dir = mapUtil.directionFix(dir)


                        DATA = workbook.getData(rowMin, rowMax, colMin, colMax)

                        excelWrite.inputData(0, DATA[0], 1, 1, filledCols)

                        if DATA[1] == False:
                            excelWrite.write(0, 1, 1 + filledCols, dir)
                            filledCols = filledCols + 1


                excelWrite.write(0, 1, 5, '' )
                excelWrite.write(1, 0, 0, filename)
                newWorkbookSave = excelWrite.getWorkbook()

                newWorkbookSave.save(os.path.join(DIRECTORY_PATH_FORMATTED_MAINLINE, newFilename))

                #TODO: Add these file paths to config folder
                if 'NAME_ERROR' not in newFilename:
                    shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE, filename),
                                DIRECTORY_PATH_ORIGINAL)
                else:
                    shutil.move(os.path.join(DIRECTORY_PATH_MAINLINE,filename), DIRECTORY_PATH_ORIGINAL)

def CountsUnlimitedFormat():
    for filename in mainlineDir:
        if filename.endswith('xlsx') or filename.endswith('XLSX'):
            workbook = SheetUtil.excelUtil(os.path.join(DIRECTORY_PATH_MAINLINE, filename))
            sheet = workbook.getSheet(0)

            if workbook.findCell([MAINLINE_KEYWORDS[1]])[0]:
                print 'Entering ' + filename
                DATE = workbook.getDate('Date:')
                print(DATE)


CountsUnlimitedFormat()
