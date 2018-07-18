############################################################################################
#                                                                                          #
#     #####  ####### ####### ######      #####  #     # #     # #     # ####### ######     #
#    #     # #     # #     # #     #    #     # #     # ##   ## ##   ## #       #     #    #
#    #       #     # #     # #     #    #       #     # # # # # # # # # #       #     #    #
#    #  #### #     # #     # #     #     #####  #     # #  #  # #  #  # #####   ######     #
#    #     # #     # #     # #     #          # #     # #     # #     # #       #   #      #
#    #     # #     # #     # #     #    #     # #     # #     # #     # #       #    #     #
#     #####  ####### ####### ######      #####   #####  #     # #     # ####### #     #    #
#                                                                                          #
#                                                                                          #
############################################################################################
#################################Written by Donald Chung####################################

"""
   TODO

   Flagging possible errors
   Create config file
   Optimize Code
   Make Code OOP
   Write GUI

"""

#Imports
import shutil, os, xlrd, xlwt, googlemaps,re, math, configparser
from xlrd import *
from xlwt import easyxf
from xlutils.copy import copy
from collections import defaultdict
from collections import deque
import calendar


#File paths
DIR_PATH = 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\\'
DIR_PATH_TURNS = 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\pre_ctdrac_turning'
DIR_PATH_MAINLINE = 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\pre_ctdrac_mainline'
DIR_PATH_DUPE = 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\pre_ctdrac_duplicate_ignore'
fileDir = os.listdir(DIR_PATH)

#Template Files
MAINLINE_TEMPLATE = 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\\templates\MAINLINE_TEMPLATE.xlsx'

#Number - Month conversions
MONTHS = {v: k for k,v in enumerate(calendar.month_name)}

#Respective Keywords for turning/mainline
TURNING_KEYWORDS = ['TURNING MOVEMENT COUNT', 'Note: U-Turn volumes for bikes are included in Left-Turn, if any.',
                    'I N T E R S E C T I O N   T U R N I N G   M O V E M E N T   S U M M A R Y', 'CROSSING GUARD COUNT',
                    'Turning Movement Count']
MAINLINE_KEYWORDS = ['24-HOUR ADT COUNT SUMMARY', 'Counts Unlimited, Inc.', 'IDAX 24-HOUR ADT COUNT SUMMARY',
                     '24 Hour Directional Volume Count']

#Google Maps API Key
GOOGLE_MAPS = googlemaps.Client(key='AIzaSyC1xt0nxhwwMxmWPIHgsMygniUZfd7sU04')



#Sorts files into turns and mainline data
def excelSort():

    #Loops through all files in directory
    for filename in fileDir:

        #Check if the file is an excel file
        if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith('XLSX'):

            #Opens file and goes to the first sheet
            workbook = open_workbook(os.path.join(DIR_PATH, filename))
            sheet = workbook.sheet_by_index(0)

            #Checks for turning keywords to confirm it is a turning file and moves file to turns
            if (findCell(sheet, TURNING_KEYWORDS)[0]):
                print("Moved " + filename + " to turns folder")
                shutil.move(os.path.join(DIR_PATH, filename), DIR_PATH_TURNS)

            #Checks for mainine keywords to confirm it is a mainline file and moves file to mainline
            elif (findCell(sheet, MAINLINE_KEYWORDS)[0]):
                print("Moved " + filename + " to mainline folder")
                shutil.move(os.path.join(DIR_PATH, filename), DIR_PATH_MAINLINE)


#Remove PDF duplicates of files
def pdfDuplicates():

    #Loops through all files in directory and looks for PDF files
    for filename in fileDir:
        if filename.endswith(".pdf"):

            #Store all instances of underscore in filename
            baseUnderscore = findCharInString(filename, '_')

            #Loops through file directory again and find another unique file
            for matchFile in fileDir:
                if filename != matchFile:

                    #Compares substrings, splitting by underscores in filename, need 3 matches to be considered duplicates
                    matchFound = 0  # Need 3 for a 'match'
                    for endIndex in baseUnderscore:
                        if filename[: endIndex] == matchFile[: endIndex]:
                            matchFound = matchFound + 1
                        else:
                            break

                        #If three matches are found, compares filename lengths to decide the movements
                        if matchFound == 3:
                            try:
                                if filename.__len__() > matchFile.__len__():
                                    print("Moved " + matchFile + " to duplicates folder")
                                    shutil.move(os.path.join(DIR_PATH, matchFile), DIR_PATH_DUPE)
                                elif filename.__len__() < matchFile.__len__():
                                    print("Moved " + filename + " to duplicates folder")
                                    shutil.move(os.path.join(DIR_PATH, filename), DIR_PATH_DUPE)
                            except Exception as err:
                                print("Unable to move file " + str(err))

#Unsupported - Files with multiple days
#Supports ADT 24 Hour Count file format
def ADTdracFormat():
            i = 0

            #Loops through all files in directory to find workbooks
            for filename in fileDir:

                if filename.endswith(".xlsx") or filename.endswith(".xls") or filename.endswith('XLSX'):

                    workbook = open_workbook(os.path.join(DIR_PATH, filename))
                    sheet = workbook.sheet_by_index(0)

                    #Searches for ADT Keyword to confirm valid filetype
                    if findCell(sheet, MAINLINE_KEYWORDS[0]):
                        print 'Entering ' + filename

                        #Variable declerations

                        #Stores the row col location in array to store in dictionary
                        NBCoords = [False]
                        SBCoords = [False]
                        WBCoords = [False]
                        EBCoords = [False]
                        DIRECTIONS = {'NB': NBCoords, 'SB': SBCoords, 'EB': EBCoords, 'WB': WBCoords}

                        #Data to be added to new spreadsheet
                        DATA = deque([])

                        # Get date
                        dateRow = findCell(sheet, ['DATE:'])[1]
                        DATE = getRightCell(sheet, dateRow).value
                        DATE = dateFormat(DATE)

                        #Get Location
                        locationRow = findCell(sheet, ['LOCATION:'])[1]
                        LOCATION = getRightCell(sheet, locationRow).value

                        #rename file
                        newFilename = str(newNaming(LOCATION)[0])

                        #Used to heck if the street is in SoMa grid
                        tempMainline = newNaming(LOCATION)[1]
                        tempStreet = newNaming(LOCATION)[2]
                        tempStreet2 = newNaming(LOCATION)[3]

                        #T/F in SoMa Grid
                        inSoMa = SoMaCheck(tempMainline, tempStreet)


                        # Get location of the NB/SB cells etct
                        if findCell(sheet, ['NB Total Volume', 'NB Total Vol', 'NB']):
                            NBCoords[0] = True
                            NBCoords.append(findCell(sheet, ['NB Total Volume', 'NB Total Vol', 'NB'])[1])
                            NBCoords.append(findCell(sheet, ['NB Total Volume', 'NB Total Vol', 'NB'])[2])
                        if findCell(sheet, ['SB Total Volume', 'SB Total Vol', 'SB']):
                            SBCoords[0] = True
                            SBCoords.append(findCell(sheet, ['SB Total Volume', 'SB Total Vol', 'SB'])[1])
                            SBCoords.append(findCell(sheet, ['SB Total Volume', 'SB Total Vol', 'SB'])[2])
                        if findCell(sheet, ['WB Total Volume', 'WB Total Vol', 'WB']):
                            WBCoords[0] = True
                            WBCoords.append(findCell(sheet, ['WB Total Volume', 'WB Total Vol', 'WB'])[1])
                            WBCoords.append(findCell(sheet, ['WB Total Volume', 'WB Total Vol', 'WB'])[2])
                        if findCell(sheet, ['EB Total Volume', 'EB Total Vol', 'EB']):
                            EBCoords[0] = True
                            EBCoords.append(findCell(sheet, ['EB Total Volume', 'EB Total Vol', 'EB'])[1])
                            EBCoords.append(findCell(sheet, ['EB Total Volume', 'EB Total Vol', 'EB'])[2])
                        print DIRECTIONS

                        # Write new file
                        templateWorkbook = open_workbook(MAINLINE_TEMPLATE)
                        templateSheet = templateWorkbook.sheet_by_index(0)
                        tempBook = copy(templateWorkbook)
                        tempSheet = tempBook.get_sheet(0)
                        tempBook.get_sheet('YYYY.MM.DD').name = DATE
                        sourceSheet = tempBook.get_sheet(1)

                        filledCols = 0

                        for key in DIRECTIONS:
                            if DIRECTIONS.get(key)[0] == True:
                                #Temp Var to hold the direction
                                dir = key

                                #Checks the next empty right cell in the template sheet
                                if emptyLeftCell(templateSheet, 1)[0]:
                                    coords = emptyLeftCell(templateSheet, 1)

                                    tempSheet.write(coords[1], coords[2] + filledCols, '')

                                    #Check if left cell of direction cell is empty, then get bounds of data
                                    if checkEmptyCell(sheet, DIRECTIONS.get(dir)[1], DIRECTIONS.get(dir)[2] - 2):
                                        rowMin = DIRECTIONS.get(dir)[1] + 3
                                        rowMax = DIRECTIONS.get(dir)[1] + 3 + 24
                                        colMin = DIRECTIONS.get(dir)[2] - 2
                                        colMax = DIRECTIONS.get(dir)[2] + 2


                                    elif checkEmptyCell(sheet, DIRECTIONS.get(dir)[1], DIRECTIONS.get(dir)[2] - 1):
                                        rowMin = DIRECTIONS.get(dir)[1] + 3
                                        rowMax = DIRECTIONS.get(dir)[1] + 3 + 24
                                        colMin = DIRECTIONS.get(dir)[2] - 1
                                        colMax = DIRECTIONS.get(dir)[2] + 3

                                    #work, add whitelist
                                    if inSoMa == True:
                                        dir = directionFix(tempMainline, tempStreet, tempStreet2, dir)
                                        if dir == 'err':
                                            if dir == 'WB':
                                                dir = 'SB'
                                            elif dir == 'EB':
                                                dir = 'NB'
                                            elif dir == 'NB':
                                                dir = 'EB'
                                            elif dir == 'SB':
                                                dir = 'WB'

                                    #Loop to copy the data
                                    #FLAG var for empty datasets, does't copy the data if blank
                                    for row in range(rowMin, rowMax):
                                        FLAGGED = False
                                        for col in range(colMin, colMax):
                                            if checkEmptyCell(sheet, row, col):
                                                FLAGGED = True
                                                break
                                            DATA.append(sheet.cell(row, col).value)

                                        if FLAGGED:
                                            break


                                    #Writes copied data into new sheet
                                    done = 1
                                    print DATA
                                    while DATA.__len__() > 0:
                                        tempSheet.write(coords[1] + done, coords[2] + filledCols, DATA.popleft())
                                        done = done + 1


                                #If not Flagged then write in the direction at the top row
                                if FLAGGED == False:
                                    tempSheet.write(coords[1], coords[2] + filledCols, dir)
                                    filledCols = filledCols + 1

                        #Write the source sheet and remove marker
                        #Save newly created workbook and move the original file
                        sourceSheet.write(0,0, filename)
                        tempSheet.write(1, 5, '')
                        tempBook.save(os.path.join(DIR_PATH, newFilename))

                      #  if 'NAME_ERROR' not in newFilename:
                      #      shutil.move(os.path.join(DIR_PATH, filename), 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\Working File Formats')
                      #  else:
                      #      shutil.move(os.path.join(DIR_PATH,filename), 'Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\Needs Review')




                        print "Success with " + filename

                        i = i + 1



##########################################################
#                                                        #
#                     HELPER FUNCTIONS                   #
#                                                        #
##########################################################


# Find all indexes of single characters in a string
#Parameters: String, Character to find
def findCharInString(s, ch):
    return [i for i, ltr in enumerate(s) if ltr == ch]

# string var must be an array
#finds cell matching to a string exactly
#Parameters: Sheet of interest, string to find
def findCell(sheet, string):
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            for text in string:
                if sheet.cell(row, col).value == text:
                    return [True, row, col]

    return False

#Check next cell is not empty and returns the data
#Parameters: Sheet of interest, row to search
def getRightCell(sheet, row):
    data = ''

    for col in range(sheet.ncols):
        if sheet.cell_type(row, col) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            data = sheet.cell(row, col)

    return data

#Mostly used for mainline sheets
#Finds the next empty cell
#Parameters: Sheet of interest, row to search
def emptyLeftCell(sheet, row):
    col = 0
    while col < 5:
        if sheet.cell_type(row, col) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            return [True, row, col]
        else:
            col = col + 1

#Check if cell is empty
def checkEmptyCell(sheet, row, col):
    if sheet.cell_type(row, col) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return True
    else:
        return False

#Formats 'dayname', 'Month Name' 'Number', Year to YYYY.MM.DD
#eg. Tuesday, November 23, 2016 to 2016.11.23
#Parameters: Date in words format
def dateFormat(DATE):

    #Splitting substring
    tempDate = DATE[DATE.find(',') + 2:]
    year = tempDate[tempDate.find(',') + 2:]
    month = tempDate[:tempDate.find(' ')]
    day = tempDate[tempDate.find(' ') + 1:tempDate.find(',')]

    #Add extra zero in front if the values are less than 10
    month = MONTHS[month]
    if int(month) < 10:
        month = '0' + str(month)

    if int(day) < 10:
        day = '0' + str(day)
    formattedDate = str(year) + '.' + str(month) + '.' + str(day)

    return formattedDate

#Find all instances of substring in one string
#Return type: Array
def find_all(string, substring):
    return [m.start() for m in re.finditer(substring, string)]


def newNaming(location):
    isVertical = False

    StringsToRemove = ['Street', 'between', ' and', 'Bet.', '&', 'Avenue', 'between', 'and', 'San Francisco', ' St', 'Ave', 'Blvd', 'Boulevard']

    for string in StringsToRemove:
        location = location.replace(string, '')

    mainline = location[:location.find(' ')]
    location = location[location.find(' '):].strip()
    street1 = location[:location.find(' ')].strip()
    street2 = location[location.find(' '):].strip()

    geocode_result1 = GOOGLE_MAPS.geocode(mainline + ' and ' + street1 + ', San Francisco, California')[0]
    geocode_result2 = GOOGLE_MAPS.geocode(mainline + ' and ' + street2 + ', San Francisco, California')[0]


    lat_1 = geocode_result1['geometry']['location']['lat']
    lng_1 = geocode_result1['geometry']['location']['lng']
    lat_2 = geocode_result2['geometry']['location']['lat']
    lng_2 = geocode_result2['geometry']['location']['lng']


    if SoMaCheck(mainline, street1) == True:
        originLat = (lat_1 + lat_2) / 2
        originLng = (lng_1 + lng_2) / 2
        origin = (originLat, originLng)
        point1 = (lat_1, lng_1)
        point2 = (lat_2, lng_2)

        rotated1 = rotate(origin, point1, 0.785398 )
        rotated2 = rotate(origin, point2, 0.785398)
        lat_1 = rotated1[0]
        lng_1 = rotated1[1]
        lat_2 = rotated2[0]
        lng_2 = rotated2[1]


    lat_difference = abs(lat_2 - lat_1)
    lng_difference = abs(lng_2 - lng_1)


    if lat_difference > lng_difference:
        isVertical = True

    if isVertical:
        if lat_1 > lat_2:
            filename = mainline + '_' + street1 + '.' + street2
        else:
            filename = mainline + '_' + street2 + '.' + street1

    else:
        if lng_1 < lng_2:
            filename = mainline + '_' + street1 + '.' + street2
        else:
            filename = mainline + '_' + street2 + '.' + street1

    filename = filename + '.xls'

    return [filename, mainline, street1, street2]



#Checks if the street is valid name
def streetValidation(street):
    streetBook = open_workbook('Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\Streets\street_extract.xlsx')
    sheet = streetBook.sheet_by_index(0)
    for row in range(sheet.nrows):

        if street.lower().rstrip() == sheet.cell(row, 0).value.lower():
            return True

    return False

#Checks if intersection is in the SoMa Grid
def SoMaCheck(street1, street2):
    MAX_LAT = 37.795365
    MIN_LNG = -122.423582
    MIN_LAT = 37.768014
    MAX_LNG = -122.387442


    geocode_result1 = GOOGLE_MAPS.geocode(street1 + ' and ' + street2 + ', San Francisco, California')[0]
    lat = geocode_result1['geometry']['location']['lat']
    lng = geocode_result1['geometry']['location']['lng']

    if lat < MAX_LAT and lat > MIN_LAT and lng > MIN_LNG and lng < MAX_LNG:

        y = (1.26787 * lat) -170.313

        if y < lng:
            return True


    return False

#Works in SoMa, but does not work in irregular grids
#Return true if NB/SB, false if EB/WB
def directionFix(mainline, street1, street2, dir):
    #Clear whitespace
    mainline = mainline.replace(' ', '')
    street1 = street1.replace(' ', '')
    street2 = street2.replace(' ', '')

    streetBook = open_workbook('Q:\Data\Observed\Streets\Counts\CtDrac2016_Donald\Streets\SoMa_Directions.xls')
    sheet = streetBook.sheet_by_index(0)
    isVertical = False

    geocode_result1 = GOOGLE_MAPS.geocode(mainline + ' and ' + street1 + ', San Francisco, California')[0]
    geocode_result2 = GOOGLE_MAPS.geocode(mainline + ' and ' + street2 + ', San Francisco, California')[0]

    lat1 = geocode_result1['geometry']['location']['lat']
    lng1 = geocode_result1['geometry']['location']['lng']
    lat2 = geocode_result2['geometry']['location']['lat']
    lng2 = geocode_result2['geometry']['location']['lng']

    #Repackage into seperate method later
    if SoMaCheck(mainline, street1) == True:
        originLat = (lat1 + lat2) / 2
        originLng = (lng1 + lng2) / 2
        origin = (originLat, originLng)
        point1 = (lat1, lng1)
        point2 = (lat2, lng2)

        rotated1 = rotate(origin, point1, 0.785398 )
        rotated2 = rotate(origin, point2, 0.785398)
        lat1 = rotated1[0]
        lng1 = rotated1[1]
        lat2 = rotated2[0]
        lng2 = rotated2[1]

        lat_difference = abs(lat2 - lat1)
        lng_difference = abs(lng2 - lng1)

        if lat_difference > lng_difference:
            isVertical = True

        for row in range(sheet.nrows):
            if sheet.cell(row, 0).value == mainline:
                if sheet.cell(row, 1).value == 'No':
                    if (isVertical and (dir == 'EB' or dir == 'WB')) or (not isVertical and (dir =='NB' or dir =='SB')):
                        return 'err'
                else:
                    return sheet.cell(row, 2).value


#Written by Mark Dickinson
#https://stackoverflow.com/questions/34372480/rotate-point-about-another-point-in-degrees-python
def rotate(origin, point, angle):
    """
    Rotate a point counterclockwise by a given angle around a given origin.

    The angle should be given in radians.
    """
    ox, oy = origin
    px, py = point

    qx = ox + math.cos(angle) * (px - ox) - math.sin(angle) * (py - oy)
    qy = oy + math.sin(angle) * (px - ox) + math.cos(angle) * (py - oy)
    return qx, qy

#newNaming('16th St Bet. Folsom St & Shotwell St')
#print newNaming('Hayes Street between Pierce Street and Steiner Street  San Francisco')
#print newNaming('9th Street between Folsom Street and Harrison Street ')
ADTdracFormat()

#Main method
module = int(intput("To run a function, enter the corresponding number\n"
                    "1. Sort "))



