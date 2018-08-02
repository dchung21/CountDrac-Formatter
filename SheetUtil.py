import os, xlrd, xlwt, googlemaps, calendar, math, configparser
from xlrd import *
from xlutils.copy import copy
from collections import deque

MONTHS = {v: k for k, v in enumerate(calendar.month_name)}
CONFIG = configparser.ConfigParser()
CONFIG.read('CONFIG.ini')
SOMA_DIRECTION = CONFIG.get("BASE_FILES", "SOMA_DIRECTION")
STREETS = CONFIG.get("BASE_FILES", "STREETS")


#Handles excel reads
class excelUtil:
    DIRECTION_KEYWORDS = [['NB Total Volume', 'NB Total Vol', 'NB', 'Northbound'],
                          ['SB Total Volume', 'SB Total Vol', 'SB', 'Southbound'],
                          ['WB Total Volume', 'WB Total Vol', 'WB', 'Westbound'],
                          ['EB Total Volume', 'EB Total Vol', 'EB', 'Eastbound']
                          ]

    #Parameters: file path to excel workbook
    def __init__(self, filepath):
        self.filepath = filepath
        self.workbook = open_workbook(self.filepath)


    #Parameters: Sheet index
    #Opens and returns sheet by index
    def getSheet(self, i):
        self.sheet = self.workbook.sheet_by_index(i)
        return self.sheet

    #Parameters: String array
    #Return [0]: True/False stating if the string was found in the sheet
    #Return [1] and [2]: Row and col
    def findCell(self, string):
        for text in string:
            for row in range(self.sheet.nrows):
                for col in range(self.sheet.ncols):
                    if self.sheet.cell(row, col).value == text:
                        return [True, row, col]
        return [False]

    #Get the next not empty right cell
    #Parameters: row
    #Returns data in the cell
    def getRightCell(self, row):
        data = ''

        for col in range(self.sheet.ncols):
            if self.sheet.cell_type(row, col) not in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                data = self.sheet.cell(row, col)

        return data

    #Get the next empty right cell
    #Parameters: row
    #Return [0]: True/False stating if there are any empty cells
    #Return [1] [2]: Row and col
    #Prone to crashes bc of out of bounds exceptions
    #TODO: Fix unchecked exceptions
    def emptyRightCell(self, row):
        col = 0
        while col < 5:
            if self.sheet.cell_type(row, col) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) or sheet.cell(row, col) == '':
                return [True, row, col]
            else:
                col = col + 1

    #Check if the cell is empty
    #Parameters: Row and col
    #Return True/False if the cell really is empty
    def checkEmptyCell(self, row, col):
        if self.sheet.cell_type(row, col) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            return True
        else:
            return False

    #Gets and formats date
    #Return formatted date
    def getDate(self, dateKeyword):
        dateRow = self.findCell([dateKeyword])[1]
        DATE = self.getRightCell(dateRow)
        dateValue = DATE.value

        if DATE.ctype == 1:

            return dateFormat(dateValue)

        elif DATE.ctype == 3:
            return dateFormat(xldate_as_tuple(dateValue, self.workbook.datemode))

    #Gets location string from cell
    def getLocation(self, fileType):
        if fileType == 'ADT':
            locationRow = self.findCell(['LOCATION:'])[1]
            LOCATION = self.getRightCell(locationRow).value
            return LOCATION

        elif fileType == 'CountsUnlimited':
            mainlineRow = self.findCell(['Street:'])[1]
            streetRow = self.findCell(['Segment:'])[1]
            mainline = self.getRightCell(mainlineRow).value
            streets = self.getRightCell(streetRow).value
            return mainline + ' ' + streets

    #Finds the direction of the sheet (NB,SB,EB,WB,
    #Returns dictionary {'NB': NBCoords, 'SB': SBCoords, 'EB': EBCoords, 'WB': WBCoords}
    #eg. NBCoords would be...
    # NBCoords[0] True/False depending on if NB was found in the sheet
    # NBCoords[1] [2] Row and col of where the NB cell is
    def findDirectionCell(self):

        NBCoords = [False]
        SBCoords = [False]
        WBCoords = [False]
        EBCoords = [False]
        DIRECTIONS = {'NB': NBCoords, 'SB': SBCoords, 'EB': EBCoords, 'WB': WBCoords}
        tempDir = [NBCoords, SBCoords, WBCoords, EBCoords]


        for dir, key in zip(tempDir, excelUtil.DIRECTION_KEYWORDS):
            tempLoc = self.findCell(key)
            if tempLoc[0]:
                dir[0] = True
                dir.append(tempLoc[1])
                dir.append(tempLoc[2])

        return DIRECTIONS

    #Parameters: Filepath to template workbook
    #Copies file template
    #Returns copied workbook
    def createNewWorkbook(self, filepath):
        templateWorkbook = open_workbook(filepath)
        templateSheet = templateWorkbook.sheet_by_index(0)
        tempBook = copy(templateWorkbook)

        return tempBook

    #Gets the count data
    #Parameters: bounds of the data (rows and cols)
    def getData(self, rowMin, rowMax, colMin, colMax):
        DATA = deque([])
        FLAGGED = False

        for row in range(rowMin, rowMax):

            for col in range(colMin, colMax):
                if self.checkEmptyCell(row, col):
                    FLAGGED = True
                    break
                DATA.append(self.sheet.cell(row, col).value)

            if FLAGGED:
                break

        return [DATA, FLAGGED]


    def checkNumberInstances(self, string):
        i = 0

        for row in range(self.sheet.nrows):
            for col in range(self.sheet.ncols):
                if self.sheet.cell(row, col).value == string:
                    i = i + 1

        return i

    def dayValidation(self, data):
        if sum(data) == 0:
            return False

        return True


class multiExcelUtil(excelUtil, object):

    def __init__(self, filepath):
        super(multiExcelUtil, self).__init__(filepath)
        self.sheet = self.workbook.sheet_by_index(0)

    def getAllInstances(self, string):
        instances = []
        for text in string:
            for row in range(self.sheet.nrows):
                for col in range(self.sheet.ncols):
                    if self.sheet.cell(row, col).value == text:
                        instances.append([row, col])

        return instances

    def findDirectionCell(self):

        NBCoords = [False]
        SBCoords = [False]
        WBCoords = [False]
        EBCoords = [False]
        DIRECTIONS = {'NB': NBCoords, 'SB': SBCoords, 'EB': EBCoords, 'WB': WBCoords}
        tempDir = [NBCoords, SBCoords, WBCoords, EBCoords]

        for dir, key in zip(tempDir, excelUtil.DIRECTION_KEYWORDS):
            tempLoc = self.getAllInstances(key)
            if tempLoc.__len__() > 0:
                dir[0] = True
                dir.append(tempLoc)

        return DIRECTIONS




#Handle excel writes
class excelWrite:

    #Parameters: workbook and formatted date string
    def __init__(self, workbook):
        self.workbook = workbook

    #Writes in workbook
    #Parameters: index of sheet, row, col, string you want to input
    def write(self, sheetIndex, row, col, content):
        self.workbook.get_sheet(sheetIndex).write(row, col, content)

    #Inputs count data
    #Parameter: index of sheet, data in stack (deque), row, col, number of cols filled
    def inputData(self, sheetIndex, data, row, col, filledCol):
        print data
        done = 1
        while data.__len__() > 0:
            self.workbook.get_sheet(sheetIndex).write(row + done, col + filledCol, data.popleft())
            done = done + 1

    #Gets a sheet by index
    #Returns sheet
    def getSheet(self, sheetIndex):
        return self.workbook.get_sheet(sheetIndex)

    #Returns workbook
    def getWorkbook(self):
        return self.workbook

    def changeSheetName(self, index, name):
        sheetName = 'Sheet' + (str)(index + 1)
        self.workbook.get_sheet(sheetName).name = name



#Handles map functions
class mapUtil:

    #Parameters: location string from cell, google maps api key
    def __init__(self, location, mapsAPI):
        GOOGLE_MAPS = googlemaps.Client(key=mapsAPI)
        self.location = location

        StringsToRemove = ['Street', 'between', ' and', 'Bet.', '&', 'Avenue', 'between', 'and', 'San Francisco', ' St',
                           'Ave', 'Blvd', 'Boulevard', ' -', 'Drive', 'Way']
        for string in StringsToRemove:
            self.location = self.location.replace(string, '')

        self.mainline = self.location[:self.location.find(' ')]
        self.location = self.location[self.location.find(' '):].strip()
        self.street1 = self.location[:self.location.find(' ')].strip()
        self.street2 = self.location[self.location.find(' '):].strip()

        geocode_result1 = GOOGLE_MAPS.geocode(self.mainline + ' and ' + self.street1 + ', San Francisco, California')[0]
        geocode_result2 = GOOGLE_MAPS.geocode(self.mainline + ' and ' + self.street2 + ', San Francisco, California')[0]

        self.lat_1 = geocode_result1['geometry']['location']['lat']
        self.lng_1 = geocode_result1['geometry']['location']['lng']
        self.lat_2 = geocode_result2['geometry']['location']['lat']
        self.lng_2 = geocode_result2['geometry']['location']['lng']
        self.verticalCheck()
        self.GMAPS_API = mapsAPI

    #Returns formatted filename
    def mainlineNaming(self):
        matchedStreets = 0
        name_error = ''

        streetBook = excelUtil(STREETS)
        street = streetBook.getSheet(0)
        tempArray = [self.mainline, self.street1, self.street2]

        for string in tempArray:
                for row in range(street.nrows):
                    if string.lower() == street.cell(row, 0).value:
                        matchedStreets = matchedStreets + 1
                        break


        if matchedStreets != 3:
            name_error = "NAME_ERROR_"

        if self.isVertical:
            if self.lat_1 > self.lat_2:
                filename = self.mainline + '_' + self.street1 + '.' + self.street2
            else:
                filename = self.mainline + '_' + self.street2 + '.' + self.street1

        else:
            if self.lng_1 < self.lng_2:
                filename = self.mainline + '_' + self.street1 + '.' + self.street2
            else:
                filename = self.mainline + '_' + self.street2 + '.' + self.street1

        filename = name_error + filename + '.xls'

        return filename

    #Checks if intersection is in the SoMa grid
    #True/False respectively
    def SoMaCheck(self):
        MAX_LAT = 37.795365
        MIN_LNG = -122.423582
        MIN_LAT = 37.768014
        MAX_LNG = -122.387442


        if self.lat_1 < MAX_LAT and self.lat_1 > MIN_LAT and self.lng_1 > MIN_LNG and self.lng_1 < MAX_LNG:

            y = (1.26787 * self.lat_1) - 170.313

            if y < self.lng_1:
                return True

        return False

    #Fixes the direction if incorrect
    #TODO Fix all potential bugs
    def directionFix(self, dir):
        # Clear whitespace
        self.mainline = self.mainline.replace(' ', '')
        self.street1 = self.street1.replace(' ', '')
        self.street2 = self.street2.replace(' ', '')
        error = 'false'
        #OPTIMIZE
        streetBook = open_workbook(SOMA_DIRECTION)
        sheet = streetBook.sheet_by_index(0)

        #OPTIMIZE
        for row in range(sheet.nrows):
            if self.SoMaCheck() and sheet.cell(row, 0).value == self.mainline and sheet.cell(row, 1).value == 'Yes':
                error = sheet.cell(row, 2).value
            elif self.SoMaCheck() and sheet.cell(row, 0).value == self.mainline and  sheet.cell(row, 1).value == 'No':
                if (self.isVertical and (dir == 'EB' or dir == 'WB')) or (not self.isVertical and (dir == 'NB' or dir == 'SB')):
                    error = 'err'
            elif ((self.SoMaCheck() == False) and (self.isVertical and (dir == 'EB' or dir == 'WB')) or ((not self.isVertical and (dir == 'NB' or dir == 'SB')))):
                error = 'err'

        if error == 'err':
            if dir == 'WB':
                dir = 'SB'
            elif dir == 'EB':
                dir = 'NB'
            elif dir == 'NB':
                dir = 'EB'
            elif dir == 'SB':
                dir = 'WB'
        else:
            dir = error

        return dir

    #Check if the streets are "vertical" to see if it is NB/SB or EB/WB
    def verticalCheck(self):
        if self.SoMaCheck() == True:
            originLat = (self.lat_1 + self.lat_2) / 2
            originLng = (self.lng_1 + self.lng_2) / 2
            origin = (originLat, originLng)
            point1 = (self.lat_1, self.lng_1)
            point2 = (self.lat_2, self.lng_2)

            rotated1 = self.rotate(origin, point1, 0.785398)
            rotated2 = self.rotate(origin, point2, 0.785398)
            self.lat_1 = rotated1[0]
            self.lng_1 = rotated1[1]
            self.lat_2 = rotated2[0]
            self.lng_2 = rotated2[1]

        lat_difference = abs(self.lat_2 - self.lat_1)
        lng_difference = abs(self.lng_2 - self.lng_1)

        if lat_difference > lng_difference:
            self.isVertical = True
        else:
            self.isVertical = False

        return self.isVertical

    #helper function to rotate streets
    def rotate(self, origin, point, angle):
        """
        Rotate a point counterclockwise by a given angle around a given origin.

        The angle should be given in radians.
        """
        ox, oy = origin
        px, py = point

        qx = ox + math.cos(angle) * (px - ox) - math.sin(angle) * (py - oy)
        qy = oy + math.sin(angle) * (px - ox) + math.cos(angle) * (py - oy)
        return qx, qy


    def getMainline(self):
        return self.mainline


    def getStreet1(self):
        return self.street1


    def getStreet2(self):
        return self.street2



#General methods

#Formats date
#Parameter: Date from cell
def dateFormat(DATE):
    hasNameDate = False
    nameDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

    if DATE == '-':
        return False

    for days in nameDays:
        if days in DATE:
            hasNameDate = True

    if (isinstance(DATE, str) or isinstance(DATE, unicode)) and hasNameDate:
        #Splitting substring
        tempDate = DATE[DATE.find(',') + 2:]
        year = tempDate[tempDate.find(',') + 2:]
        month = tempDate[:tempDate.find(' ')]
        day = tempDate[tempDate.find(' ') + 1:tempDate.find(',')]
        month = MONTHS[month]

    elif isinstance(DATE, tuple) and not hasNameDate:
        year = DATE[0]
        month = DATE[1]
        day = DATE[2]

    elif (isinstance(DATE, str) or isinstance(DATE, unicode)) and not hasNameDate:
        year = DATE[DATE.find(',') + 1:].strip()
        DATE = DATE[:DATE.find(',')]
        month = DATE[:DATE.find(' ')].strip()
        day = DATE[DATE.find(' '):].strip()
        month = MONTHS[month]



    #Add extra zero in front if the values are less than 10

    if int(month) < 10:
        month = '0' + str(month)

    if int(day) < 10:
        day = '0' + str(day)
    formattedDate = str(year) + '.' + str(month) + '.' + str(day)


    return formattedDate

#Find all instances of a char in string
#Return: Array of ints
def findCharInString(s, ch):
    return [i for i, ltr in enumerate(s) if ltr == ch]

def dayValidation(data):
    sum = 0

    for num in data:
        sum = sum + num

    if sum > 0:
        return True
    else:
        return False





