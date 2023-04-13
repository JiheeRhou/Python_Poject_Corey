import os, openpyxl

# Excel columns
columnList = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
# boolean for whether the first file is loaded 
isLoadedFirstFile = False
# boolean for whether the second file is loaded 
isLoadedFile = False

class CompareFile:
    def __init__(self):
        global isLoadedFile

        self.firstMaxColumn = 0
        self.firstMaxRow = 1
        self.secondMaxColumn = 0
        self.secondMaxRow = 1
        self.indexOfFirstSheetId = 0
        self.indexOfSecondSheetId = 0
        self.ArrSecondSheetInstructorIds = []
        self.ArrNewInstructorIds = []

        file = "C:\\example\\file1.xlsx"
        if self.existsFile(file):
            self.workbook = openpyxl.load_workbook(file, data_only=True)
            self.firstSheet = self.workbook.worksheets[0]
            self.secondSheet = self.workbook.worksheets[1]
            isLoadedFile = True

    def existsFile(self, file):
        if not os.path.exists(file):
            print("ERROR!!")
            print("The file ( ", file, " ) does not exist.")
            return False
        else:
            return True

    def getFirstSheetInfo(self):
        while not self.firstSheet[columnList[self.firstMaxColumn]+'1'].value == None:
            
            if self.firstSheet[columnList[self.firstMaxColumn]+'1'].value == "Instructor ID":
                self.indexOfFirstSheetId = self.firstMaxColumn

            self.firstMaxColumn = self.firstMaxColumn + 1

        while not self.firstSheet[columnList[self.indexOfFirstSheetId] + str(self.firstMaxRow)].value == None:
            
            self.firstMaxRow = self.firstMaxRow + 1
        
    def getSecondSheetInfo(self):
        while not self.secondSheet[columnList[self.secondMaxColumn]+'1'].value == None:
            
            if self.secondSheet[columnList[self.secondMaxColumn]+'1'].value == "Instructor ID":
                self.indexOfSecondSheetId = self.secondMaxColumn

            self.secondMaxColumn = self.secondMaxColumn + 1

        while not self.secondSheet[columnList[self.indexOfSecondSheetId] + str(self.secondMaxRow)].value == None:
            id = self.secondSheet[columnList[self.indexOfSecondSheetId] + str(self.secondMaxRow)].value
            if not id == "Instructor ID":
                self.ArrSecondSheetInstructorIds.append(id)
                
            self.secondMaxRow = self.secondMaxRow + 1
            
        print(self.ArrSecondSheetInstructorIds)

    def getNewInstructorIds(self):
        for i in range (1, self.firstMaxRow):
            id = self.firstSheet[columnList[self.indexOfFirstSheetId]+str(i)].value
            if id not in self.ArrSecondSheetInstructorIds:
                self.ArrNewInstructorIds.append(id)
                
        print(self.ArrNewInstructorIds)
    
    def createThirdFile(self):
        thirdSheet = self.workbook.create_sheet("third sheet")
        indexOfFirstSheetFN = 0
        indexOfFirstSheetLN = 0
        indexOfFirstSheetEmail = 0
        id = ""
        firstName = ""
        lastName = ""
        email = ""
        
        for i in range (self.firstMaxColumn):
            if self.firstSheet[columnList[i]+'1'].value == "Instructor First Name":
                indexOfFirstSheetFN = i
            elif self.firstSheet[columnList[i]+'1'].value == "Instructor Last Name":
                indexOfFirstSheetLN = i
            elif self.firstSheet[columnList[i]+'1'].value == "Instructor Email":
                indexOfFirstSheetEmail = i
                
        for i in range (1, self.firstMaxRow):
            id = self.firstSheet[columnList[self.indexOfFirstSheetId]+str(i)].value

            if id not in self.ArrSecondSheetInstructorIds:
                firstName = self.firstSheet[columnList[indexOfFirstSheetFN]+str(i)].value
                lastName = self.firstSheet[columnList[indexOfFirstSheetLN]+str(i)].value
                email = self.firstSheet[columnList[indexOfFirstSheetEmail]+str(i)].value
                thirdSheet.append([id, firstName, lastName, email])
                
        self.workbook.save("C:\\example\\file2.xlsx")

compareFile = CompareFile()

if isLoadedFile:
    compareFile.getFirstSheetInfo()
    compareFile.getSecondSheetInfo()
    compareFile.getNewInstructorIds()
    compareFile.createThirdFile()