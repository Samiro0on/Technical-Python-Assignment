# Samir       # mahmoudsamir109@gmail.com       # 23/02/2019        # github == @samiro0on
# another approach
import pandas
import re
from threading import *

class MyTool(Thread):

    def __init__(self,fileName):
        self.fileName = fileName

    def access(self):
        file1 = pandas.ExcelFile(self.fileName)
        sheet1 = pandas.read_excel(file1, index_col=[0], sheet_name='Sheet1')
        print(sheet1)  # if you wanna see the whole sheet
        # convert data frame to dictionary where headers are keys and values = data in a list
        myDic = sheet1.to_dict('list')
        myStr = list(myDic.values())
        # output now is list of list so i need to remove the outer one
        myNewStr = []
        for lists in myStr:
            for L1 in lists:
                myNewStr.append(L1)
        # print("o/p as one list ", myNewStr)

    # method to take sheet in excel in dataframe formate and convert it to list or string
    def decodedStr(self):
        sheet = pandas.read_excel(self.fileName, index_col=[0], sheet_name='Sheet1')
        row = sheet.shape[1]
        col = sheet.shape[0]
        midData = []
        for t in range(row):
            for n in range(col):
                # accessing data frame
                midData.append(sheet.iloc[n, t])
        # print("my str in a list : ", midData)
        # remove list and join the elements
        decodedLine = ' '.join(str(s) for s in midData)
        # print("input as one line : ", decodedLine)
        return decodedLine
    # now i have the data as one line with spaces

    def noSpaces(self, oneLine):
        strWithoutSpaces = ''.join(str(s) for s in oneLine)
        # print("input line without space : ", strWithoutSpaces)
        return strWithoutSpaces

    # now i have the data one line without spaces

class AllAlpaNum(Thread):

    def onlyAlphNum(self,str):
        #  import re and replace any char upper or lower or numeric with "" == nothing
        newStr = re.sub('[^0-9a-zA-Z]+', '', str)
        print("connected alphnumeric line: ", newStr)

        # strWithoutSpaces = re.sub('[^0-9a-zA-Z]+', '', strWithoutSpaces)
        # print("haaaa: ", strWithoutSpaces)

    # another way to solve the same problem
    def noSpecial(self, str):
        x = re.findall("[^!@#$%^&*~]", str)
        return "".join(x)
    # print(noSpecial(oneLine))

class NoDuplication(Thread):

    def removeDoubleSpecialChar(self, str):
        s1 = re.sub("!!", ' ', str)
        s2 = re.sub("##", ' ', s1)
        s3 = re.sub("@@", ' ', s2)
        s4 = re.sub("%%", ' ', s3)
        s5 = re.sub("&&", ' ', s4)
        s6 = re.sub("~~", ' ', s5)
        return s6

class ConsequentNum(Thread):

    def getTripleDigits(self, str):
        x = re.findall("[0-9][0-9][0-9]", str)
        return x

    # print(getTripleDigits(oneLine))

class DuplicateWords(Thread):

    def noDuplication(self, str):
        lastOP = []
        here = set()
        for words in str.split():
    # first if condition
            if words not in here:
                lastOP.append(words)
                here.add(words)
        return " ".join(lastOP)
    # output comparing words

    def noDouble(self, st1):
        op = []
        for n in st1:
            if n not in op:
                op.append(n)
        return " ".join(op)
    # output comparing char



if __name__ == "__main__":
# access certain sheet pass excelFileName.xlsx inside the class
    s1 = MyTool("book.xlsx")
    # if you dont wanna see the whole excel sheet comment the nextLine
    s1.access()
    # one line string but have white spaces
    oneLine = s1.decodedStr()
    # one line string but all connected
    oneLineWithoutSpaces = s1.noSpaces(oneLine)
    # the next 2 lines print one string and all alphanumeric char connected
    # comment the next 2 lines if y dont need to see that
    s2 = AllAlpaNum()
    s2.start()

    s2.onlyAlphNum(oneLineWithoutSpaces)
    # one line without special char
    print("one line without special char : ", s2.noSpecial(oneLine))

    s3 = NoDuplication()
    s3.start()
    print("line with white spaces due to duplicate special char : ", s3.removeDoubleSpecialChar(oneLine))
############################################################################################################
    s4 = ConsequentNum()
    s4.start()
    output4 = s4.getTripleDigits(oneLineWithoutSpaces)
    print("list of all occurrences of three consequent numeric characters : ", output4)
############################################################################################################
    testline = s3.removeDoubleSpecialChar(oneLine)
    s5 = DuplicateWords()
    s5.start()
    output5 = s5.noDuplication(testline)
    print("List of all occurrences of duplicated words : ", output5)
    output6 = s5.noDouble(testline)
    print("List of all occurrences of duplicated char : ", output6)
##########################################################################################################

# before every proper output required you will find a line of "############"
# go to command prompt and write down pip install pandas
# go to command prompt and write down pip install re
# go to command prompt and write down pip install xlrd
# if it is your first time to use those API's and first time to access excel sheet
# version of python and interpreter
# thank you