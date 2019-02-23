# Samir       # mahmoudsamir109@gmail.com       # 22/02/2019        # github == @samiro0on
# input Excel sheet .xlsx of n rows and m columns
import pandas
import re

try:
    file1 = pandas.ExcelFile("book.xlsx")
    sheet1 = pandas.read_excel(file1, index_col=[0], sheet_name='Sheet1')

    print(sheet1)     # if you wanna see the whole sheet
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
    def decodedStr(sheet):
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

    oneLine = decodedStr(sheet1)
    # now i have the data as one line with spaces

    def noSpaces(oneLine):
        strWithoutSpaces = ''.join(str(s) for s in oneLine)
        # print("input line without space : ", strWithoutSpaces)
        return strWithoutSpaces

    oneLineWithoutSpaces = noSpaces(myNewStr)
        # now i have the data one line without spaces
    ############################################################################

    def onlyAlphNum(str):
        #  import re and replace any char upper or lower or numeric with "" == nothing
        newStr = re.sub('[^0-9a-zA-Z]+', '', str)
        print("connected alphnumeric line: ", newStr)
    onlyAlphNum(oneLine)
        # strWithoutSpaces = re.sub('[^0-9a-zA-Z]+', '', strWithoutSpaces)
        # print("haaaa: ", strWithoutSpaces)

    # another way to solve the same problem
    def noSpecial(str):
        x = re.findall("[^!@#$%^&*~]", str)
        return "".join(x)
    # print(noSpecial(oneLine))

#############################################################################
# replace the multiple occurrences of non-alphanumeric to space

    def removeDoubleSpecialChar(str):
        s1 = re.sub("!!", ' ', str)
        s2 = re.sub("##", ' ', s1)
        s3 = re.sub("@@", ' ', s2)
        s4 = re.sub("%%", ' ', s3)
        s5 = re.sub("&&", ' ', s4)
        s6 = re.sub("~~", ' ', s5)
        return s6
    testline = removeDoubleSpecialChar(oneLine)
    # print(removeDoubleSpecialChar(oneLine))
    # print(removeDoubleSpecialChar(oneLineWithoutSpaces))

##########################################################################3
# if three consequent characters are numeric add them to a list

    def getTripleDigits(str):
        x = re.findall("[0-9][0-9][0-9]", str)
        return x

    # print(getTripleDigits(oneLine))
    print("list of all occurrences of three consequent numeric characters : ", getTripleDigits(oneLineWithoutSpaces))

###############################################################3

# no duplication words

    def noDuplication(str):
        lastOP = []
        here = set()
        for words in str.split():
    # first if condition
            if words not in here:
                lastOP.append(words)
                here.add(words)
        return " ".join(lastOP)
    # output comparing words
    print("List of all occurrences of duplicated words : ", noDuplication(testline))

    def noDouble(st1):
        op = []
        for n in st1:
            if n not in op:
                op.append(n)
        return " ".join(op)
    # output comparing char
    print("List of all occurrences of duplicated char : ", noDouble(testline))

# method parse # method head
except IOError:
    print("cant read the file ")


except Exception as error:
    print("error is in : ", error.args[0])

finally:
    print("thank you")

# i used 1 if condition
# threading and multiprocessing
# exceptional handling
# documentation
# naming
# still my logic have weakness points
# still need to know more about threading in methods and classes
