from xlrd import open_workbook
import os,re
import xlsxwriter

filename = "7_1 copy.xlsx"

if (os.path.isfile(filename)):
        print(filename)
        wb = open_workbook(filename)
        SheetName = wb.sheet_names()
        InputRow = []
        MetaData_List = []
        for s in wb.sheets():
            for row in range(s.nrows):
                col_value = []
                for col in range(s.ncols):
                    value = (s.cell(row, col).value)
                    try:
                        value = str(int(value))
                    except:
                        pass
                    col_value.append(value)
                InputRow.append(col_value)
            MetaData_List.append(InputRow)
            InputRow = []


PlaceHolderList = []

#FinalList.append(MetaData_List[0][0])
for index in range(0,len(MetaData_List[1][0])):
    PlaceHolderList.append(MetaData_List[1][0][index])
#print(PlaceHolderList)
SerialNumber = 0
SerialNumberList = []

for SheetIndex in range(1,len(SheetName)):
    print(SheetName[SheetIndex])
    FinalList = []
    FinalList.append(MetaData_List[0][0])
    for row_index in range(1,len(MetaData_List[SheetIndex])):
        TempList = []
        VarWorksheet = []
        for row in range(1,len(MetaData_List[0])):
            temp = []
            for column in range(0,len(MetaData_List[0][row])):
                temp.append(MetaData_List[0][row][column])
            VarWorksheet.append(temp)
            SerialNumber = SerialNumber + 1
            SerialNumberList.append(SerialNumber)
        #print(VarWorksheet)
        for col_index in range(1,len(MetaData_List[SheetIndex][row_index])):
            PlaceHolderValue = MetaData_List[SheetIndex][row_index][col_index]
            #print(PlaceHolderValue)
            PlaceHolder      = PlaceHolderList[col_index]
            #print(PlaceHolder)



            for TestCaseSerialNumber in range(0,len(VarWorksheet)):
                for Column in range(1,len(VarWorksheet[TestCaseSerialNumber])):
                    VarWorksheet[TestCaseSerialNumber][Column] = re.sub(PlaceHolder,PlaceHolderValue,VarWorksheet[TestCaseSerialNumber][Column])

                #print(SerialNumber)
                #VarWorksheet[TestCaseSerialNumber][0] = SerialNumber

        #print(len(VarWorksheet))
        for data in VarWorksheet:
            FinalList.append(data)
    print(len(FinalList))
    for rows in range(1,len(FinalList)):
        FinalList[rows][0] = SerialNumberList[rows - 1]


    #print(MetaData_List[0][1])


    FileName = 'TC_Result_UC11_Reg_1_4.xlsx'
    workbook = xlsxwriter.Workbook(FileName)
    worksheet  = workbook.add_worksheet("TestCases")
    for i, l in enumerate(FinalList):
        for j, col in enumerate(l) :
            worksheet.write(i, j, col)
    workbook.close()

