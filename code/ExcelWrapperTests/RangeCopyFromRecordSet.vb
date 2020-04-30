Imports NUnit.Framework
Imports Microsoft.Office.Interop
Imports ExcelWrapper

''' <summary>
''' Verifies that the Range.CopyFromRecordset method works the same
''' way in Excel and ExcelWrapper.
''' </summary>
Public Class RangeCopyFromRecordset
    Dim excelApp As Excel.Application
    Dim excelWbk As Excel.Workbook
    Dim excelWks As Excel.Worksheet
    Dim wrapperApp As Application
    Dim wrapperWbk As Workbook
    Dim wrapperWks As Worksheet

    <SetUp>
    Public Sub Setup()
        ' Set up excel worksheet
        excelApp = New Excel.Application
        excelWbk = excelApp.Workbooks.Add
        excelWks = excelWbk.Worksheets(1)

        ' Set up wrapper worksheet
        wrapperApp = New Application
        wrapperWbk = wrapperApp.Workbooks.Add
        wrapperWks = wrapperWbk.Worksheets(1)
    End Sub

    <TearDown>
    Public Sub TearDown()
        ' Trick Excel into thinking the workbook is saved so it 
        ' allows us to quit
        excelWbk.Saved = True
        excelWbk.Close(SaveChanges:=False)
        excelWks = Nothing
        excelWbk = Nothing
        excelApp.Quit()
        excelApp = Nothing

        ' Dispose of the wrapper resources too
        wrapperApp.Quit()
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper copy a record set with
    ''' three fields and four records the same way.
    ''' </summary>
    <Test>
    Public Sub ThreeFieldsFourRecords()
        ' First we construct the record set we will use.
        Dim rs As ADODB.Recordset = New ADODB.Recordset()
        rs.Fields.Append("VarCharField", ADODB.DataTypeEnum.adVarChar, 255)
        rs.Fields.Append("IntField", ADODB.DataTypeEnum.adInteger, 5)
        rs.Fields.Append("DateField", ADODB.DataTypeEnum.adDBDate)
        rs.Open(CursorType:=ADODB.CursorTypeEnum.adOpenForwardOnly)

        ' Add the test data
        Dim fieldsArray As Object() = {"VarCharField", "IntField", "DateField"}
        Dim values1 As Object() = {"test", 2, #08/21/1998#}
        Dim values2 As Object() = {"hello", 3, #03/09/2019#}
        Dim values3 As Object() = {"world", 5, #07/20/2020#}
        Dim values4 As Object() = {"C sharp", 7, #05/10/2004#}
        rs.AddNew(fieldsArray, values1)
        rs.AddNew(fieldsArray, values2)
        rs.AddNew(fieldsArray, values3)
        rs.AddNew(fieldsArray, values4)
        rs.MoveFirst()

        ' Copy from record set in both Excel and the wrapper
        Dim excelStartCell As Excel.Range = excelWks.Cells(2, 3)
        Dim wrapperStartCell As Range = wrapperWks.Cells(2, 3)
        excelStartCell.CopyFromRecordset(rs)
        wrapperStartCell.CopyFromRecordset(rs)

        ' Make sure the outcome is the same on the range A1:H8.
        For row As Integer = 1 To 8
            For col As Integer = 1 To 8
                Dim excelVal As Object = excelWks.Cells(row, col).Value
                Dim wrapperVal As Object = wrapperWks.Cells(row, col).Value
                Assert.AreEqual(excelVal, wrapperVal)
            Next
        Next


    End Sub

End Class
