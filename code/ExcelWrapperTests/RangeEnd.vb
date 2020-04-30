Imports NUnit.Framework
Imports Microsoft.Office.Interop
Imports ExcelWrapper

''' <summary>
''' Verifies that the Range.End method works the same way in Excel and
''' ExcelWrapper.
''' </summary>
Public Class RangeEnd
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
        ' Trick Excel into thinking the workbook is saved so it allows us to quit
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
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several nonempty cells, and starting on
    ''' a nonempty cell.
    ''' </summary>
    <Test>
    Public Sub NonemptyCellsStartingWithNonempty()
        ' populate a column in each worksheet with a series of nonempty cells
        excelWks.Cells(3, 2).Value = 2
        wrapperWks.Cells(3, 2).Value = 2
        excelWks.Cells(4, 2).Value = "Test"
        wrapperWks.Cells(4, 2).Value = "Test"
        excelWks.Cells(5, 2).Value = 0.0
        wrapperWks.Cells(5, 2).Value = 0.0
        excelWks.Cells(6, 2).Value = "Another test"
        wrapperWks.Cells(6, 2).Value = "Another test"

        ' start at the top nonempty cell
        Dim excelStart As Excel.Range = excelWks.Cells(3, 2)
        Dim wrapperStart As Range = wrapperWks.Cells(3, 2)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlDown)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlDown)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several nonempty cells, and starting on
    ''' an empty cell.
    ''' </summary>
    <Test>
    Public Sub NonemptyCellsStartingWithEmpty()
        ' populate a column in each worksheet with a series of nonempty cells
        excelWks.Cells(3, 2).Value = 2
        wrapperWks.Cells(3, 2).Value = 2
        excelWks.Cells(4, 2).Value = "Test"
        wrapperWks.Cells(4, 2).Value = "Test"
        excelWks.Cells(5, 2).Value = 0.0
        wrapperWks.Cells(5, 2).Value = 0.0
        excelWks.Cells(6, 2).Value = "Another test"
        wrapperWks.Cells(6, 2).Value = "Another test"

        ' start on the empty cell below the nonempty cells
        Dim excelStart As Excel.Range = excelWks.Cells(7, 2)
        Dim wrapperStart As Range = wrapperWks.Cells(7, 2)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlUp)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlUp)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several empty cells, and starting on
    ''' a nonempty cell.
    ''' </summary>
    <Test>
    Public Sub EmptyCellsStartingWithNonempty()
        excelWks.Cells(7, 10).Value = 2
        wrapperWks.Cells(7, 10).Value = 2
        excelWks.Cells(7, 5).Value = "Test"
        wrapperWks.Cells(7, 5).Value = "Test"

        ' start on the nonempty cell to the right
        Dim excelStart As Excel.Range = excelWks.Cells(7, 10)
        Dim wrapperStart As Range = wrapperWks.Cells(7, 10)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlToLeft)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlToLeft)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several empty cells, and starting on
    ''' an empty cell.
    ''' </summary>
    <Test>
    Public Sub EmptyCellsStartingWithEmpty()
        excelWks.Cells(7, 10).Value = 2
        wrapperWks.Cells(7, 10).Value = 2

        ' start on an empty cell to the left
        Dim excelStart As Excel.Range = excelWks.Cells(7, 1)
        Dim wrapperStart As Range = wrapperWks.Cells(7, 1)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlToRight)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlToRight)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several empty cells, going all the way to
    ''' the edge of the sheet.
    ''' </summary>
    <Test>
    Public Sub EmptyCellsToEdge()
        Dim excelStart As Excel.Range = excelWks.Cells(14, 6)
        Dim wrapperStart As Range = wrapperWks.Cells(14, 6)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlToLeft)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlToLeft)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' dealing with a series of several nonempty cells, going all the way to
    ''' the edge of the sheet.
    ''' </summary>
    <Test>
    Public Sub NonemptyCellsToEdge()
        excelWks.Cells(1, 8).Value = 2
        wrapperWks.Cells(1, 8).Value = 2
        excelWks.Cells(2, 8).Value = "test"
        wrapperWks.Cells(2, 8).Value = "test"
        excelWks.Cells(3, 8).Value = 1.73
        wrapperWks.Cells(3, 8).Value = 1.73

        Dim excelStart As Excel.Range = excelWks.Cells(3, 8)
        Dim wrapperStart As Range = wrapperWks.Cells(3, 8)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlUp)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlUp)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' starting at a row 1 cell and going up.
    ''' </summary>
    <Test>
    Public Sub AlreadyAtTop()
        excelWks.Cells(1, 18).Value = "test"

        Dim excelStart As Excel.Range = excelWks.Cells(1, 18)
        Dim wrapperStart As Range = wrapperWks.Cells(1, 18)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlUp)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlUp)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper find the same range end when
    ''' starting at a column 1 cell and going left.
    ''' </summary>
    <Test>
    Public Sub AlreadyAtLeft()
        Dim excelStart As Excel.Range = excelWks.Cells(10, 1)
        Dim wrapperStart As Range = wrapperWks.Cells(10, 1)

        Dim excelEnd As Excel.Range = excelStart.End(Excel.XlDirection.xlToLeft)
        Dim wrapperEnd As Range = wrapperStart.End(XlDirection.xlToLeft)

        Dim excelEndAddress As String =
            excelEnd.Address(ReferenceStyle:=Excel.XlReferenceStyle.xlR1C1)
        Dim wrapperEndAddress As String = wrapperEnd.Address()

        Assert.AreEqual(excelEndAddress, wrapperEndAddress)
    End Sub
End Class
