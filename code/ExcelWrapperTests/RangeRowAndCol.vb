Imports NUnit.Framework
Imports Microsoft.Office.Interop
Imports ExcelWrapper

''' <summary>
''' Verifies that the Range.Row and Range.Column properties work the same way
''' in ExcelWrapper and Excel.
''' </summary>
Public Class RangeRowAndCol
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
    ''' Verifies that Excel and ExcelWrapper give the same row for
    ''' a range representing a single cell.
    ''' </summary>
    <Test>
    Public Sub SingleCellRow()
        Dim excelRange As Excel.Range = excelWks.Range("K14")
        Dim wrapperRange As Range = wrapperWks.Range("K14")
        Dim excelRow As Integer = excelRange.Row
        Dim wrapperRow As Integer = wrapperRange.Row

        Assert.AreEqual(excelRow, wrapperRow)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same column for
    ''' a range representing a single cell.
    ''' </summary>
    <Test>
    Public Sub SingleCellColumn()
        Dim excelRange As Excel.Range = excelWks.Range("K14")
        Dim wrapperRange As Range = wrapperWks.Range("K14")
        Dim excelCol As Integer = excelRange.Column
        Dim wrapperCol As Integer = wrapperRange.Column

        Assert.AreEqual(excelCol, wrapperCol)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same row for
    ''' a range spanning multiple but (finitely many) rows, specified 
    ''' with the top row first.
    ''' </summary>
    <Test>
    Public Sub TopToBottomRangeRow()
        Dim excelRange As Excel.Range = excelWks.Range("K14:Y20")
        Dim wrapperRange As Range = wrapperWks.Range("K14:Y20")
        Dim excelRow As Integer = excelRange.Row
        Dim wrapperRow As Integer = wrapperRange.Row

        Assert.AreEqual(excelRow, wrapperRow)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same row for
    ''' a range spanning multiple but (finitely many) rows, specified 
    ''' with the bottom row first.
    ''' </summary>
    <Test>
    Public Sub BottomToTopRangeRow()
        Dim excelRange As Excel.Range = excelWks.Range("Y20:K14")
        Dim wrapperRange As Range = wrapperWks.Range("Y20:K14")
        Dim excelRow As Integer = excelRange.Row
        Dim wrapperRow As Integer = wrapperRange.Row

        Assert.AreEqual(excelRow, wrapperRow)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same column for
    ''' a range spanning multiple but (finitely many) columns, specified 
    ''' with the leftmost column first.
    ''' </summary>
    <Test>
    Public Sub LeftToRightRangeColumn()
        Dim excelRange As Excel.Range = excelWks.Range("K14:Y20")
        Dim wrapperRange As Range = wrapperWks.Range("K14:Y20")
        Dim excelCol As Integer = excelRange.Column
        Dim wrapperCol As Integer = wrapperRange.Column

        Assert.AreEqual(excelCol, wrapperCol)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same column for
    ''' a range spanning multiple but (finitely many) columns, specified 
    ''' with the leftmost column first.
    ''' </summary>
    <Test>
    Public Sub RightToLeftRangeColumn()
        Dim excelRange As Excel.Range = excelWks.Range("Y20:K14")
        Dim wrapperRange As Range = wrapperWks.Range("Y20:K14")
        Dim excelCol As Integer = excelRange.Column
        Dim wrapperCol As Integer = wrapperRange.Column

        Assert.AreEqual(excelCol, wrapperCol)
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper give the same row for
    ''' a range made up of all the cells in a worksheet.
    ''' </summary>
    <Test>
    Public Sub AllCellsRow()
        Dim excelRange As Excel.Range = excelWks.Cells
        Dim wrapperRange As Range = wrapperWks.Cells
        Dim excelRow As Integer = excelRange.Row
        Dim wrapperRow As Integer = wrapperRange.Row

        Assert.AreEqual(excelRow, wrapperRow)
    End Sub

    ''' Verifies that Excel and ExcelWrapper give the same column for
    ''' a range made up of all the cells in a worksheet.
    ''' </summary>
    <Test>
    Public Sub AllCellsColumn()
        Dim excelRange As Excel.Range = excelWks.Cells
        Dim wrapperRange As Range = wrapperWks.Cells
        Dim excelCol As Integer = excelRange.Column
        Dim wrapperCol As Integer = wrapperRange.Column

        Assert.AreEqual(excelCol, wrapperCol)
    End Sub
End Class
