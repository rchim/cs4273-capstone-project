Imports NUnit.Framework
Imports Microsoft.Office.Interop
Imports ExcelWrapper

''' <summary>
''' Verifies that the Range.End method works the same way in Excel and
''' ExcelWrapper.
''' </summary>
Public Class RangeSort
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
    ''' Verifies that Excel and ExcelWrapper sort a range the same way, when sorting by
    ''' columns in ascending order. (This is the use case that appears in ErrFunctions.vb)
    ''' </summary>
    <Test>
    Public Sub AscendingByColumns()
        ' populate the range B3:E5 in each worksheet with a block of nonempty cells.
        Dim vals(2, 3) As Object

        vals(0, 0) = 5
        vals(0, 1) = 3
        vals(0, 2) = 10
        vals(0, 3) = "test"

        vals(1, 0) = 6
        vals(1, 1) = 1
        vals(1, 2) = 9
        vals(1, 3) = "hi"

        vals(2, 0) = 7
        vals(2, 1) = 2
        vals(2, 2) = 8
        vals(2, 3) = 14

        For row As Integer = 3 To 5
            For col As Integer = 2 To 5
                excelWks.Cells(row, col).Value = vals(row - 3, col - 2)
                wrapperWks.Cells(row, col).Value = vals(row - 3, col - 2)
            Next
        Next

        ' Sort on both Excel and the wrapper by their column C value.
        excelWks.Range("B3:E5").Sort(
            Key1:=excelWks.Range("C100"),
            Order1:=Excel.XlSortOrder.xlAscending,
            Orientation:=Excel.XlSortOrientation.xlSortColumns
        )
        wrapperWks.Range("B3:E5").Sort(
            key:=wrapperWks.Range("C100"),
            order:=ExcelWrapper.XlSortOrder.xlAscending,
            orientation:=ExcelWrapper.XlSortOrientation.xlSortColumns
        )

        ' Make sure the outcome is the same on the range A1:H8.
        For row As Integer = 1 To 8
            For col As Integer = 1 To 8
                Dim excelVal As Object = excelWks.Cells(row, col).Value
                Dim wrapperVal As Object = wrapperWks.Cells(row, col).Value
                Assert.AreEqual(excelVal, wrapperVal)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper sort a range the same way, when sorting by
    ''' columns in descending order.
    ''' </summary>
    <Test>
    Public Sub DescendingByColumns()
        ' populate the range B3:E5 in each worksheet with a block of nonempty cells.
        Dim vals(2, 3) As Object

        vals(0, 0) = 5
        vals(0, 1) = 3
        vals(0, 2) = 10
        vals(0, 3) = "test"

        vals(1, 0) = 6
        vals(1, 1) = 1
        vals(1, 2) = 9
        vals(1, 3) = "hi"

        vals(2, 0) = 7
        vals(2, 1) = 2
        vals(2, 2) = 8
        vals(2, 3) = 14

        For row As Integer = 3 To 5
            For col As Integer = 2 To 5
                excelWks.Cells(row, col).Value = vals(row - 3, col - 2)
                wrapperWks.Cells(row, col).Value = vals(row - 3, col - 2)
            Next
        Next

        ' Sort on both Excel and the wrapper by their column C value.
        excelWks.Range("B3:E5").Sort(
            Key1:=excelWks.Range("C100"),
            Order1:=Excel.XlSortOrder.xlDescending,
            Orientation:=Excel.XlSortOrientation.xlSortColumns
        )
        wrapperWks.Range("B3:E5").Sort(
            key:=wrapperWks.Range("C100"),
            order:=ExcelWrapper.XlSortOrder.xlDescending,
            orientation:=ExcelWrapper.XlSortOrientation.xlSortColumns
        )

        ' Make sure the outcome is the same on the range A1:H8.
        For row As Integer = 1 To 8
            For col As Integer = 1 To 8
                Dim excelVal As Object = excelWks.Cells(row, col).Value
                Dim wrapperVal As Object = wrapperWks.Cells(row, col).Value
                Assert.AreEqual(excelVal, wrapperVal)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper sort a range the same way, when sorting by
    ''' rows in ascending order.
    ''' </summary>
    <Test>
    Public Sub AscendingByRows()
        ' populate the range B3:E5 in each worksheet with a block of nonempty cells.
        Dim vals(2, 3) As Object

        vals(0, 0) = 3
        vals(0, 1) = 5
        vals(0, 2) = 10
        vals(0, 3) = "test"

        vals(1, 0) = 6
        vals(1, 1) = 1
        vals(1, 2) = 9
        vals(1, 3) = "hi"

        vals(2, 0) = 2
        vals(2, 1) = 7
        vals(2, 2) = 8
        vals(2, 3) = 14

        For row As Integer = 3 To 5
            For col As Integer = 2 To 5
                excelWks.Cells(row, col).Value = vals(row - 3, col - 2)
                wrapperWks.Cells(row, col).Value = vals(row - 3, col - 2)
            Next
        Next

        ' Sort on both Excel and the wrapper by their row 4 value.
        excelWks.Range("B3:E5").Sort(
            Key1:=excelWks.Range("ABC4"),
            Order1:=Excel.XlSortOrder.xlAscending
        )
        wrapperWks.Range("B3:E5").Sort(
            key:=wrapperWks.Range("ABC4"),
            order:=ExcelWrapper.XlSortOrder.xlAscending
        )

        ' Make sure the outcome is the same on the range A1:H8.
        For row As Integer = 1 To 8
            For col As Integer = 1 To 8
                Dim excelVal As Object = excelWks.Cells(row, col).Value
                Dim wrapperVal As Object = wrapperWks.Cells(row, col).Value
                Assert.AreEqual(excelVal, wrapperVal)
            Next
        Next
    End Sub

    ''' <summary>
    ''' Verifies that Excel and ExcelWrapper sort a range the same way, when sorting by
    ''' rows in ascending order.
    ''' </summary>
    <Test>
    Public Sub DescendingByRows()
        ' populate the range B3:E5 in each worksheet with a block of nonempty cells.
        Dim vals(2, 3) As Object

        vals(0, 0) = 3
        vals(0, 1) = 5
        vals(0, 2) = 10
        vals(0, 3) = "test"

        vals(1, 0) = 6
        vals(1, 1) = 1
        vals(1, 2) = 9
        vals(1, 3) = "hi"

        vals(2, 0) = 2
        vals(2, 1) = 7
        vals(2, 2) = 8
        vals(2, 3) = 14

        For row As Integer = 3 To 5
            For col As Integer = 2 To 5
                excelWks.Cells(row, col).Value = vals(row - 3, col - 2)
                wrapperWks.Cells(row, col).Value = vals(row - 3, col - 2)
            Next
        Next

        ' Sort on both Excel and the wrapper by their row 4 value.
        excelWks.Range("B3:E5").Sort(
            Key1:=excelWks.Range("ABC4"),
            Order1:=Excel.XlSortOrder.xlDescending
        )
        wrapperWks.Range("B3:E5").Sort(
            key:=wrapperWks.Range("ABC4"),
            order:=ExcelWrapper.XlSortOrder.xlDescending
        )

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
