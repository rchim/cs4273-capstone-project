Public Class TesterForm
    Private Sub ButtonPopulateSpreadsheet_Click(sender As Object, e As EventArgs) Handles btnPopulateSpreadsheet.Click
        ' This code to set up the workbook (and much of the later code)
        ' is taken verbatim from ErrFunctions.vb for the HCI website,
        ' except we use ExcelWrapper instead of Excel, because our library
        ' ExcelWrapper is replacing Excel.
        Dim xlApp As New ExcelWrapper.Application

        ' In VB, no parens on a function/constructor call just passes no args
        Dim xlWbk As ExcelWrapper.Workbook = xlApp.Workbooks.Add
        Dim xlWks1 As ExcelWrapper.Worksheet = xlWbk.Worksheets(1)
        Dim xlWks2 As ExcelWrapper.Worksheet = xlWbk.Worksheets.Item(2)
        Dim xlWks3 As ExcelWrapper.Worksheet = xlWbk.Worksheets(3)

        Try
            ' We'll use the hourglass cursor to show that the Worksheet
            ' is being populated.
            Me.Cursor = Cursors.WaitCursor

            xlWks1.Name = "strCurrFkorg"
            xlWks2.Name = "Summary Page"
            Dim xlWks4 As ExcelWrapper.Worksheet = xlWbk.Worksheets.Add()
            xlWks4.Name = "The best sheet"
            xlWks4.Name = "No really its the best"

            xlWks3.Delete()

            xlWks2.Cells.NumberFormat = "@"
            xlWks2.Columns(4).NumberFormat = "0.00"
            xlWks2.Activate()
            xlWks2.Cells(5, 1).Value = "Field"
            xlWks2.Cells(5, 2).Value = "Description"
            xlWks2.Cells(5, 3).Value = "# Errors"
            xlWks2.Range("D5").Value = "Percent"
            xlWks2.Cells(5, 5).Value = "Total Hosp Records:"
            xlWks2.Cells(1, 5).Value = "Total Excel Rows:"

            ' Watch how the entry in column 4 gets a different number format,
            ' since we set it above.
            xlWks2.Cells(8, 3).Value = 1
            xlWks2.Cells(8, 4).Value = 2
            xlWks2.Cells(8, 5).Value = 3

            '   top row color
            xlWks2.Cells(1, 1).Interior.ColorIndex = 15
            xlWks2.Cells(1, 2).Interior.ColorIndex = 15
            xlWks2.Cells(1, 3).Interior.ColorIndex = 15
            xlWks2.Cells(1, 4).Interior.ColorIndex = 15
            xlWks2.Cells(1, 5).Interior.ColorIndex = 15
            '   fields header color
            xlWks2.Cells(2, 2).Interior.ColorIndex = 15
            xlWks2.Cells(3, 2).Interior.ColorIndex = 15
            xlWks2.Cells(5, 1).Interior.ColorIndex = 15
            xlWks2.Cells(5, 2).Interior.ColorIndex = 15
            xlWks2.Cells(5, 3).Interior.ColorIndex = 15
            xlWks2.Cells(5, 4).Interior.ColorIndex = 15
            xlWks2.Cells(5, 5).Interior.ColorIndex = 15

            xlWks2.Cells.ColumnWidth = 5
            xlWks2.Cells(4, 2).ColumnWidth = 25

            xlWks2.Columns.AutoFit()

            xlWks2.Cells(5, 5).Activate()

            xlApp.ActiveWorkbook.SaveAs(txtExcelFilePath.Text, fileFormat:=56)
            xlWks2.Rows(1).EntireRow.Delete() 'A change after the save...
            xlApp.ActiveWorkbook.SaveAs("C:\\Users\\Ryan\\Desktop\\test2.xlsx", fileFormat:=56)
            xlApp.ActiveWorkbook.Close()
        Catch ex As Exception
            MessageBox.Show("ERROR: " + ex.Message)
        Finally
            Me.Cursor = Cursors.Default ' Get rid of the hourglass cursor.
            MessageBox.Show("DONE")
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtExcelFilePath.TextChanged

    End Sub
End Class
