Imports NUnit.Framework
Imports Microsoft.Office.Interop
Imports ExcelWrapper

''' <summary>
''' Verifies that getting and setting cell values works the same way
''' in ExcelWrapper and Excel. In particular, we need "Is Nothing" checks 
''' and "<> Nothing" checks to work the same when applied to cell values 
''' retrieved from the two libraries.
''' </summary>
Public Class CellValues
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
    ''' Verifies that getting an empty (never-touched) cell value gives
    ''' the same result in Excel and ExcelWrapper.
    ''' </summary>
    <Test>
    Public Sub GetEmpty()
        Dim excelVal As Object = excelWks.Cells(3, 2).Value
        Dim wrapperVal As Object = wrapperWks.Cells(3, 2).Value

        ' Make sure Excel's empty cell val and the wrapper val are equal
        ' in a loose sense
        Assert.AreEqual(excelVal, wrapperVal,
            $"Excel val was {excelVal} but wrapper val was {wrapperVal}")

        ' Make sure Excel's val and the wrapper val have the same 
        ' type (AreEqual already checks for this if one of them Is Nothing)
        If ((excelVal IsNot Nothing) And (wrapperVal IsNot Nothing)) Then
            Dim excelValType As Type = excelVal.GetType()
            Dim wrapperValType As Type = wrapperVal.GetType()

            Assert.AreEqual(excelValType.ToString, wrapperValType.ToString)
        End If
    End Sub

    ''' <summary>
    ''' Verifies that setting and getting a cell value gives the same
    ''' result in Excel and ExcelWrapper.
    ''' </summary>
    <TestCase(0)>
    <TestCase(0.0)>
    <TestCase(3)>
    <TestCase(-2.0)>
    <TestCase("my string")>
    <TestCase("")>
    <TestCase("0")>
    <TestCase("3")>
    Public Sub SetAndGetNumbersAndStrings(value)
        excelWks.Cells(3, 2).Value = value
        wrapperWks.Cells(3, 2).Value = value

        Dim excelVal As Object = excelWks.Cells(3, 2).Value
        Dim wrapperVal As Object = wrapperWks.Cells(3, 2).Value

        ' Make sure Excel's val and the wrapper val are equal
        ' in a loose sense
        Assert.AreEqual(excelVal, wrapperVal,
            $"Excel val was {excelVal} but wrapper val was {wrapperVal}")

        ' Make sure Excel's val and the wrapper val have the same 
        ' type (AreEqual already checks for this if one of them Is Nothing)
        If ((excelVal IsNot Nothing) And (wrapperVal IsNot Nothing)) Then
            Dim excelValType As Type = excelVal.GetType()
            Dim wrapperValType As Type = wrapperVal.GetType()

            Assert.AreEqual(excelValType.ToString, wrapperValType.ToString)
        End If
    End Sub

    <Test>
    Public Sub SetAndGetDate()
        excelWks.Cells(3, 2).Value = #03/09/2019#
        wrapperWks.Cells(3, 2).Value = #03/09/2019#

        Dim excelVal As Object = excelWks.Cells(3, 2).Value
        Dim wrapperVal As Object = wrapperWks.Cells(3, 2).Value

        ' Make sure Excel's val and the wrapper val are equal
        ' in a loose sense
        Assert.AreEqual(excelVal, wrapperVal,
            $"Excel val was {excelVal} but wrapper val was {wrapperVal}")

        ' Make sure Excel's val and the wrapper val have the same 
        ' type (AreEqual already checks for this if one of them Is Nothing)
        If ((excelVal IsNot Nothing) And (wrapperVal IsNot Nothing)) Then
            Dim excelValType As Type = excelVal.GetType()
            Dim wrapperValType As Type = wrapperVal.GetType()

            Assert.AreEqual(excelValType.ToString, wrapperValType.ToString)
        End If
    End Sub

End Class
