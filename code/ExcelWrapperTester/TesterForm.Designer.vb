<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TesterForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnPopulateSpreadsheet = New System.Windows.Forms.Button()
        Me.txtExcelFilePath = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 20)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Excel File:"
        '
        'btnPopulateSpreadsheet
        '
        Me.btnPopulateSpreadsheet.Location = New System.Drawing.Point(41, 80)
        Me.btnPopulateSpreadsheet.Name = "btnPopulateSpreadsheet"
        Me.btnPopulateSpreadsheet.Size = New System.Drawing.Size(117, 84)
        Me.btnPopulateSpreadsheet.TabIndex = 1
        Me.btnPopulateSpreadsheet.Text = "Populate Spreadsheet"
        Me.btnPopulateSpreadsheet.UseVisualStyleBackColor = True
        '
        'txtExcelFilePath
        '
        Me.txtExcelFilePath.Location = New System.Drawing.Point(123, 31)
        Me.txtExcelFilePath.Name = "txtExcelFilePath"
        Me.txtExcelFilePath.Size = New System.Drawing.Size(614, 26)
        Me.txtExcelFilePath.TabIndex = 2
        Me.txtExcelFilePath.Text = "C:\\Users\\Ryan\\Desktop\\test.xlsx"
        '
        'TesterForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkOrange
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.txtExcelFilePath)
        Me.Controls.Add(Me.btnPopulateSpreadsheet)
        Me.Controls.Add(Me.Label1)
        Me.Name = "TesterForm"
        Me.Text = "Excel Wrapper Tester"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents btnPopulateSpreadsheet As Button
    Friend WithEvents txtExcelFilePath As TextBox
End Class
