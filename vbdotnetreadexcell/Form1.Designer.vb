<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.LabelCellValue = New System.Windows.Forms.Label()
        Me.ButtonSelectFile = New System.Windows.Forms.Button()
        Me.ComboBoxSheets = New System.Windows.Forms.ComboBox()
        Me.ButtonGetCellValue = New System.Windows.Forms.Button()
        Me.TextBoxCellReference = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'LabelCellValue
        '
        Me.LabelCellValue.AutoSize = True
        Me.LabelCellValue.Location = New System.Drawing.Point(13, 33)
        Me.LabelCellValue.Name = "LabelCellValue"
        Me.LabelCellValue.Size = New System.Drawing.Size(77, 13)
        Me.LabelCellValue.TabIndex = 0
        Me.LabelCellValue.Text = "LabelCellValue"
        '
        'ButtonSelectFile
        '
        Me.ButtonSelectFile.Location = New System.Drawing.Point(16, 63)
        Me.ButtonSelectFile.Name = "ButtonSelectFile"
        Me.ButtonSelectFile.Size = New System.Drawing.Size(260, 23)
        Me.ButtonSelectFile.TabIndex = 1
        Me.ButtonSelectFile.Text = "ButtonSelectFile"
        Me.ButtonSelectFile.UseVisualStyleBackColor = True
        '
        'ComboBoxSheets
        '
        Me.ComboBoxSheets.FormattingEnabled = True
        Me.ComboBoxSheets.Location = New System.Drawing.Point(16, 108)
        Me.ComboBoxSheets.Name = "ComboBoxSheets"
        Me.ComboBoxSheets.Size = New System.Drawing.Size(121, 21)
        Me.ComboBoxSheets.TabIndex = 2
        '
        'ButtonGetCellValue
        '
        Me.ButtonGetCellValue.Location = New System.Drawing.Point(154, 106)
        Me.ButtonGetCellValue.Name = "ButtonGetCellValue"
        Me.ButtonGetCellValue.Size = New System.Drawing.Size(180, 23)
        Me.ButtonGetCellValue.TabIndex = 4
        Me.ButtonGetCellValue.Text = "ButtonGetCellValue"
        Me.ButtonGetCellValue.UseVisualStyleBackColor = True
        '
        'TextBoxCellReference
        '
        Me.TextBoxCellReference.Location = New System.Drawing.Point(12, 188)
        Me.TextBoxCellReference.Name = "TextBoxCellReference"
        Me.TextBoxCellReference.Size = New System.Drawing.Size(236, 20)
        Me.TextBoxCellReference.TabIndex = 5
        Me.TextBoxCellReference.Text = "A2"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.TextBoxCellReference)
        Me.Controls.Add(Me.ButtonGetCellValue)
        Me.Controls.Add(Me.ComboBoxSheets)
        Me.Controls.Add(Me.ButtonSelectFile)
        Me.Controls.Add(Me.LabelCellValue)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelCellValue As Label
    Friend WithEvents ButtonSelectFile As Button
    Friend WithEvents ComboBoxSheets As ComboBox
    Friend WithEvents ButtonGetCellValue As Button
    Friend WithEvents TextBoxCellReference As TextBox
End Class
