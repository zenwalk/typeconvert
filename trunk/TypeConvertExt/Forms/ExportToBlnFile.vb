Option Strict Off
Option Explicit On 
Imports System.Runtime.InteropServices
<ComVisible(False)> Public Class ExportToBlnFile
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public WithEvents InsideFlag As System.Windows.Forms.CheckBox
    Public WithEvents OpenBlnFile As System.Windows.Forms.Button
    Public WithEvents BlnFileName As System.Windows.Forms.TextBox
    Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
    Public WithEvents ExportButton As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.InsideFlag = New System.Windows.Forms.CheckBox
        Me.OpenBlnFile = New System.Windows.Forms.Button
        Me.BlnFileName = New System.Windows.Forms.TextBox
        Me.CancelButton_Renamed = New System.Windows.Forms.Button
        Me.ExportButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.SuspendLayout()
        '
        'InsideFlag
        '
        Me.InsideFlag.BackColor = System.Drawing.SystemColors.Control
        Me.InsideFlag.Checked = True
        Me.InsideFlag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.InsideFlag.Cursor = System.Windows.Forms.Cursors.Default
        Me.InsideFlag.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.InsideFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.InsideFlag.Location = New System.Drawing.Point(8, 48)
        Me.InsideFlag.Name = "InsideFlag"
        Me.InsideFlag.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.InsideFlag.Size = New System.Drawing.Size(208, 16)
        Me.InsideFlag.TabIndex = 5
        Me.InsideFlag.Text = "Region outside areas is to be blanked"
        '
        'OpenBlnFile
        '
        Me.OpenBlnFile.BackColor = System.Drawing.SystemColors.Control
        Me.OpenBlnFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.OpenBlnFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenBlnFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OpenBlnFile.Location = New System.Drawing.Point(264, 21)
        Me.OpenBlnFile.Name = "OpenBlnFile"
        Me.OpenBlnFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OpenBlnFile.Size = New System.Drawing.Size(25, 25)
        Me.OpenBlnFile.TabIndex = 4
        Me.OpenBlnFile.Text = "..."
        '
        'BlnFileName
        '
        Me.BlnFileName.AcceptsReturn = True
        Me.BlnFileName.AutoSize = False
        Me.BlnFileName.BackColor = System.Drawing.SystemColors.Window
        Me.BlnFileName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.BlnFileName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BlnFileName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.BlnFileName.Location = New System.Drawing.Point(8, 24)
        Me.BlnFileName.MaxLength = 0
        Me.BlnFileName.Name = "BlnFileName"
        Me.BlnFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.BlnFileName.Size = New System.Drawing.Size(249, 19)
        Me.BlnFileName.TabIndex = 2
        Me.BlnFileName.Text = ""
        '
        'CancelButton_Renamed
        '
        Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.CancelButton_Renamed.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CancelButton_Renamed.Location = New System.Drawing.Point(312, 40)
        Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
        Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CancelButton_Renamed.Size = New System.Drawing.Size(81, 25)
        Me.CancelButton_Renamed.TabIndex = 1
        Me.CancelButton_Renamed.Text = "Cancel"
        '
        'ExportButton
        '
        Me.ExportButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExportButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExportButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ExportButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExportButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExportButton.Location = New System.Drawing.Point(312, 8)
        Me.ExportButton.Name = "ExportButton"
        Me.ExportButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ExportButton.Size = New System.Drawing.Size(81, 25)
        Me.ExportButton.TabIndex = 0
        Me.ExportButton.Text = "Export"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(137, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Blanking file path"
        '
        'ExportToBlnFile
        '
        Me.AcceptButton = Me.ExportButton
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.CancelButton_Renamed
        Me.ClientSize = New System.Drawing.Size(402, 73)
        Me.Controls.Add(Me.InsideFlag)
        Me.Controls.Add(Me.OpenBlnFile)
        Me.Controls.Add(Me.BlnFileName)
        Me.Controls.Add(Me.CancelButton_Renamed)
        Me.Controls.Add(Me.ExportButton)
        Me.Controls.Add(Me.Label1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Location = New System.Drawing.Point(184, 250)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExportToBlnFile"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Export to GoldenSoftware Blanking file"
        Me.ResumeLayout(False)

    End Sub
#End Region


    Private Sub OpenBlnFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OpenBlnFile.Click

        SaveFileDialog.Filter() = "GoldenSoftware Blanking (*.bln)|*.bln"

        If SaveFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            BlnFileName.Text = SaveFileDialog.FileName
        End If
    End Sub
End Class