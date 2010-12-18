Public Class ConvertToGE
    Inherits System.Windows.Forms.Form
    Private isUVRenderer As Boolean

#Region " Код, автоматически созданный конструктором форм Windows "

    Public Sub New()
        MyBase.New()

        'Этот вызов требуется конструктором форм Windows.
        InitializeComponent()

        'Добавьте код инициализации после вызова InitializeComponent()

    End Sub

    'Форма переопределяет метод Dispose для очистки списка компонентов.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Требуется конструктором форм Windows
    Private components As System.ComponentModel.IContainer

    'ПРИМЕЧАНИЕ: следующая процедура требуется для конструктора форм Windows.
    'Ее можно изменить в конструкторе форм Windows.  
    'Не изменяйте ее в редакторе исходного текста.
    Public WithEvents OpenBlnFile As System.Windows.Forms.Button
    Public WithEvents CancelButton_Renamed As System.Windows.Forms.Button
    Public WithEvents SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Public WithEvents ExportButton As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SetValue As System.Windows.Forms.TextBox
    Friend WithEvents cbAltitudeMode As System.Windows.Forms.ComboBox
    Friend WithEvents cbAttributeValue As System.Windows.Forms.ComboBox
    Friend WithEvents lbAltitudeMode As System.Windows.Forms.Label
    Friend WithEvents gbExtrude As System.Windows.Forms.GroupBox
    Friend WithEvents ExtrudeFlag As System.Windows.Forms.CheckBox
    Friend WithEvents rbAttribute As System.Windows.Forms.RadioButton
    Friend WithEvents rbValue As System.Windows.Forms.RadioButton
    Public WithEvents kmlFileName As System.Windows.Forms.TextBox
    Friend WithEvents LabelFieldList As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PlacemarkFlag As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cbOpenInGE As System.Windows.Forms.CheckBox
    Friend WithEvents DistributeFlag As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.OpenBlnFile = New System.Windows.Forms.Button
        Me.kmlFileName = New System.Windows.Forms.TextBox
        Me.CancelButton_Renamed = New System.Windows.Forms.Button
        Me.SaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.ExportButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbAltitudeMode = New System.Windows.Forms.ComboBox
        Me.lbAltitudeMode = New System.Windows.Forms.Label
        Me.gbExtrude = New System.Windows.Forms.GroupBox
        Me.cbAttributeValue = New System.Windows.Forms.ComboBox
        Me.SetValue = New System.Windows.Forms.TextBox
        Me.rbAttribute = New System.Windows.Forms.RadioButton
        Me.rbValue = New System.Windows.Forms.RadioButton
        Me.ExtrudeFlag = New System.Windows.Forms.CheckBox
        Me.LabelFieldList = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.PlacemarkFlag = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cbOpenInGE = New System.Windows.Forms.CheckBox
        Me.DistributeFlag = New System.Windows.Forms.CheckBox
        Me.gbExtrude.SuspendLayout()
        Me.SuspendLayout()
        '
        'OpenBlnFile
        '
        Me.OpenBlnFile.BackColor = System.Drawing.SystemColors.Control
        Me.OpenBlnFile.Cursor = System.Windows.Forms.Cursors.Default
        Me.OpenBlnFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenBlnFile.ForeColor = System.Drawing.SystemColors.ControlText
        Me.OpenBlnFile.Location = New System.Drawing.Point(266, 21)
        Me.OpenBlnFile.Name = "OpenBlnFile"
        Me.OpenBlnFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.OpenBlnFile.Size = New System.Drawing.Size(25, 25)
        Me.OpenBlnFile.TabIndex = 10
        Me.OpenBlnFile.Text = "..."
        '
        'kmlFileName
        '
        Me.kmlFileName.AcceptsReturn = True
        Me.kmlFileName.AutoSize = False
        Me.kmlFileName.BackColor = System.Drawing.SystemColors.Window
        Me.kmlFileName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.kmlFileName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kmlFileName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.kmlFileName.Location = New System.Drawing.Point(13, 24)
        Me.kmlFileName.MaxLength = 0
        Me.kmlFileName.Name = "kmlFileName"
        Me.kmlFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.kmlFileName.Size = New System.Drawing.Size(249, 19)
        Me.kmlFileName.TabIndex = 8
        Me.kmlFileName.Text = ""
        '
        'CancelButton_Renamed
        '
        Me.CancelButton_Renamed.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton_Renamed.Cursor = System.Windows.Forms.Cursors.Default
        Me.CancelButton_Renamed.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton_Renamed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CancelButton_Renamed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CancelButton_Renamed.Location = New System.Drawing.Point(192, 328)
        Me.CancelButton_Renamed.Name = "CancelButton_Renamed"
        Me.CancelButton_Renamed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.CancelButton_Renamed.Size = New System.Drawing.Size(81, 25)
        Me.CancelButton_Renamed.TabIndex = 7
        Me.CancelButton_Renamed.Text = "Cancel"
        '
        'ExportButton
        '
        Me.ExportButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExportButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExportButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ExportButton.Enabled = False
        Me.ExportButton.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExportButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExportButton.Location = New System.Drawing.Point(104, 328)
        Me.ExportButton.Name = "ExportButton"
        Me.ExportButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ExportButton.Size = New System.Drawing.Size(81, 25)
        Me.ExportButton.TabIndex = 6
        Me.ExportButton.Text = "Export"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(12, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(137, 17)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Google Earth file path"
        '
        'cbAltitudeMode
        '
        Me.cbAltitudeMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbAltitudeMode.Enabled = False
        Me.cbAltitudeMode.Items.AddRange(New Object() {"Clamped To Ground", "Relative To Ground", "Absolute"})
        Me.cbAltitudeMode.Location = New System.Drawing.Point(16, 176)
        Me.cbAltitudeMode.Name = "cbAltitudeMode"
        Me.cbAltitudeMode.Size = New System.Drawing.Size(256, 21)
        Me.cbAltitudeMode.TabIndex = 11
        '
        'lbAltitudeMode
        '
        Me.lbAltitudeMode.Enabled = False
        Me.lbAltitudeMode.Location = New System.Drawing.Point(16, 160)
        Me.lbAltitudeMode.Name = "lbAltitudeMode"
        Me.lbAltitudeMode.Size = New System.Drawing.Size(100, 16)
        Me.lbAltitudeMode.TabIndex = 12
        Me.lbAltitudeMode.Text = "Altitude Mode"
        '
        'gbExtrude
        '
        Me.gbExtrude.Controls.Add(Me.cbAttributeValue)
        Me.gbExtrude.Controls.Add(Me.SetValue)
        Me.gbExtrude.Controls.Add(Me.rbAttribute)
        Me.gbExtrude.Controls.Add(Me.rbValue)
        Me.gbExtrude.Enabled = False
        Me.gbExtrude.Location = New System.Drawing.Point(16, 208)
        Me.gbExtrude.Name = "gbExtrude"
        Me.gbExtrude.Size = New System.Drawing.Size(256, 81)
        Me.gbExtrude.TabIndex = 13
        Me.gbExtrude.TabStop = False
        Me.gbExtrude.Text = "Extrude based on"
        '
        'cbAttributeValue
        '
        Me.cbAttributeValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbAttributeValue.Location = New System.Drawing.Point(100, 16)
        Me.cbAttributeValue.Name = "cbAttributeValue"
        Me.cbAttributeValue.Size = New System.Drawing.Size(148, 21)
        Me.cbAttributeValue.TabIndex = 3
        '
        'SetValue
        '
        Me.SetValue.Enabled = False
        Me.SetValue.Location = New System.Drawing.Point(99, 48)
        Me.SetValue.Name = "SetValue"
        Me.SetValue.Size = New System.Drawing.Size(149, 20)
        Me.SetValue.TabIndex = 2
        Me.SetValue.Text = ""
        '
        'rbAttribute
        '
        Me.rbAttribute.Checked = True
        Me.rbAttribute.Location = New System.Drawing.Point(8, 16)
        Me.rbAttribute.Name = "rbAttribute"
        Me.rbAttribute.TabIndex = 1
        Me.rbAttribute.TabStop = True
        Me.rbAttribute.Text = "Attribute field"
        '
        'rbValue
        '
        Me.rbValue.Location = New System.Drawing.Point(8, 46)
        Me.rbValue.Name = "rbValue"
        Me.rbValue.TabIndex = 0
        Me.rbValue.Text = "Value"
        '
        'ExtrudeFlag
        '
        Me.ExtrudeFlag.Location = New System.Drawing.Point(16, 136)
        Me.ExtrudeFlag.Name = "ExtrudeFlag"
        Me.ExtrudeFlag.TabIndex = 14
        Me.ExtrudeFlag.Text = "Extrude"
        '
        'LabelFieldList
        '
        Me.LabelFieldList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LabelFieldList.Location = New System.Drawing.Point(71, 96)
        Me.LabelFieldList.Name = "LabelFieldList"
        Me.LabelFieldList.Size = New System.Drawing.Size(192, 21)
        Me.LabelFieldList.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(12, 100)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(59, 16)
        Me.Label2.TabIndex = 16
        Me.Label2.Text = "Label field"
        '
        'PlacemarkFlag
        '
        Me.PlacemarkFlag.Checked = True
        Me.PlacemarkFlag.CheckState = System.Windows.Forms.CheckState.Checked
        Me.PlacemarkFlag.Location = New System.Drawing.Point(13, 45)
        Me.PlacemarkFlag.Name = "PlacemarkFlag"
        Me.PlacemarkFlag.Size = New System.Drawing.Size(216, 24)
        Me.PlacemarkFlag.TabIndex = 17
        Me.PlacemarkFlag.Text = "Convert each object as a placemark"
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(5, 128)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(287, 3)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        '
        'cbOpenInGE
        '
        Me.cbOpenInGE.Checked = True
        Me.cbOpenInGE.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbOpenInGE.Location = New System.Drawing.Point(16, 296)
        Me.cbOpenInGE.Name = "cbOpenInGE"
        Me.cbOpenInGE.Size = New System.Drawing.Size(152, 24)
        Me.cbOpenInGE.TabIndex = 19
        Me.cbOpenInGE.Text = "Open in Google Earth "
        '
        'DistributeFlag
        '
        Me.DistributeFlag.Location = New System.Drawing.Point(13, 68)
        Me.DistributeFlag.Name = "DistributeFlag"
        Me.DistributeFlag.Size = New System.Drawing.Size(251, 24)
        Me.DistributeFlag.TabIndex = 20
        Me.DistributeFlag.Text = "Distribute features using a categories legend"
        '
        'ConvertToGE
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(298, 367)
        Me.Controls.Add(Me.DistributeFlag)
        Me.Controls.Add(Me.cbOpenInGE)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.PlacemarkFlag)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LabelFieldList)
        Me.Controls.Add(Me.ExtrudeFlag)
        Me.Controls.Add(Me.gbExtrude)
        Me.Controls.Add(Me.lbAltitudeMode)
        Me.Controls.Add(Me.cbAltitudeMode)
        Me.Controls.Add(Me.OpenBlnFile)
        Me.Controls.Add(Me.kmlFileName)
        Me.Controls.Add(Me.CancelButton_Renamed)
        Me.Controls.Add(Me.ExportButton)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ConvertToGE"
        Me.Text = "Convert To Google Earth"
        Me.gbExtrude.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Property isUniqueValueRenderer() As Boolean
        Get
            isUniqueValueRenderer = isUVRenderer
        End Get
        Set(ByVal Value As Boolean)
            isUVRenderer = Value
        End Set
    End Property
    Private Sub OpenBlnFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OpenBlnFile.Click

        SaveFileDialog.Filter() = "Google Earth file (*.kml)|*.kml"

        If SaveFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            kmlFileName.Text = SaveFileDialog.FileName
        End If
    End Sub
    Private Sub ExtrudeFlag_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExtrudeFlag.CheckedChanged
        lbAltitudeMode.Enabled = ExtrudeFlag.Checked
        cbAltitudeMode.Enabled = ExtrudeFlag.Checked
        gbExtrude.Enabled = ExtrudeFlag.Checked
        EnableChange()
    End Sub

    Private Sub rbAttribute_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbAttribute.CheckedChanged
        EnableChange()
    End Sub

    Private Sub rbValue_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbValue.CheckedChanged
        EnableChange()
    End Sub
    Private Sub EnableChange()
        cbAttributeValue.Enabled = rbAttribute.Checked And ExtrudeFlag.Checked
        SetValue.Enabled = rbValue.Checked And ExtrudeFlag.Checked
    End Sub

    Private Sub kmlFileName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles kmlFileName.TextChanged
        ExportButton.Enabled = (kmlFileName.TextLength > 0)
    End Sub

    Private Sub PlacemarkFlag_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PlacemarkFlag.CheckedChanged
        LabelFieldList.Enabled = PlacemarkFlag.Checked And (LabelFieldList.Items.Count > 0)
        'DistributeFlag.Enabled = PlacemarkFlag.Checked And isUVRenderer
        DistributeFlag.Enabled = isUVRenderer
    End Sub

    Private Sub ConvertToGE_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cbAltitudeMode.Text = cbAltitudeMode.GetItemText(cbAltitudeMode.Items.Item(0))
        'DistributeFlag.Enabled = PlacemarkFlag.Checked And isUVRenderer
        DistributeFlag.Enabled = isUVRenderer
    End Sub
End Class
