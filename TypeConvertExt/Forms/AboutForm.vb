Imports Microsoft.Win32

Public Class AboutForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents VersionTypeLabel As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BlogLinkLabel As System.Windows.Forms.LinkLabel
    Friend WithEvents HomeLinkLabel As System.Windows.Forms.LinkLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AboutForm))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.BlogLinkLabel = New System.Windows.Forms.LinkLabel()
        Me.HomeLinkLabel = New System.Windows.Forms.LinkLabel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.VersionTypeLabel = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.LinkLabel1)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.BlogLinkLabel)
        Me.GroupBox2.Controls.Add(Me.HomeLinkLabel)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.VersionTypeLabel)
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(229, 272)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(85, 173)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(134, 13)
        Me.LinkLabel1.TabIndex = 25
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "support@geoblogspot.com"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(11, 174)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Support e-mail:"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(77, 226)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "Ok"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(74, 74)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(115, 16)
        Me.Label6.TabIndex = 22
        Me.Label6.Tag = ""
        Me.Label6.Text = "Build: 15.12.2010"
        '
        'BlogLinkLabel
        '
        Me.BlogLinkLabel.Location = New System.Drawing.Point(11, 93)
        Me.BlogLinkLabel.Name = "BlogLinkLabel"
        Me.BlogLinkLabel.Size = New System.Drawing.Size(70, 18)
        Me.BlogLinkLabel.TabIndex = 21
        Me.BlogLinkLabel.TabStop = True
        Me.BlogLinkLabel.Text = "Blog"
        '
        'HomeLinkLabel
        '
        Me.HomeLinkLabel.Location = New System.Drawing.Point(11, 74)
        Me.HomeLinkLabel.Name = "HomeLinkLabel"
        Me.HomeLinkLabel.Size = New System.Drawing.Size(70, 23)
        Me.HomeLinkLabel.TabIndex = 20
        Me.HomeLinkLabel.TabStop = True
        Me.HomeLinkLabel.Text = "Home page"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(11, 111)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(100, 16)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Programming:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(27, 145)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(152, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Developer:   Michael Barsky"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(27, 129)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(152, 16)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Leader:        Valery Hronusov"
        '
        'VersionTypeLabel
        '
        Me.VersionTypeLabel.Location = New System.Drawing.Point(75, 52)
        Me.VersionTypeLabel.Name = "VersionTypeLabel"
        Me.VersionTypeLabel.Size = New System.Drawing.Size(149, 16)
        Me.VersionTypeLabel.TabIndex = 8
        Me.VersionTypeLabel.Tag = ""
        Me.VersionTypeLabel.Text = "Trial version"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.Control
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(7, 14)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(57, 34)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 10
        Me.PictureBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.Location = New System.Drawing.Point(66, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(162, 35)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Typeconvert for ArcGIS 10" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "v.3.0.0"
        '
        'AboutForm
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(240, 290)
        Me.Controls.Add(Me.GroupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "About"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Const HomeURL As String = "http://xbbster.googlepages.com/typeconvert"
    Const GISLabURL As String = "http://xbbster.googlepages.com"
    Const BlogURL As String = "http://gisplanet.blogspot.com"
    Const BuyURL As String = HomeURL
    Const RegROOT As String = "SOFTWARE\\GIS Center\\TypeConvert"

    Private Sub RegistryForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Try
        '    Dim reg As RegistryKey
        '    reg = Registry.CurrentUser.OpenSubKey(RegROOT, True)
        '    RegName.Text = reg.GetValue("RegName", "")
        '    RegKey.Text = reg.GetValue("RegKey", "")
        '    Select Case CheckKey(RegName.Text, RegKey.Text)
        '        Case mdlReg.enumVersionType.vtFull
        VersionTypeLabel.Text = "Full version"
        '        Case mdlReg.enumVersionType.vtExclusive
        'VersionTypeLabel.Text = "Exclusive version"
        '        Case mdlReg.enumVersionType.vtLimited
        'If DaysLeft() > 0 Then
        '    VersionTypeLabel.Text = "Trial version (" & DaysLeft() & " days left)"
        'ElseIf DaysLeft() = 0 Then
        '    VersionTypeLabel.Text = "Trial version (Last day of use)"
        'Else
        '    VersionTypeLabel.Text = "Time of use has expired"
        'End If
        '    End Select

        'Catch ex As Exception

        'End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Process.Start(BuyURL)
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles HomeLinkLabel.LinkClicked
        Process.Start(HomeURL)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Process.Start(HomeURL)
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start(BuyURL)
    End Sub
    Private Sub BlogLinkLabel_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles BlogLinkLabel.LinkClicked
        Process.Start(BlogURL)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Process.Start("mailto:support@geoblogspot.com")

        Catch ex As Exception

        End Try
    End Sub
End Class
