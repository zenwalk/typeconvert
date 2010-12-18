Imports System.ComponentModel
Imports System.Configuration.Install
Imports Install
Public Class Installer

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add initialization code after the call to InitializeComponent

    End Sub
    Public Overrides Sub Install(ByVal stateSaver As System.Collections.IDictionary)
        MyBase.Install(stateSaver)

        Dim params As Specialized.StringDictionary = Me.Context.Parameters

        If params.ContainsKey("extension") Then
            Dim cmd1 As String = """" + Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles) + "\ArcGIS\bin\ESRIRegAsm.exe" + """"

            Dim part1 As String = params.Item("extension")
            Dim part2 As String = " /p:Desktop /s"
            Dim cmd2 As String = """" + part1 + """" + part2
            ExecuteCommand(cmd1, cmd2, 10000)
            If params.ContainsKey("geoprocessing") Then
                part1 = params.Item("geoprocessing")
                cmd2 = """" + part1 + """" + part2
                ExecuteCommand(cmd1, cmd2, 10000)
            End If
        End If
    End Sub

    Public Overrides Sub Uninstall(ByVal savedState As System.Collections.IDictionary)

        MyBase.Uninstall(savedState)

        Dim params As Specialized.StringDictionary = Me.Context.Parameters
        If params.ContainsKey("extension") Then
            Dim cmd1 As String = """" + Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles) + "\ArcGIS\bin\ESRIRegAsm.exe" + """"
            Dim part1 As String = params.Item("extension")
            Dim part2 As String = " /p:Desktop /u /s"
            Dim cmd2 As String = """" + part1 + """" + part2

            ExecuteCommand(cmd1, cmd2, 10000)

            If params.ContainsKey("geoprocessing") Then
                part1 = params.Item("geoprocessing")
                cmd2 = """" + part1 + """" + part2
                ExecuteCommand(cmd1, cmd2, 10000)
            End If
        End If
    End Sub


    Public Shared Sub ExecuteCommand(ByVal Command1 As String, ByVal Command2 As String, ByVal Timeout As Integer)

        'Set up a ProcessStartInfo using your path to the executable (Command1) and the command line arguments (Command2).
        Dim ProcessInfo As ProcessStartInfo = New ProcessStartInfo(Command1, Command2)
        ProcessInfo.CreateNoWindow = True
        ProcessInfo.UseShellExecute = False

        'Invoke the process.
        Dim Process As Process = Process.Start(ProcessInfo)
        Process.WaitForExit(Timeout)

        'Finish.
        'Dim ExitCode As Integer = Process.ExitCode
        Process.Close()
        'Return ExitCode
    End Sub

End Class
