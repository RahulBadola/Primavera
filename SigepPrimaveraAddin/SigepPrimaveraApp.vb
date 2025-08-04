Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports System.Reflection
Imports System.IO
Imports System.Text
Imports Microsoft.Win32
Imports System.Linq

<ComSourceInterfaces(GetType(SigepPrimaveraInterface)), _
ComVisible(True), _
ClassInterface(ClassInterfaceType.AutoDual), _
ProgId("PrimaveraSigepLauncher.Application")> _
Public Class SigepPrimaveraApp
    Implements SigepPrimaveraInterface

    Public Shared idpm6parameter As Integer = 0
    Private Shared _POProj As Integer = 0
    'il vbs passa al exe il codice progetto 
    <ComVisible(False), STAThread()> _
    Public Shared Sub Main(ByVal args As String())
        ''args = New String() {"-1", "1", "", "saveas"} '' uncomment to test 
        Dim isSaveAs As Boolean = args.Contains("saveas", StringComparer.InvariantCultureIgnoreCase)

        Dim user As String = String.Empty
        If args.Length >= 3 Then
            user = args(2)
        End If
        ''Changes by Anubrij
        ''Implement publish as functionality 
        If isSaveAs Then
            _POProj = ReadPOProjFromCfg(args(0), user)
        Else
            _POProj = -1
        End If


        Dim t As New Threading.Thread(New Threading.ParameterizedThreadStart(AddressOf runTask))
        t.IsBackground = True
        If args(1) = 1 Then
            t.Start(args(0))
        End If
        Dim direction As Integer = args(1)
        Dim id As Integer = args(0)

        If _POProj = -1 Then
            '  idpm6parameter = id
            ' MsgBox("Params : " & args(0) & "-" & args(1) & "-" & args(2))
            ExecuteLauncher(id, direction, user)
        Else
            ''Changes by Anubrij
            ''Implement publish as functionality 
            If _POProj = 0 Then
                Return
            End If
            ExecuteLauncher(id, _POProj, direction, user)
        End If



    End Sub


    <ComVisible(True)>
    Public Sub Run(ByVal projId As String, ByVal direction As String, ByVal user As String) Implements SigepPrimaveraInterface.Run
        AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", Assembly.GetExecutingAssembly().Location & ".config")
        ExecuteLauncher(CInt(projId), CInt(direction), CStr(user))
    End Sub
    ''' <summary>
    ''' Added by Anubrij 
    ''' </summary>
    ''' <param name="user"></param>
    ''' <returns>Get the selected result from project browser popup</returns>
    ''' <remarks></remarks>
    Private Shared Function ReadPOProjFromCfg(p6ProjId As Integer, ByVal user As String) As Integer
        Dim result As Integer = 0
        result = ProjBrowser.SelectProject(p6ProjId, user)
        Return result
    End Function
    Public Shared Sub runTask(ByVal o As Object)
        '  Console.WriteLine("inizio task monitor")

        Dim s As New SplashScreen
        s.Location.Offset(300, 900)
        s.Show()
        s.Activate()
        s.BringToFront()
        s.Focus()
        s.TopMost = True
        Dim cfg As New ConfigHelper()
        Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

        ws.Timeout = 20 * 60 * 1000 '10 minuti
        Dim str As String = String.Empty
        While 1 = 1
            '     Console.WriteLine("ciclo task monitor - " & CInt(o)
            Try
                str = (ws.DisplayStatus(CInt(o)).Split("-")(1))
            Catch

                Exit Sub
            End Try
            s.Label3.Text = str.Replace("OK", "")
            s.Refresh()
            Threading.Thread.Sleep(4000)
            If idpm6parameter = 1 Then
                '         Console.WriteLine("uscito task monitor - " & CInt(o))
                s.Hide()
                s.Dispose()
                Threading.Thread.CurrentThread.Abort()
                '  Threading.Thread.CurrentThread.Join()
                Exit While
            End If
        End While

    End Sub
    ''' <summary>
    ''' Added by Anubrij
    ''' </summary>
    ''' <param name="idPM5">primavera id</param>
    ''' <param name="idPO">sigep id</param>
    ''' <param name="direction"></param>
    ''' <param name="user"></param>
    ''' <remarks></remarks>
    Private Shared Sub ExecuteLauncher(ByVal idPM5 As String, ByVal idPO As Integer, ByVal direction As Integer, ByVal user As String)
        Dim cfg As New ConfigHelper()
        Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

        ws.Timeout = 20 * 60 * 1000 '10 minuti
        Dim printResponse As Boolean
        '   Console.WriteLine("inizio task principale - " & idPM5)
        WriteLog("Processing start at :" & Now)
        '/////////////// 1: se stiamo aggiornando SIGEP da menu di PM5//////////////////////////////////////
        'il parametro di ingresso è  il codice del progetto primavera appena modificato
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ' commented in order to have a single "ExecuteLauncher" procedure - MP 2015-12-22
        'If idPO = 0 Then
        '    MsgBox("No project selected")
        '    Return
        'End If
        ' commented in order to have a single "ExecuteLauncher" procedure - MP 2015-12-22
        If idPO > 0 Then            ' added "if" condition in order to have a single "ExecuteLauncher" procedure - MP 2015-12-22
            If Not ws.UpdateProjectAssociation(idPM5, idPO) Then
                MsgBox("Project publish fail")
                Return
            End If
        End If                      ' END - added "if" condition in order to have a single "ExecuteLauncher" procedure - MP 2015-12-22
        If direction = 1 Then
            Try
                '  Dim s As New SplashScreen
                '  Threading.Thread.Sleep(3000)
                '  s.Show()
                '  s.Opacity = 0
                '  s.Refresh()
                '  s.Activate()
                '  s.BringToFront()
                '  s.Focus()
                '  s.TopMost = True

                printResponse = ws.UpdateProject(idPM5, "pm6")
                ' s.Refresh()
                If printResponse = False Then
                    '    s.Label1.Text = "Publishing Aborted."
                    MsgBox("Another user is still publishing this project or this project not exists in Sigep", MsgBoxStyle.Critical, "Error")
                    Return
                Else
                End If
                idpm6parameter = 1
                Threading.Thread.Sleep(2000)
                's.Refresh()
                WriteLog("Processing end at :" & Now)
                's.Hide()
                's.Dispose()


                Dim msgPartialUpdate As Integer
                msgPartialUpdate = ws.GetSigepProjectStatus(idPM5)
                If msgPartialUpdate = 2 Then
                    MsgBox("The project has been successfully published to SIGEP", MsgBoxStyle.Information, "Publishing")
                ElseIf msgPartialUpdate >= 0 Then
                    MsgBox("The project has been successfully published to SIGEP." & vbLf & "SIGEP project status is not ""In Progress""; actual dates and progress will not be updated ", MsgBoxStyle.Information, "Publishing")
                ElseIf msgPartialUpdate = -1 Then
                    MsgBox("Error Retriving SIGEP Project Status Information")
                End If

            Catch ex As Exception
                WriteLog(ex.Message)
            End Try

            '' '' ''    '///// 0: adesso uso questa direzione per aprire da sigep PM.exe 
            '' '' ''    '// il parametro di ingresso è  il codice  del progetto sigep 
        ElseIf direction = 0 Then
            Try
                Dim p6Passw As String = String.Empty
                Try
                    p6Passw = ws.GetUserPrimaveraPwd(user)
                    ws.ReplacePM6StartProject(ws.GetAssociateProject(idPM5), user)
                    WriteLog("user: " & user & "; passw: " & p6Passw)
                Catch ex As Exception
                    WriteLog(ex.Message & "impostato a : " & ws.Timeout)
                End Try
                'per aprire pm6.exe sull'utente sigep occorre che projects passi a run_app l'utente sigep e il run_app lo passi a questo exe. 
                'poi si fa il replace  della stringa contenuta nel config    'username=admin /password=admin' con il nome effettivo (la pwd è meglio lasciarla univoca)
                'Dim exePath As String = System.Configuration.ConfigurationManager.AppSettings("pm5_Path")
                'Dim workDir As String = System.Configuration.ConfigurationManager.AppSettings("pm5_workdir")
                'Dim p6DBAlias As String = System.Configuration.ConfigurationManager.AppSettings("pm5_Alias")

                Dim exePath As String = cfg.GetConfigKey("pm5_Path")
                Dim workDir As String = cfg.GetConfigKey("pm5_workdir")
                Dim p6DBAlias As String = cfg.GetConfigKey("pm5_Alias")

                'Dim str As String = String.Empty
                'str = Replace(System.Configuration.ConfigurationManager.AppSettings("pm5_Path"), "username=admin", "username=" & user)
                'str = Replace(System.Configuration.ConfigurationManager.AppSettings("pm5_Path"), "/username=admin /password=admin", "/username=" & user & " " & "/password=" & p6Passw)

                Dim info As New ProcessStartInfo()
                info.FileName = exePath
                info.WorkingDirectory = workDir
                info.Arguments = "/username=" & user & " /password=" & p6Passw & " /alias=" & p6DBAlias
                info.UseShellExecute = True
                System.Diagnostics.Process.Start(info)
                WriteLog("Exe Path: " & exePath)
                WriteLog("Working directory: " & workDir)
                WriteLog("Arguments: " & info.Arguments)
                'WriteLog("shell command: " & str)
                'Shell(str)
            Catch ex As Exception
                WriteLog(ex.Message)
            End Try
        End If
    End Sub

    Private Shared Sub ExecuteLauncher(ByVal idPM5 As String, ByVal direction As Integer, ByVal user As String)
        ExecuteLauncher(idPM5, 0, direction, user)
        'Dim cfg As New ConfigHelper()
        'Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

        'ws.Timeout = 20 * 60 * 1000 '10 minuti
        'Dim printResponse As Boolean
        ''   Console.WriteLine("inizio task principale - " & idPM5)
        'WriteLog("Processing start at :" & Now)
        ''/////////////// 1: se stiamo aggiornando SIGEP da menu di PM5//////////////////////////////////////
        ''il parametro di ingresso è  il codice del progetto primavera appena modificato
        ''///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        'If direction = 1 Then
        '    Try
        '        '  Dim s As New SplashScreen
        '        '  Threading.Thread.Sleep(3000)
        '        '  s.Show()
        '        '  s.Opacity = 0
        '        '  s.Refresh()
        '        '  s.Activate()
        '        '  s.BringToFront()
        '        '  s.Focus()
        '        '  s.TopMost = True

        '        printResponse = ws.UpdateProject(idPM5, "pm6")
        '        ' s.Refresh()
        '        If printResponse = False Then
        '            '    s.Label1.Text = "Publishing Aborted."
        '            MsgBox("Another user is still publishing this project or this project not exists in Sigep", MsgBoxStyle.Critical, "Error")
        '        Else
        '        End If
        '        idpm6parameter = 1
        '        Threading.Thread.Sleep(2000)
        '        's.Refresh()
        '        WriteLog("Processing end at :" & Now)
        '        's.Hide()
        '        's.Dispose()


        '        Dim msgPartialUpdate As Integer
        '        msgPartialUpdate = ws.GetSigepProjectStatus(idPM5)
        '        If msgPartialUpdate = 2 Then
        '            MsgBox("The project has been successfully published to SIGEP", MsgBoxStyle.Information, "Publishing")
        '        ElseIf msgPartialUpdate >= 0 Then
        '            MsgBox("The project has been successfully published to SIGEP." & vbLf & "SIGEP project status is not ""In Progress""; actual dates and progress will not be updated ", MsgBoxStyle.Information, "Publishing")
        '        ElseIf msgPartialUpdate = -1 Then
        '            MsgBox("Error Retriving SIGEP Project Status Information")
        '        End If

        '    Catch ex As Exception
        '        WriteLog(ex.Message & "impostato a : " & ws.Timeout)
        '    End Try

        '    '' '' ''    '///// 0: adesso uso questa direzione per aprire da sigep PM.exe 
        '    '' '' ''    '// il parametro di ingresso è  il codice  del progetto sigep 
        'ElseIf direction = 0 Then
        '    Try
        '        Dim p6Passw As String = String.Empty
        '        Try
        '            p6Passw = ws.GetUserPrimaveraPwd(user)
        '            ws.ReplacePM6StartProject(ws.GetAssociateProject(idPM5), user)
        '            WriteLog("user: " & user & "; passw: " & p6Passw)
        '        Catch ex As Exception
        '            WriteLog(ex.Message)
        '        End Try

        '        Dim exePath As String = cfg.GetConfigKey("pm5_Path")
        '        Dim workDir As String = cfg.GetConfigKey("pm5_workdir")
        '        Dim p6DBAlias As String = cfg.GetConfigKey("pm5_Alias")

        '        'exePath = System.Configuration.ConfigurationManager.AppSettings("pm5_Path")
        '        'workDir = System.Configuration.ConfigurationManager.AppSettings("pm5_workdir")
        '        'p6DBAlias = System.Configuration.ConfigurationManager.AppSettings("pm5_Alias")
        '        'per aprire pm6.exe sull'utente sigep occorre che projects passi a run_app l'utente sigep e il run_app lo passi a questo exe. 
        '        'poi si fa il replace  della stringa contenuta nel config    'username=admin /password=admin' con il nome effettivo (la pwd è meglio lasciarla univoca)
        '        ''Dim str As String = String.Empty
        '        ''str = Replace(System.Configuration.ConfigurationManager.AppSettings("pm5_Path"), "username=admin", "username=" & user)
        '        ''str = Replace(System.Configuration.ConfigurationManager.AppSettings("pm5_Path"), "/username=admin /password=admin", "/username=" & user & " " & "/password=" & p6Passw)

        '        Dim info As New ProcessStartInfo()
        '        info.FileName = exePath
        '        info.WorkingDirectory = workDir
        '        info.Arguments = "/username=" & user & " /password=" & p6Passw & " /alias=" & p6DBAlias
        '        info.UseShellExecute = True
        '        System.Diagnostics.Process.Start(info)
        '        WriteLog("Exe Path: " & exePath)
        '        WriteLog("Working directory: " & workDir)
        '        WriteLog("Arguments: " & info.Arguments)
        '        'WriteLog("shell command: " & str)
        '        'Shell(str)
        '    Catch ex As Exception
        '        WriteLog(ex.Message)
        '    End Try
        'End If
    End Sub

    <ComRegisterFunction()> _
    Public Shared Sub RegisterClass(ByVal key As String)
        Dim builder As New StringBuilder(key)
        builder.Replace("HKEY_CLASSES_ROOT\", "")
        Dim key4 As RegistryKey = Registry.ClassesRoot.OpenSubKey(builder.ToString, True)
        key4.CreateSubKey("Control").Close()
        Dim key3 As RegistryKey = key4.OpenSubKey("InprocServer32", True)
        key3.SetValue("CodeBase", Assembly.GetExecutingAssembly.CodeBase)
        key3.Close()
        key4.Close()
    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub UnregisterClass(ByVal key As String)
        Dim builder As New StringBuilder(key)
        builder.Replace("HKEY_CLASSES_ROOT\", "")
        Dim key2 As RegistryKey = Registry.ClassesRoot.OpenSubKey(builder.ToString, True)
        key2.DeleteSubKey("Control", False)
        key2.OpenSubKey("InprocServer32", True)
        key2.DeleteSubKey("CodeBase", False)
        key2.Close()
    End Sub

    Public Shared Sub WriteLog(ByVal text As String)
        Try
            Dim cfg As New ConfigHelper()
            Dim logpath As String = cfg.GetConfigKey("logpath")
            If String.IsNullOrEmpty(logpath) Then Return
            FileIO.FileSystem.WriteAllText(logpath, Now.ToString("yyyy-MM-dd HH:mm:ss ") & text & vbCrLf, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
