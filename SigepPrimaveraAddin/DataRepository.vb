Imports System.Data
Imports PO.DataClient
Imports PO.Objects

Public NotInheritable Class DataRepository

    Private Shared _Current As DataRepository
    Private Shared syncroLock As New Object

    Private _ProjStatus As Dictionary(Of Integer, String)
    Private _ProjVersions As Dictionary(Of Integer, String)
    Private _Customers As Dictionary(Of String, String)
    Private _Locations As Dictionary(Of Integer, String)
    Private _Company As Dictionary(Of Integer, String)
    Public _companyID As Integer = 1
    Protected Sub New()
    End Sub

    Protected Sub WriteLog(ByVal text As String)
        Try
            Dim cfg As New ConfigHelper()
            Dim logPath As String = cfg.GetConfigKey("logpath")
            If String.IsNullOrEmpty(logPath) Then Return
            FileIO.FileSystem.WriteAllText(logpath, text & Now.ToString(" [yyyy-MM-dd HH:mm:ss]") & vbCrLf, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Shared ReadOnly Property Current(Optional ByVal companyID As Integer = 1) As DataRepository
        Get
            If (_Current Is Nothing) Then
                SyncLock syncroLock
                    _Current = New DataRepository()
                    _Current._companyID = companyID
                End SyncLock
            End If
            Return _Current
        End Get
    End Property
    Public ReadOnly Property Locations() As Dictionary(Of Integer, String)
        Get
            If _Locations Is Nothing Then
                Try

                    Dim companyId As Integer = _companyID '' System.Configuration.ConfigurationManager.AppSettings("companyId")
                    Dim cfg As New ConfigHelper()
                    Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))
                    Dim buf() As Byte = ws.ProjLocationsList(companyId)
                    Dim data As New ZDataSet(buf)

                    _Locations = New Dictionary(Of Integer, String)
                    _Locations.Add(-1, String.Empty)
                    For Each row As DataRow In data.Tables(0).Rows
                        _Locations.Add(Convert.ToInt32(row("C_LOCATION")), row("S_LOCATION"))
                    Next
                Catch ex As Exception
                    WriteLog(ex.Message)
                End Try

            End If
            Return _Locations
        End Get
    End Property
    Public ReadOnly Property Customers() As Dictionary(Of String, String)
        Get
            If _Customers Is Nothing Then
                Try
                    Dim companyId As Integer = _companyID ''System.Configuration.ConfigurationManager.AppSettings("companyId")
                    Dim cfg As New ConfigHelper()
                    Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

                    Dim buf() As Byte = ws.ProjCustomers(companyId)
                    Dim data As New ZDataSet(buf)
                    _Customers = New Dictionary(Of String, String)
                    _Customers.Add(String.Empty, String.Empty)
                    For Each row As DataRow In data.Tables(0).Rows
                        _Customers.Add(row("C_CLI"), row("S_CLI"))
                    Next
                Catch ex As Exception
                    WriteLog(ex.Message)
                End Try
            End If
            Return _Customers
        End Get
    End Property
    Public ReadOnly Property ProjStatus() As Dictionary(Of Integer, String)
        Get
            If _ProjStatus Is Nothing Then

                Dim cfg As New ConfigHelper()
                Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

                Dim buf() As Byte = ws.ProjStatusList()
                Dim data As New ZDataSet(buf)

                _ProjStatus = New Dictionary(Of Integer, String)
                _ProjStatus.Add(-1, String.Empty)
                For Each row As DataRow In data.Tables(0).Rows
                    _ProjStatus.Add(Convert.ToInt32(row("STATUSID")), row("STATUSNAME"))
                Next

            End If
            Return _ProjStatus
        End Get
    End Property

    Public Function CompanyList(userId As String) As Dictionary(Of Integer, String)
        If _Company Is Nothing Then

            Dim cfg As New ConfigHelper()
            Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

            Dim buf As Byte() = ws.CompanyList(userId)
            Dim data As New ZDataSet(buf)

            _Company = New Dictionary(Of Integer, String)
            For Each row As DataRow In data.Tables(0).Rows
                _Company.Add(Convert.ToInt32(row("C_AZD")), row("S_AZD"))
            Next

        End If
        Return _Company
    End Function

    Public ReadOnly Property ProjSnapshotversion(ByVal projID As Integer) As Dictionary(Of Integer, String)
        Get
            Dim cfg As New ConfigHelper()
            Dim ws As New PM6DBServiceController.PM6DBServiceController(cfg.GetConfigKey("WebServiceUrl"))

            Dim buf() As Byte = ws.ProjectSanpshots(projID)
            Dim data As New ZDataSet(buf)
            _ProjVersions = New Dictionary(Of Integer, String)
            _ProjVersions.Add(projID, "Current version")
            For Each row As DataRow In data.Tables(0).Rows
                _ProjVersions.Add(Convert.ToInt32(row("C_PROG_SNAPSHOT")), row("CREATION_DATE"))
            Next
            Return _ProjVersions
        End Get
    End Property

End Class


   