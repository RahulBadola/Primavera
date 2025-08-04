Imports PO.DataClient

Public Class P6Calendar
    Implements IDisposable
    Public Enum P6CalendarType
        Base = 1
        Project
        Resource
        Task
    End Enum

#Region "Properties related variables"
    Private _LastException As Exception
    Private _CalendarId As Integer
    Private _CalendarName As String
    Private _CalendarType As P6CalendarType
    Private _IsDefault As Boolean
    Private _PrivateForResource As Boolean
    Private _BaseCalendar As P6Calendar
    Private _ProjectId As System.Nullable(Of Integer)
    Private _TaskId As System.Nullable(Of Integer)
    Private _ResourceId As System.Nullable(Of Integer)
    Private _HoursPerDay As Decimal
    Private _HoursPerWeek As Decimal
    Private _HoursPerMonth As Decimal
    Private _HoursYear As Decimal
    Private _CalendarExceptions As String
#End Region

    Public Function LoadCalendar(calendarId As Integer) As P6Calendar
        Me.CalendarId = calendarId
        Dim stmt As String = SelectCalendarStmt() & " and CLNDR_ID = " & calendarId.ToString() & " "
        If Not ReadCalendar(stmt) Then
            Return Nothing
        End If
        Return Me
    End Function
    Public Function LoadProjectCalendar(projectId As Integer) As P6Calendar
        Me.ProjectId = projectId
        Dim stmt As String = SelectCalendarStmt() & " and CLNDR_ID in (select CLNDR_ID from PROJECT where PROJ_ID = " & projectId.ToString() & ") "
        If Not ReadCalendar(stmt) Then
            Return Nothing
        End If
        Return Me
    End Function
    Public Function LoadTaskCalendar(projectId As Integer, taskId As Integer) As P6Calendar
        Me.ProjectId = projectId
        Me.TaskId = taskId
        Dim stmt As String = SelectCalendarStmt() & " and CLNDR_ID in (select CLNDR_ID from TASK where PROJ_ID = " & projectId.ToString() & " and TASK_ID = " & taskId.ToString() & ") "
        If Not ReadCalendar(stmt) Then
            Return Nothing
        End If
        Return Me
    End Function
    Public Function LoadResourceCalendar(resourceId As Integer) As P6Calendar
        Me.ResourceId = resourceId
        Dim stmt As String = SelectCalendarStmt() & " and CLNDR_ID in (select CLNDR_ID from RSRC where RSRC_ID = " & resourceId.ToString() & ") "
        If Not ReadCalendar(stmt) Then
            Return Nothing
        End If
        Return Me
    End Function
    Private Function SelectCalendarStmt() As String
        Return " select CLNDR_ID, DEFAULT_FLAG, RSRC_PRIVATE, CLNDR_NAME " &
            " , PROJ_ID, BASE_CLNDR_ID, CLNDR_TYPE, CLNDR_DATA " &
            " , DAY_HR_CNT, WEEK_HR_CNT, YEAR_HR_CNT, MONTH_HR_CNT " &
            " from CALENDAR " &
            " where DELETE_SESSION_ID is null and DELETE_DATE is null "
    End Function
    Private Function ReadCalendar(stmt As String) As Boolean

        Dim conn As IDataClient = GetDBConn()

        Try
            Dim data As DataSet = conn.ExecuteQuery(stmt)
            If data.Tables(0).Rows.Count = 0 Then
                Me.LastException = New Exception("No calendar retrieved")
                Return False
            End If
            If data.Tables(0).Rows.Count > 1 Then
                Me.LastException = New Exception("Calendar data error!")
                Return False
            End If

            Dim row As DataRow = data.Tables(0).Rows(0)

            Me.CalendarId = Convert.ToInt32(row("CLNDR_ID"))
            Me.CalendarName = DirectCast(row("CLNDR_NAME"), String)
            Me.CalendarType = EncodeP6CalendarType(DirectCast(row("CLNDR_TYPE"), String))
            If row("PROJ_ID") IsNot DBNull.Value Then Me.ProjectId = Convert.ToInt32(row("PROJ_ID"))

            'this.TaskId = 
            'this.ResourceId = 
            If row("DAY_HR_CNT") IsNot DBNull.Value Then Me.HoursPerDay = Convert.ToInt32(row("DAY_HR_CNT"))
            If row("WEEK_HR_CNT") IsNot DBNull.Value Then Me.HoursPerWeek = Convert.ToInt32(row("WEEK_HR_CNT"))
            If row("YEAR_HR_CNT") IsNot DBNull.Value Then Me.HoursYear = Convert.ToInt32(row("YEAR_HR_CNT"))
            If row("MONTH_HR_CNT") IsNot DBNull.Value Then Me.HoursPerMonth = Convert.ToInt32(row("MONTH_HR_CNT"))
            Me.PrivateForResource = (DirectCast(row("RSRC_PRIVATE"), String) = "Y")
            Me.IsDefault = (DirectCast(row("DEFAULT_FLAG"), String) = "Y")
            Me.BaseCalendar = Nothing 'If row("BASE_CLNDR_ID") IsNot DBNull.Value Then Me.BaseCalendar = LoadCalendar(Convert.ToInt32(row("BASE_CLNDR_ID")))
            Me.CalendarExceptions = ReadCalendarExceptions()             'CLNDR_DATA 

            Return True

        Catch ex As Exception
            Me.LastException = ex
            Return False
        End Try
    End Function

#Region "Properties"
    Public Property LastException() As Exception
        Get
            Return _LastException
        End Get
        Private Set
            _LastException = Value
        End Set
    End Property
    Public Property CalendarId() As Integer
        Get
            Return _CalendarId
        End Get
        Private Set
            _CalendarId = Value
        End Set
    End Property
    Public Property CalendarName() As String
        Get
            Return _CalendarName
        End Get
        Private Set
            _CalendarName = Value
        End Set
    End Property
    Public Property CalendarType() As P6CalendarType
        Get
            Return _CalendarType
        End Get
        Private Set
            _CalendarType = Value
        End Set
    End Property
    Public Property IsDefault() As Boolean
        Get
            Return _IsDefault
        End Get
        Private Set
            _IsDefault = Value
        End Set
    End Property
    Public Property PrivateForResource() As Boolean
        Get
            Return _PrivateForResource
        End Get
        Private Set
            _PrivateForResource = Value
        End Set
    End Property
    Public Property BaseCalendar() As P6Calendar
        Get
            Return _BaseCalendar
        End Get
        Private Set
            _BaseCalendar = Value
        End Set
    End Property
    Public Property ProjectId() As System.Nullable(Of Integer)
        Get
            Return _ProjectId
        End Get
        Private Set
            _ProjectId = Value
        End Set
    End Property
    Public Property TaskId() As System.Nullable(Of Integer)
        Get
            Return _TaskId
        End Get
        Private Set
            _TaskId = Value
        End Set
    End Property
    Public Property ResourceId() As System.Nullable(Of Integer)
        Get
            Return _ResourceId
        End Get
        Private Set
            _ResourceId = Value
        End Set
    End Property
    Public Property HoursPerDay() As Decimal
        Get
            Return _HoursPerDay
        End Get
        Private Set
            _HoursPerDay = Value
        End Set
    End Property
    Public Property HoursPerWeek() As Decimal
        Get
            Return _HoursPerWeek
        End Get
        Private Set
            _HoursPerWeek = Value
        End Set
    End Property
    Public Property HoursPerMonth() As Decimal
        Get
            Return _HoursPerMonth
        End Get
        Private Set
            _HoursPerMonth = Value
        End Set
    End Property
    Public Property HoursYear() As Decimal
        Get
            Return _HoursYear
        End Get
        Private Set
            _HoursYear = Value
        End Set
    End Property
    Public Property CalendarExceptions() As String
        Get
            Return _CalendarExceptions
        End Get
        Private Set
            _CalendarExceptions = Value
        End Set
    End Property
#End Region

    Private Function GetDBConn() As IDataClient

        Dim dsn As String = System.Configuration.ConfigurationManager.AppSettings("pm5_DSN")
        Dim p6DBType As PO.Base.DBType = PO.Base.DBType.Oracle
        Dim service As String
        Dim user As String
        Dim passw As String
        Dim initialDB As String
        Dim providerName As String

        If (String.Compare(System.Configuration.ConfigurationManager.AppSettings("pm5_USE_ORACLE"), "true", True) <> 0) Then p6DBType = PO.Base.DBType.SQLServer

        Select Case p6DBType
            Case PO.Base.DBType.Oracle
                PO.Base.DBUtilities.SplitConnString(dsn, False, service, user, passw, initialDB, providerName)
                Return New OracleNETDataClient(service, user, passw, False)
            Case PO.Base.DBType.SQLServer
                If (PO.Base.DBUtilities.UseIntegratedSecurity(dsn)) Then
                    Dim conn As New SQLServerDataClient("", "", "", "", "", False)
                    conn.SetConnectionString(dsn)
                    Return conn
                Else
                    PO.Base.DBUtilities.SplitConnString(dsn, False, service, user, passw, initialDB, providerName)
                    Return New SQLServerDataClient(service, user, passw, initialDB, String.Empty, False)
                End If
            Case Else : Throw New Exception("DB Type not managed!")
        End Select


    End Function
    Private Function EncodeP6CalendarType(p6CalendarTypeCode As String) As P6CalendarType
        Select Case p6CalendarTypeCode.ToUpper()
            Case "CA_BASE"
                Return P6CalendarType.Base
            Case "CA_PROJECT"
                Return P6CalendarType.Project
            Case "CA_RSRC"
                Return P6CalendarType.Resource
            Case Else
                Throw New Exception("Calendar type " & p6CalendarTypeCode & " is not recognized!")
        End Select
    End Function
    Private Function ReadCalendarExceptions() As String
        'DataRow row = null;
        'DataTable t = null;
        't.CreateDataReader
        'OracleType.Blob blob = 
        '' Fetch the BLOB data through OracleDataReader using OracleBlob type
        'OracleBlob blob = oraImgReader.GetOracleBlob(2);

        '' Create a byte array of the size of the Blob obtained
        'Byte[] byteArr = new Byte[blob.Length];

        '' Read blob data into byte array
        'int i = blob.Read(byteArr, 0, System.Convert.ToInt32(blob.Length));

        '' Get the primitive byte data into in-memory data stream
        'MemoryStream memStream = new MemoryStream(byteArr);

        '' Attach the in-memory data stream to the PictureBox
        'picEmpPhoto.Image = Image.FromStream(memStream);

        '' Fit the image to the PictureBox size
        'picEmpPhoto.SizeMode = PictureBoxSizeMode.StretchImage;
        Return String.Empty
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
