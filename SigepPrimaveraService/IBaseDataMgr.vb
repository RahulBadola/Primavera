Imports System.Data
Imports PO.DataClient
Imports PO.Base
Imports System.Security

Namespace PO.DBInterface

    Public Enum DataSource
        PO
        Primavera
    End Enum
    ''' <summary>
    ''' Added by Anubrij
    ''' </summary>
    ''' <remarks>Database intraction class used in POsearch class</remarks>
    Public MustInherit Class IBaseDataMgr

        Protected _LastException As Exception
        Protected _Conn As IDataClient
        Protected _FormatProvider As IFormatProvider = System.Globalization.CultureInfo.InvariantCulture

        Friend Sub New(ByVal source As DataSource)

            _Conn = GetDBConnection(source)

        End Sub
        Public ReadOnly Property LastException() As Exception
            Get
                Return _LastException
            End Get
        End Property

        'Public Shared Function GetDBType() As DBConnType
        '    Dim cp As Boolean = System.Configuration.ConfigurationManager.AppSettings("sigep_USE_ORACLE")
        '    If cp Then Return DBConnType.Oracle
        '    Return DBConnType.SQLServer
        '    'Dim cpVal As Integer
        '    'If (Not Integer.TryParse(cp, cpVal)) Then Throw New Exception("DB Type '" & cp & "' is not defined!")
        '    'If (Not [Enum].IsDefined(GetType(DBConnType), cpVal)) Then Throw New Exception("DB Type value '" & cp & "' is not defined!")
        '    'Return cpVal
        'End Function
        Public Shared Function GetDBType(dataSourceId As DataSource) As PO.Base.DBType

            Dim cfgkey As String = String.Empty
            Select Case dataSourceId
                Case DataSource.PO : cfgkey = "sigep_USE_ORACLE"
                Case DataSource.Primavera : cfgkey = "pm5_USE_ORACLE"
            End Select
            Dim cfgvalue As String = ConfigurationManager.AppSettings(cfgkey)

            If (String.Compare(cfgvalue, "true", True) = 0) Then : Return Base.DBType.Oracle
            Else : Return Base.DBType.SQLServer
            End If

        End Function
        Public Shared Function IsNull(DBType As PO.Base.DBType) As String
            Select Case DBType
                Case Base.DBType.Oracle : Return "nvl"
                Case Base.DBType.SQLServer : Return "isnull"
                Case Else : Throw New Exception("DB type " & DBType.ToString() & " not managed")
            End Select
        End Function
        Public Shared Function Concat(DBType As PO.Base.DBType) As String
            Select Case DBType
                Case Base.DBType.Oracle : Return "||"
                Case Base.DBType.SQLServer : Return "+"
                Case Else : Throw New Exception("DB type " & DBType.ToString() & " not managed")
            End Select
        End Function
        Public Shared Function LeftColumnNameDelimiter(DBType As PO.Base.DBType) As String
            Select Case DBType
                Case Base.DBType.Oracle : Return """"
                Case Base.DBType.SQLServer : Return "["
                Case Else : Throw New Exception("DB type " & DBType.ToString() & " not managed")
            End Select
        End Function
        Public Shared Function RightColumnNameDelimiter(DBType As PO.Base.DBType) As String
            Select Case DBType
                Case Base.DBType.Oracle : Return """"
                Case Base.DBType.SQLServer : Return "]"
                Case Else : Throw New Exception("DB type " & DBType.ToString() & " not managed")
            End Select
        End Function
        Public Shared Function DatetimeToDBFormula(val As DateTime?, time As String, connType As IDataClient.eConnectionType) As String

            If Not val.HasValue Then Return "NULL"
            Return DatetimeToDBFormula(val.Value, time, connType)

        End Function

        Public Shared Function DatetimeToDBFormula(val As DateTime, time As String, connType As IDataClient.eConnectionType) As String
            If IsDBNull(val) Then Return "NULL"
            If DateTime.Compare(val, DateTime.MinValue) = 0 Then Return "NULL"

            Dim strDateTime = val.ToString("yyyyMMdd") & " " & time
            Select Case connType
                Case IDataClient.eConnectionType.Oracle, IDataClient.eConnectionType.OracleNET
                    Dim format As String = "YYYYMMDD"
                    If Not String.IsNullOrEmpty(time) Then format &= " HH24:MI"
                    Return "to_date('" & strDateTime & "', '" & format & "')"
                Case IDataClient.eConnectionType.SqlServer
                    Return "'" & strDateTime & "'"
                Case Else : Throw New Exception("Connection type " & connType.ToString() & " not managed")
            End Select

        End Function

        <Permissions.PermissionSet(Permissions.SecurityAction.Assert, Unrestricted:=True)>
        Friend Shared Function GetDBConnection(ByVal source As DataSource) As IDataClient
            Return GetDBConnection(source, False)
        End Function
        <Permissions.PermissionSet(Permissions.SecurityAction.Assert, Unrestricted:=True)>
        Friend Shared Function GetDBConnection(ByVal source As DataSource, permanent As Boolean) As IDataClient

            Dim dsn As String
            Dim service As String
            Dim user As String
            Dim passw As String
            Dim initialDB As String
            Dim providerName As String

            ' TODO: differentiate
            Select Case source
                Case DataSource.PO
                    dsn = PwCrypt.DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("sigep_DSN"))
                Case DataSource.Primavera
                    dsn = System.Configuration.ConfigurationManager.AppSettings("pm5_DSN") 'PwCrypt.DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("pm5_DSN"))
            End Select


            Select Case GetDBType(source)
                Case Base.DBType.Oracle
                    DBUtilities.SplitConnString(dsn, False, service, user, passw, initialDB, providerName)
                    Return New OracleNETDataClient(service, user, passw, False)
                Case Base.DBType.SQLServer
                    If (DBUtilities.UseIntegratedSecurity(dsn)) Then
                        Dim conn As New SQLServerDataClient("", "", "", "", "", False)
                        conn.SetConnectionString(dsn)
                        Return conn
                    Else
                        DBUtilities.SplitConnString(dsn, False, service, user, passw, initialDB, providerName)
                        Return New SQLServerDataClient(service, user, passw, initialDB, String.Empty, False)
                    End If
                    'Dim conn As IDataClient = New OleDBDataClient("", "", "", "", "", False)
                    'conn.SetConnectionString(dsn)
                    'Return conn
                Case Else : Throw New Exception("DB Type not managed!")
            End Select

        End Function
        Friend Shared Function TestDBConnection(ByVal source As DataSource) As Exception

            Dim conn As IDataClient = GetDBConnection(source)
            Dim curTime As DateTime = DateTime.Now
            Dim stmt As String = String.Empty
            Try

                Select Case GetDBType(source)
                    Case Base.DBType.Oracle : stmt = "select SYSDATE from DUAL"
                    Case Base.DBType.SQLServer : stmt = "select GETDATE()"
                    Case Else : Throw New Exception("DB Type not managed!")
                End Select

                Dim data As DataSet = conn.ExecuteQuery(stmt)
                If Not conn.LastException Is Nothing Then Throw conn.LastException
                curTime = DirectCast(data.Tables(0).Rows(0)(0), DateTime)
                Return Nothing

            Catch ex As Exception
                Return ex
            Finally

                conn.Dispose()
                conn = Nothing

            End Try

        End Function

    End Class

End Namespace
