Imports System.Web.Services
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports PO.Base
Imports PO.DataClient
Imports PO.Objects
Imports System.Linq
Imports PO.DBInterface

<System.Web.Services.WebService(Namespace:="http://projectobjects.com/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class PM6DBServiceController
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function UpdateProject(ByVal idPM6 As Integer, ByVal source As String) As Boolean
        Dim esito As Boolean
        'devo controllare che nn sia già partito un import
        Dim sp As New SigepPrimaveraApp()
        sp._Source = source
        esito = sp.ExecuteLauncher(idPM6, 1)
        Return esito
    End Function
    ''' <summary>
    ''' Added by Anubrij (POINDIA)
    ''' </summary>
    ''' <param name="idPM6">PM6 project ID</param>
    ''' <param name="sigepID">Sigep project ID</param>
    ''' <returns>Return boolean</returns>
    ''' <remarks>Update the Primavera sigep project association before performing publish as</remarks>
    <WebMethod()>
    Public Function UpdateProjectAssociation(ByVal idPM6 As Integer, ByVal sigepID As String) As Boolean

        'devo controllare che nn sia già partito un import
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Dim sp As New SigepPrimaveraApp()
        sp._Source = sigepID
        Return sp.UpdateProjectAssociation(conn, idPM6, sigepID)

    End Function
    <WebMethod()>
    Public Function ImportProject(ByVal IdSigep As Integer) As Boolean
        Dim esito As Boolean
        'devo controllare che nn sia già partito un update
        Dim sp As New SigepPrimaveraApp
        esito = sp.ExecuteLauncher(IdSigep, 0)
        Return esito
    End Function

    <WebMethod()>
    Public Function DeleteAllIntegrationAPIUsession(ByVal appname As String) As Boolean
        Try
            Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
            conn.ExecuteNonQuery("delete from usession where app_name='" & appname & "' ")
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    '///restituisce il valore del campo note in corrispondenza del progetto con dato idPrimavera , e idlog maggiore
    <WebMethod()>
    Public Function DisplayStatus(ByVal idPM6 As Integer) As String
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Dim concat As String = "+"
        Dim ret As String = String.Empty
        Try
            If String.Compare("true", ConfigurationManager.AppSettings("pm5_USE_ORACLE"), True) = 0 Then
                concat = DBUtilities.StrConcatOperator(DBType.Oracle)
            Else
                concat = DBUtilities.StrConcatOperator(DBType.SQLServer)
            End If
            If Not conn.ExecuteScalar("select status " & concat & " '-' " & concat & " note as status from pm6_log where idpm6 = " & idPM6 & " and id = (select max(id) from pm6_log where idpm6=" & idPM6 & ") ", ret) Then Return "fault"
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return ret
        Catch ex As Exception
            Return "fault"
        End Try
    End Function

    <WebMethod()>
    Public Function ReplacePM6StartProject(ByVal idPM6 As Integer, ByVal user As String) As Boolean

        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
        'Dim dbtype As Boolean = ConfigurationManager.AppSettings("pm5_USE_ORACLE")
        Dim stmt As String = "SIGEP5_REPLACEPROJECT"
        Dim params As IDbDataParameter() = Nothing
        If String.Compare("true", ConfigurationManager.AppSettings("pm5_USE_ORACLE"), True) = 0 Then
            params = New OracleParameter() {
                New OracleParameter("nuovo", OracleType.VarChar) _
                , New OracleParameter("j", OracleType.VarChar)
                }
        Else
            params = New SqlParameter() {
                New SqlParameter("@new", SqlDbType.VarChar) _
                , New SqlParameter("@j", SqlDbType.VarChar)
            }
        End If

        Try
            params(0).Value = idPM6
            params(1).Value = user
            conn.ExecuteNonQuery(stmt, params)
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return True
        Catch ex As Exception
            Return False
        Finally
        End Try

    End Function
    '<WebMethod()> _
    'Public Function ReplacePM6StartProject2(ByVal idPM6 As Integer, ByVal user As String) As Boolean
    '    Dim userCode As Integer
    '    userCode = DbUtility.QueryValue("select user_id from users where user_name='" & user & "'", Primavera6.datasource.pm5)
    '    Dim esito As Boolean
    '    Dim user_data_id As Integer = DbUtility.QueryValue("select  user_data_id  from userdata where user_id= " & userCode & " and topic_name='pm_settings'", Primavera6.datasource.pm5)
    '    ' Dim user_data As String = DbUtility.QueryValue("select  sigep5_blobtoclob(user_data)  from  userdata where user_data_id= " & user_data_id, Primavera6.datasource.pm5)
    '    Dim user_data As String = LoadTextBlob(userCode)
    '    Dim _connectionstring As String = System.Configuration.ConfigurationManager.AppSettings("pm5_DSN")
    '    Dim conn As New OleDbConnection(_connectionstring)
    '    Dim cmd As New OleDbCommand()
    '    cmd.Connection = conn
    '    Try
    '        If conn.State <> ConnectionState.Open Then
    '            conn.Open()
    '        End If
    '        Dim clob As IO.FileStream
    '        cmd.CommandText = "select  sigep5_blobtoclob(user_data)  from  userdata where user_data_id= " & user_data_id
    '        clob = cmd.ExecuteScalar()
    '        esito = True
    '    Catch ex As Exception
    '        esito = False
    '    Finally
    '    End Try
    '    Return esito
    'End Function

    <WebMethod()>
    Public Function GetAssociateProject(ByVal idSigep As Integer) As Integer
        Dim ret As Integer = 0
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Try
            If Not conn.ExecuteScalar("select PM6_PROJID from prog_t056 where c_prog = " & idSigep, ret) Then Return 0
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return ret
        Catch ex As Exception
            return 0
        End Try
    End Function

#Region "Procedure to check if a P6 project is assigned to SIGEP project"
    ''' <summary>
    '''checks if a P6 schedule is already linked to a SIGEP project different from the one the user is publishing to"
    ''' Added By Indramani
    ''' </summary>
    ''' <param name="idPM">PM6 project ID</param>
    ''' <returns>Associated sigep project id </returns>
    ''' <remarks></remarks>
    <WebMethod()>
    Public Function GetPM6AssociateProject(ByVal idPM As Integer) As Integer
        Dim ret As Integer = 0
        Try
            Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
            If Not conn.ExecuteScalar("select c_prog from prog_t056 where pm6_projid = " & idPM, ret) Then Return 0
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return ret
        Catch ex As Exception
            ret = 0
        End Try
        Return ret
    End Function
    <WebMethod()>
    Public Function GetPM6AssociateProjectExceptCurrentProject(ByVal idPM As Integer) As Integer
        Dim ret As Integer = 0
        Try
            Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
            If Not conn.ExecuteScalar("select c_prog from prog_t056 where pm6_projid not IN( " & idPM & ")", ret) Then Return 0
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return ret
        Catch ex As Exception
            ret = 0
        End Try
        Return ret
    End Function
    <WebMethod()>
    Public Function GetSigepProjectInfo(ByVal idSigep As Integer) As Byte()
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Dim sqlStm As String = "SELECT S_prog_Nom,S_prog  from prog_t056 where c_prog = " & idSigep
        Dim data As DataSet = conn.ExecuteQuery(sqlStm)
        Return ZDataSet.ZipDataSet(data)
    End Function
#End Region

    <WebMethod()>
    Public Function GetSigepProjectStatus(ByVal idPM As Integer) As Integer
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Dim status As Integer = -1
        Try
            If IBaseDataMgr.GetDBType(DataSource.PO) = DBType.Oracle Then
                If Not conn.ExecuteScalar("select f_sta from prog_t056 where c_prog in  ( select nvl(c_prog,0)  from prog_t056 where   pm6_projid = " & idPM & " and rownum = 1) ", status) Then Return -1
            ElseIf IBaseDataMgr.GetDBType(DataSource.PO) = DBType.SQLServer Then
                If Not conn.ExecuteScalar("select f_sta from prog_t056 where c_prog in  ( select top 1 isnull(c_prog,0)  from prog_t056 where   pm6_projid = " & idPM & ")", status) Then Return -1
            End If
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return status
        Catch ex As Exception
            return -1
        End Try
    End Function

    <WebMethod()>
    Public Function GetUserPrimaveraPwd(userId As String) As String
        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Dim passw As String = Nothing
        Try
            If Not conn.ExecuteScalar("select S_PASSW from VW_PRIMAVERAUSERINFO where UPPER(C_USER) = " & PO.Base.DBUtilities.QS(userId.Trim().ToUpperInvariant()), passw) Then Return Nothing
            If Not conn.LastException Is Nothing Then Throw conn.LastException
            Return passw
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Added By Anubrij
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>User to load the project search popup in add-in</remarks>
#Region "Project search related methods"
    <WebMethod()> Public Function ProjStatusList() As Byte() ' Dictionary(Of Integer, String)
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetProjStatusList()
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function ProjLocationsList(ByVal companyId As Integer) As Byte() 'Dictionary(Of Integer, String)
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetProjLocationsList(companyId)
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function GetDefaultP6calendarList() As Byte() ' Dictionary(Of String, String)
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetDefaultP6calendarList()
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function ProjCustomers(ByVal companyId As Integer) As Byte() ' Dictionary(Of String, String)
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetProjCustomersList(companyId)
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function Projects(ByVal companyId As Integer, ByVal userId As String, ByVal descr As String, ByVal status As Integer, ByVal location As Integer, ByVal cust As String) As Byte()
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetProjects(companyId, userId, descr, status, location, cust)
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function ProjectSanpshots(ByVal projectID As Integer) As Byte()
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetProjectSnapshots(projectID)
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function CompanyList(userId As String) As Byte()
        Dim obj As New PODBSearch()
        Dim data As DataSet = obj.GetCompanyList(userId)
        Return ZDataSet.ZipDataSet(data)
    End Function
    <WebMethod()> Public Function CompanyListExt(userId As String, ByRef errMsg As String) As Byte()

        errMsg = String.Empty
        Try

            Dim obj As New PODBSearch()
            Dim data As DataSet = obj.GetCompanyList(userId)
            If Not obj.LastException Is Nothing Then Throw obj.LastException
            Return ZDataSet.ZipDataSet(data)
        Catch e As Exception
            errMsg = "Error: " & e.Message & IIf(e.InnerException Is Nothing, "", vbCrLf & "Inner exception: " & e.InnerException.Message) & vbCrLf & "Call stack: " & e.StackTrace
            Return Nothing
        End Try
    End Function
#End Region

End Class

' Class PODBSearch, used to search within projects and to provide some values lists to be used as search parameters 
Public Class PODBSearch
    Inherits PO.DBInterface.IBaseDataMgr

    Public Sub New()
        MyBase.New(PO.DBInterface.DataSource.PO)
    End Sub
    Public Function GetProjStatusList() As DataSet 'As Dictionary(Of Integer, String)
        Dim result As New DataSet() 'Dictionary(Of Integer, String)
        result.Tables.Add()
        result.Tables(0).Columns.Add("STATUSID", GetType(Integer))
        result.Tables(0).Columns.Add("STATUSNAME", GetType(String))
        result.Tables(0).LoadDataRow(New Object() {0, "Draft"}, True)
        result.Tables(0).LoadDataRow(New Object() {1, "Approved"}, True)
        result.Tables(0).LoadDataRow(New Object() {2, "Open"}, True)
        Return result
    End Function
    Public Function GetProjCustomersList(ByVal companyId As Integer) As DataSet 'As Dictionary(Of String, String)
        Try

            Dim data As DataSet = _Conn.ExecuteQuery("select C_CLI, S_CLI from cli_t023 where C_AZD = " & companyId.ToString() & " order by S_CLI")
            Return data
        Catch ex As Exception
            Return Nothing
        End Try

    End Function
    Public Function GetProjLocationsList(ByVal companyId As Integer) As DataSet 'Dictionary(Of Integer, String)

        Try
            Dim result As DataSet = _Conn.ExecuteQuery("select C_LOCATION, S_LOCATION from TAB_LOCATIONS where C_AZD = " & companyId.ToString() & " order by S_LOCATION")
            Return result

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function GetCompanyList(ByVal userId As String)
        Try

            Dim sysAdmin As Boolean = False
            Dim userCode As Long = 0

            CheckSysAdminUser(userId, userCode, sysAdmin)

            Dim stmt As String = "select c_azd, s_azd from azd_t203 azd "
            If Not sysAdmin Then
                stmt &= " where azd.c_azd in (" &
                        " select C_AZD from ASCN_GRP_AZD_T204 where C_GRP in (" &
                        "       select C_GRP from ASCN_GRP_UTEN_T020 where C_UTEN = " & userCode.ToString() & ") " &
                        "   union " &
                        " select C_AZD from ASCN_UTEN_AZD_T205 " &
                        " where C_UTEN = " & userCode.ToString() &
                        ") "
            End If
            stmt &= " order by s_azd "
            Dim result As DataSet = _Conn.ExecuteQuery(stmt)
            Return result

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetProjectSnapshots(ByVal projectID As Integer)
        Try

            Dim result As DataSet = _Conn.ExecuteQuery("SELECT C_PROG_SNAPSHOT, CREATION_DATE FROM TAB_snapshot_version WHERE C_PROG = " & projectID & " AND IS_VALID = 1 ORDER BY CREATION_DATE DESC")
            Return result

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Function CheckSysAdminUser(userId As String, ByRef userCode As Long, ByRef sysAdmin As Boolean) As Boolean

        Dim dbOwner As String = ConfigurationManager.AppSettings("DBOWNER")
        Dim grants As String = Nothing

        sysAdmin = False
        userCode = 0

        Try
            _Conn.ExecuteScalar("select c_uten from UTEN_T017 where c_user=" & DBUtilities.QS(userId.Trim().ToUpperInvariant()), userCode)
        Catch ex As Exception
            userCode = 0
        End Try

        Try
            _Conn.ExecuteScalar("SELECT " & dbOwner & ".FUNZL_UTEN_F007(" & userCode & ", '0', -1, '0') as sOutGrants from DUAL", grants)
            sysAdmin = (grants.Substring(0, 1) = "1")
        Catch
            sysAdmin = False
        End Try

        Return sysAdmin

    End Function
    Public Function GetProjects(ByVal companyId As Integer, ByVal userId As String, ByVal descr As String, ByVal status As Integer, ByVal location As Integer, ByVal cust As String) As DataSet

        Dim p6ProjName As String
        Dim p6Conn As IDataClient = GetDBConnection(PO.DBInterface.DataSource.Primavera)
        Dim p6ProjList As List(Of String)
        Dim userCode As Long = 0
        Dim sysAdmin As Boolean = False
        Dim poDBType As DBType = GetDBType(PO.DBInterface.DataSource.PO)
        Dim p6DBType As DBType = GetDBType(PO.DBInterface.DataSource.Primavera)
        Dim dbOwner As String = ConfigurationManager.AppSettings("DBOWNER")
        Dim strCatPO As String = IBaseDataMgr.Concat(poDBType)
        Dim isnullPO As String = IBaseDataMgr.IsNull(poDBType)
        Dim strCatP6 As String = IBaseDataMgr.Concat(p6DBType)
        Dim isnullP6 As String = IBaseDataMgr.IsNull(p6DBType)
        Dim stmt As String = " select pr.C_PROG, " & isnullPO & "(PM6_PROJID, 0) P6_PROJ_ID, S_PROG as Title, S_PROG_AZD_NOM as Code, loc.S_LOCATION as Location, cust.S_CLI as Customer" &
                                " , case pr.F_STA " &
                                " when 0 then 'Draft' " &
                                " when 1 then 'Approved' " &
                                " when 2 then 'Open' " &
                                " end as Status " &
                                ",  NULL as " & IBaseDataMgr.LeftColumnNameDelimiter(poDBType) & "P6 SCHEDULE" & IBaseDataMgr.RightColumnNameDelimiter(poDBType) & " " &
                                " from prog_t056 pr " &
                                " left outer join TAB_LOCATIONS loc on ltrim(rtrim(cast(loc.C_LOCATION as varchar(20)))) = pr.S_LOC and loc.C_AZD = pr.C_AZD " &
                                " left outer join CLI_T023 cust on cust.C_CLI = pr.C_CLI and cust.C_AZD = pr.C_AZD " &
                                " where F_MODELLO <> 1 and pr.c_azd = " & companyId.ToString()

        Try

            If Not String.IsNullOrEmpty(descr) Then stmt &= " and upper(pr.S_PROG" & strCatPO & "'#'" & strCatPO & "pr.S_PROG_AZD_NOM) like " & DBUtilities.QS("%" & descr.Trim().ToUpper() & "%")
            If status >= 0 Then stmt &= " and pr.F_STA = " & status.ToString()
            If location > 0 Then stmt &= " and pr.S_LOC = " & location.ToString()
            If Not String.IsNullOrEmpty(cust) Then stmt &= " and pr.C_CLI = " & DBUtilities.QS(cust)

            CheckSysAdminUser(userId, userCode, sysAdmin)

            'sysAdmin = True
            If Not sysAdmin Then
                stmt &= " and pr.C_PROG in (" &
                        " select C_PROG from ASCN_GRP_PROG_T070 where C_GRP in (" &
                        "       select c_grp from ASCN_GRP_UTEN_T020 inner join UTEN_T017 on ASCN_GRP_UTEN_T020.C_UTEN = UTEN_T017.C_UTEN " &
                        "       where UTEN_T017.C_UTEN = " & userCode.ToString() & ") " &
                        "   union " &
                        " select C_PROG from ASCN_UTEN_PROG_T071 " &
                        " where C_UTEN = " & userCode.ToString() &
                        ")"
            End If

            Dim result As DataSet = _Conn.ExecuteQuery(stmt)
            If Not _Conn.LastException Is Nothing Then Throw _Conn.LastException

            p6ProjList = (From item In result.Tables(0).AsEnumerable().Select(Function(x) x("P6_PROJ_ID").ToString()).Distinct()).ToList()

            stmt = " select  pr.proj_id, pr.proj_short_name, wbs.wbs_name" & strCatP6 & "'-EPS='" & strCatP6 & isnullP6 & "(parenteps.wbs_short_name, '-') proj_ext_name " &
            " from PROJECT pr " &
            " inner join (select * from projwbs where proj_node_flag = 'Y') wbs on pr.proj_id = wbs.proj_id " &
            " left join projwbs parenteps on wbs.PARENT_WBS_ID = parenteps.WBS_ID " &
            " where pr.PROJ_ID in (" & String.Join(",", p6ProjList.ToArray()) & ") and pr.project_flag = 'Y'"

            Dim p6Projects As DataSet = p6Conn.ExecuteQuery(stmt)
            If Not p6Conn.LastException Is Nothing Then Throw p6Conn.LastException

            Dim p6Dictionary As Dictionary(Of Integer, String) =
            (From item In p6Projects.Tables(0).AsEnumerable() Select New With
            {
                Key .Key = Convert.ToInt32(item("proj_id")),
                Key .Val = DirectCast(item("proj_ext_name"), String)
            }).Distinct().AsEnumerable().ToDictionary(Function(k) k.Key, Function(v) v.Val)

            For Each row As DataRow In result.Tables(0).Rows
                If Convert.ToInt32(row("P6_PROJ_ID")) = 0 Then Continue For
                If Not p6Dictionary.TryGetValue(Convert.ToInt32(row("P6_PROJ_ID")), p6ProjName) Then Continue For
                row("P6 SCHEDULE") = p6ProjName
            Next

            ' dropping P6 project id column...
            result.Tables(0).Columns.Remove("P6_PROJ_ID")

            Return result

        Catch ex As Exception

            'Log.WriteLog("GetProjects error: " & ex.Message)
            'Log.WriteLog("Statement: " & stmt)
            Return Nothing

        End Try

    End Function
    ''Get DefaultP6calendar 
    Public Function GetDefaultP6calendarList() As DataSet
        Try

            Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
            Return conn.ExecuteQuery("SELECT CLNDR_ID, CLNDR_NAME FROM CALENDAR order by CLNDR_NAME")

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class