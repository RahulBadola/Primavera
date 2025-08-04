Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports PO.DataClient
Imports PO.DBInterface

Public Class SigepPrimaveraApp

    Private _IdLog As Integer
    Private _InitLabel As String = "Initializing Tasks"
    Private _ImportWBSLabel As String = "Import WBS Structure"
    Private _InsertAllWBSLabel As String = "Create WBS Tasks"
    'Private _ImportMilestoneLabel As String = "Import Milestones"
    Public _Source As String = String.Empty
    'Public _DmlSigepCmd As New OleDbCommand
    'Public _LogSigepCmd As New OleDbCommand
    'Public _CnSigep As New OleDbConnection
    'Public _TransSigep As OleDbTransaction = Nothing
    Public _AbortMsg As String = String.Empty

    Private _NoTransactionDBConn As IDataClient
    Public _DMLDBConn As IDataClient

    Public Sub New()

        _NoTransactionDBConn = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO, True)
        _DMLDBConn = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO, True)

    End Sub

    ''' <summary>
    ''' Set the new log ID to be used during the import session
    ''' </summary>
    ''' <returns></returns>
    Private Function SetNewLogID() As Integer

        Dim conn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.PO)
        Select Case IBaseDataMgr.GetDBType(DataSource.PO)
            Case PO.Base.DBType.Oracle : conn.ExecuteScalar("select PM6_LOG_SEQ.nextval from dual", _IdLog)
            Case PO.Base.DBType.SQLServer : conn.ExecuteScalar("select next value for PM6_LOG_SEQ", _IdLog)
            Case Else : Throw New Exception("DB type not managed")
        End Select

        Return _IdLog

    End Function
    ''' <summary>
    ''' Removes items older than what specified in the configuration from the log table. 
    ''' </summary>
    Private Sub CleanOldLogItems()

        Try
            Dim timeout As Integer = CInt(ConfigurationManager.AppSettings("pendingOperations"))

            Select Case IBaseDataMgr.GetDBType(DataSource.PO)
                Case PO.Base.DBType.Oracle : _NoTransactionDBConn.ExecuteNonQuery("delete from  pm6_log where TO_NUMBER(sysdate-datetime) * 24*60  > " & timeout & " and status=1")
                Case PO.Base.DBType.SQLServer : _NoTransactionDBConn.ExecuteNonQuery("delete from  pm6_log where cast(getdate()-datetime as integer) * 24*60  > " & timeout & " and status=1")
                Case Else : Throw New Exception("DB type not managed")
            End Select

        Catch
        End Try

    End Sub
    ''' <summary>
    ''' Write a log message in the log table, and in a log file -if specified
    ''' </summary>
    ''' <param name="text"></param>
    ''' <param name="idsigep"></param>
    ''' <param name="idprimavera"></param>
    Public Sub WriteLog(ByVal text As String, Optional ByVal idsigep As Integer = 0, Optional ByVal idprimavera As Integer = 0)
        Try
            If idsigep > 0 AndAlso idprimavera > 0 Then
                Dim Sql As String = " update pm6_log set note = '" & text & "' where id = " & Me._IdLog & " and  IDSIGEP = " & idsigep & " and  IDPM6 = " & idprimavera
                _NoTransactionDBConn.ExecuteNonQuery(Sql)
            End If
            Dim logpath As String = ConfigurationManager.AppSettings("logpath")
            If String.IsNullOrEmpty(logpath) Then Return
            If Not System.IO.File.Exists(logpath) Then Return
            FileIO.FileSystem.WriteAllText(logpath, Now.ToString("yyyyMMdd HH:mm:ss ") & text & vbCrLf, True)
        Catch ex As Exception
        End Try
    End Sub
    ''' <summary>
    ''' Log table initialization. If it fails the process is interrupted because it would not be possible to perform some checks
    ''' </summary>
    ''' <param name="idSigep"></param>
    ''' <param name="idPrimavera"></param>
    ''' <param name="direction"></param>
    ''' <returns></returns>
    Public Function InitLog(idSigep As Integer, idPrimavera As Integer, direction As Integer) As Boolean

        Dim dbsysdate As String = "sysdate"
        If IBaseDataMgr.GetDBType(DataSource.PO) = PO.Base.DBType.SQLServer Then dbsysdate = "getdate()"
        Dim stmt As String = " insert into pm6_log (ID, DATETIME, DIRECTION, IDSIGEP, IDPM6, OPERATION, STATUS, NOTE) values    (" & _IdLog & ", " & dbsysdate & ", " & direction & " , " & idSigep & " , " & idPrimavera & ",'update',1,'')"
        _NoTransactionDBConn.ExecuteNonQuery(stmt)
        If Not _NoTransactionDBConn.LastException Is Nothing Then Return False
        Return True

    End Function
    Public Sub CloseLog(idSigep As Integer, idPrimavera As Integer, message As String, status As Integer)

        Try
            If (message.Length >= 150) Then message = Left(message.Replace("'", ""), 150)
            _NoTransactionDBConn.ExecuteNonQuery("update pm6_log set note = " & PO.Base.DBUtilities.QS(message) & ", status = " & status.ToString() & " where id = " & _IdLog & " and IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera)
        Catch
        End Try

    End Sub
    'Public Sub CheckAbort(ByVal idSigep As Integer, ByVal idPM5 As Integer)

    '    Dim res As Integer = 0
    '    Dim Sql As String = "select status from pm6_log where id = " & _IdLog & " and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPM5
    '    _NoTransactionDBConn.ExecuteScalar(Sql, res)
    '    If res = 3 Then
    '        _AbortMsg = "Aborted"
    '        Throw New Exception("Abort")
    '    End If

    'End Sub
    ''' <summary>
    '''  Check if there's a concurrent update running with the current one
    ''' </summary>
    ''' <param name="idSigep"></param>
    ''' <param name="idPM5"></param>
    ''' <param name="dir"></param>
    ''' <returns></returns>
    Public Function ConcurrentUpdate(ByVal idSigep As Integer, ByVal idPM5 As Integer, ByVal dir As Integer) As Boolean

        '07 gennaio 09------------------
        Dim counter As Integer = 0
        Dim sql As String = String.Empty
        _NoTransactionDBConn.ExecuteScalar("select count(*) from pm6_log where status = 1 and IDSIGEP = " & idSigep & " and  IDPM6 = " & idPM5 & " and id != " & _IdLog.ToString(), counter)
        If counter > 0 Then
            sql = " update pm6_log set note = 'Aborted For Simultaneous Upgrade', status = 4 where id= " & _IdLog
            _NoTransactionDBConn.ExecuteNonQuery(sql)
            Return True
        End If
        Return False
        '-----------------------------------

    End Function

    Public Function ExecuteLauncher(ByVal projId As String, ByVal direction As Integer) As Boolean

        Dim result As Boolean
        Dim idSigep As Integer, idPrimavera As Integer
        Dim sql As String = String.Empty
        Dim wbsP6Items As Integer = 0

        ' set the log ID for the session
        SetNewLogID()

        ' log cleanup
        CleanOldLogItems()

        WriteLog("Starting process")

        Try
            If direction = 1 Then
                '/////////////// 1: se stiamo aggiornando SIGEP da menu di PM5//////////////////////////////////////
                'il parametro di ingresso è  il CODICE PROGETTO PRIMAVERA appena modificato
                'UPDATE PROJECT
                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                If String.Compare(_Source, "sigep", True) = 0 Then direction = 0
                idSigep = 0
                idPrimavera = projId
                _NoTransactionDBConn.ExecuteScalar("select c_prog  from prog_t056 where pm6_projid = " & idPrimavera, idSigep)
            ElseIf direction = 0 Then
                '///// 0: adesso uso questa direzione per aggiornare SIGEP (importando un progetto PM6 in sigep) da SIGEP
                '// il parametro di ingresso è  il CODICE PROGETTO appena creato in SIGEP col nome del progetto primavera6 
                'IMPORT PROJECT
                '///////////////////////////////////
                idPrimavera = 0
                idSigep = CInt(projId)
                _NoTransactionDBConn.ExecuteScalar("select pm6_projid  from prog_t056 where c_prog  = " & projId, idPrimavera)
            End If

            WriteLog("Processing project : P6:" & idPrimavera & " - SIGEP:" & idSigep)

            'prima di iniziare verifico che nn sia già in corso un update di tali progetto
            '07 gennaio 09------------------
            If ConcurrentUpdate(idSigep, idPrimavera, direction) Then Return False
            '-----------------------------------

            ' L'inizio dell'update viene identificato dal flag status=1  sulla riga corispondente all IdSigep e idprimavera
            If Not InitLog(idSigep, idPrimavera, direction) Then Return False

            WriteLog(_InitLabel)

            '' start the transaction
            _DMLDBConn.BeginTransaction(IsolationLevel.ReadCommitted)

            ' initializes the interface with P6
            Dim p5 As New Primavera6(idPrimavera, idSigep, Me)
            If Not p5.LastException Is Nothing Then Throw p5.LastException

            '----------------------------------- S T O P ------------------------------------------------------------------------------------
            ' se faccio un aggiornamento NON IMPORTO LA WBS'
            'tranne se nn ne esiste nessuna ossia
            _DMLDBConn.ExecuteScalar("select nvl(max(p_lvl),0)  from DEFN_LVL_STRU_T014   a where a.C_PROG= " & idSigep & "  and c_stru='WBSP6'", wbsP6Items)
            If wbsP6Items = 0 Then
                result = p5.ImportWBSInSigep()
                WriteLog(_ImportWBSLabel, idSigep, idPrimavera)
                If Not p5.LastException Is Nothing OrElse Not result Then Throw New Exception("ImportWBSInSigep error: " & p5.LastException.Message)
            End If

            '-------------------------------------------------------------------

            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////
            '' ''NON ELIMINO PIU I TASK MA LI AGGIORNO; INSERISCO O CANCELLO' 9 sett 09
            ' ''printResponse = p5.deleteTaskInSigep()
            ' ''WriteLog("Delete Task In Sigep:" & printResponse)
            ' ''If printResponse = False Then
            ' ''    Throw New Exception("deleteTaskInSigep" & " " & AbortMsg)
            ' ''End If
            ' ''checkAbort(idSigep, idprimavera)
            '' ''07 gennaio 09------------------
            ' ''If exists(idSigep, idprimavera, dir) Then Exit Function
            '' ''-----------------------------------
            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////

            ' task deletion reintroduced as per request - 11/11/2013 - MP
            ' eventually check rows 890 to 903 in procedure insertAllWbsInSigep, class Primavera6
            ' to manage appropriately data in tables prog_track_t091, cost_rev_asso_task, po_task_pbs_codes, msp_task_assignment

            result = p5.DeleteTaskInSigep()
            WriteLog("Delete Task In Sigep:" & result)
            'If Not result Then Throw New Exception("DeleteTaskInSigep error")
            'CheckAbort(idSigep, idPrimavera)
            If Not p5.LastException Is Nothing Then Throw New Exception("DeleteTaskInSigep error: " & p5.LastException.Message)
            If ConcurrentUpdate(idSigep, idPrimavera, direction) Then Return False

            result = p5.InsertAllWbsInSigep()
            WriteLog(_InsertAllWBSLabel, idSigep, idPrimavera)
            If Not result Then Throw New Exception("InsertAllWbsInSigep error")
            'CheckAbort(idSigep, idPrimavera)
            '07 gennaio 09------------------
            If Not p5.LastException Is Nothing Then Throw New Exception("InsertAllWbsInSigep error: " & p5.LastException.Message)
            If ConcurrentUpdate(idSigep, idPrimavera, direction) Then Return False
            '-----------------------------------

            result = UpdatePublishTimestamp(idSigep)
            If Not result Then Throw New Exception("UpdatePublishTimestamp error")

            ' update project association - just to refresh the external time-now in SIGEP
            result = UpdateProjectAssociation(_DMLDBConn, idPrimavera, idSigep)
            If Not result Then Throw New Exception("UpdateProjectAssociation error")

            ' La fine dell'update viene identificata dal flag status=2  sulla riga corispondente all IdSigep e idPrimavera
            CloseLog(idSigep, idPrimavera, "OK", 2)
            WriteLog("Processing end" & vbCrLf & vbCrLf & vbCrLf)

            _DMLDBConn.CommitTransaction()
            Return True

        Catch ex As Exception

            If _DMLDBConn.TransactionActive Then _DMLDBConn.RollbackTransaction()

            CloseLog(idSigep, idPrimavera, ex.Message, 3)
            WriteLog(ex.Message)
            Return False

        Finally
            _DMLDBConn.ClearPoolAndDispose()
            _NoTransactionDBConn.ClearPoolAndDispose()

        End Try

    End Function

    'Public Function ExecuteLauncherOLD(ByVal projId As String, ByVal direction As Integer) As Boolean

    '    'ad ogni esecuzione verifica che non ci siano thread pendenti con status=1 (ossia nn ancora conclusi) ma datetime vecchio in base al settaggio da web.config
    '    Try
    '        Dim timeout As Integer = CInt(ConfigurationManager.AppSettings("pendingOperations"))
    '        DbUtility.QueryExec("delete from  pm6_log where TO_NUMBER(sysdate-datetime) * 24*60  > " & timeout & " and status=1", Primavera6.datasource.sigep)
    '    Catch
    '    End Try
    '    'fine
    '    Dim printResponse As Boolean
    '    WriteLog("Processing start at :")

    '    '/////////////// 1: se stiamo aggiornando SIGEP da menu di PM5//////////////////////////////////////
    '    'il parametro di ingresso è  il CODICE PROGETTO PRIMAVERA appena modificato
    '    'UPDATE PROJECT
    '    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////

    '    Dim idSigep As Integer, idPrimavera As Integer
    '    Dim sql As String = String.Empty

    '    Try
    '        If direction = 1 Then
    '            Dim dir As Integer = direction
    '            If String.Compare(_Source, "sigep", True) = 0 Then dir = 0

    '            idPrimavera = projId
    '            idSigep = DbUtility.QueryTable("select c_prog  from prog_t056 where pm6_projid = " & idPrimavera, Primavera6.datasource.sigep).Rows(0)(0)
    '            WriteLog("Processing project : " & idPrimavera & "-" & idPrimavera)

    '            'prima di iniziare verifico che nn sia già in corso un update di tali progetto
    '            '07 gennaio 09------------------
    '            If ConcurrentUpdate(idSigep, idPrimavera, dir) Then Return False
    '            '-----------------------------------

    '            ' L'inizio dell'update viene identificato dal flag status=1  sulla riga corispondente all IdSigep e idprimavera
    '            sql = " insert into pm6_log (ID, DATETIME, DIRECTION, IDSIGEP, IDPM6, OPERATION, STATUS, NOTE) values    (" & _IdLog & ", sysdate, " & dir & " , " & idSigep & " , " & idPrimavera & ",'update',1,'')"
    '            DbUtility.QueryExec(sql, Primavera6.datasource.sigep)

    '            Dim p5 As New Primavera6(idPrimavera, idSigep, Me)
    '            WriteLog(_InitLabel)

    '            ' se faccio un aggiornamento NON IMPORTO LA WBS'
    '            'tranne se nn ne esiste nessuna ossia
    '            Dim lvlMax As Integer = 0
    '            lvlMax = DbUtility.QueryValue("select nvl(max(p_lvl),0)  from DEFN_LVL_STRU_T014   a where a.C_PROG= " & idSigep & "  and c_stru='WBSP6'", Primavera6.datasource.sigep)
    '            If lvlMax = 0 Then
    '                printResponse = p5.ImportWBSInSigep()
    '                WriteLog(_ImportWBSLabel, idSigep, idPrimavera)
    '                If printResponse = False Then
    '                    Throw New Exception("ImportWBSInSigep" & " " & _AbortMsg)
    '                End If
    '            End If
    '            '-------------------------------------------------------------------

    '            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////
    '            '' ''NON ELIMINO PIU I TASK MA LI AGGIORNO; INSERISCO O CANCELLO' 9 sett 09
    '            ' ''printResponse = p5.deleteTaskInSigep()
    '            ' ''WriteLog("Delete Task In Sigep:" & printResponse)
    '            ' ''If printResponse = False Then
    '            ' ''    Throw New Exception("deleteTaskInSigep" & " " & AbortMsg)
    '            ' ''End If
    '            ' ''checkAbort(idSigep, idprimavera)
    '            '' ''07 gennaio 09------------------
    '            ' ''If exists(idSigep, idprimavera, dir) Then Exit Function
    '            '' ''-----------------------------------
    '            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////

    '            ' task deletion reintroduced as per request - 11/11/2013 - MP
    '            ' eventually check rows 890 to 903 in procedure insertAllWbsInSigep, class Primavera6
    '            ' to manage appropriately data in tables prog_track_t091, cost_rev_asso_task, po_task_pbs_codes, msp_task_assignment
    '            printResponse = p5.deleteTaskInSigep()
    '            WriteLog("Delete Task In Sigep:" & printResponse)
    '            If printResponse = False Then
    '                Throw New Exception("deleteTaskInSigep" & " " & _AbortMsg)
    '            End If
    '            CheckAbort(idSigep, idPrimavera)
    '            If ConcurrentUpdate(idSigep, idPrimavera, dir) Then Return False

    '            printResponse = p5.insertAllWbsInSigep()
    '            WriteLog(_InsertAllWBSLabel, idSigep, idPrimavera)
    '            If printResponse = False Then
    '                Throw New Exception("insertAllWbsInSigep" & " " & _AbortMsg)
    '            End If

    '            CheckAbort(idSigep, idPrimavera)
    '            '07 gennaio 09------------------
    '            If ConcurrentUpdate(idSigep, idPrimavera, dir) Then Return False
    '            '-----------------------------------

    '            ' MP 7/7/2011: le milestones si devono "fermare" nella msp_tasks
    '            'printResponse = p5.insertAllMilestoneInSigep()
    '            'WriteLog(importMilestoneLabel, idSigep, idprimavera)
    '            'If printResponse = False Then
    '            '    Throw New Exception("insertAllMilestoneInSigep" & " " & AbortMsg)
    '            'End If

    '            'checkAbort(idSigep, idprimavera)
    '            ''07 gennaio 09------------------
    '            'If exists(idSigep, idprimavera, dir) Then Exit Function
    '            ' MP 7/7/2011 END [le milestones si devono "fermare" nella msp_tasks]

    '            '-----------------------------------
    '            '' '' ''6-marzo-09 non aggiorno più le date, descrizione e titolo in sigep
    '            '' '' ''printResponse = p5.updateProjectDates()
    '            '' '' ''If printResponse = False Then
    '            '' '' ''    Throw New Exception("updateProjectDates")
    '            '' '' ''End If

    '            ' update timestamp
    '            Try
    '                DbUtility.QueryExec("UPDATE PROG_T056 SET LAST_P6_UPDATE = " & PO.Base.DBUtilities.ServerDate(p5.SigepDBType) & " where C_PROG = " & idSigep, Primavera6.datasource.sigep)
    '            Catch ex As Exception
    '                Throw ex
    '            End Try

    '            ' update project association - just to refresh the external time-now in SIGEP
    '            printResponse = updateProjectAssociation(idPrimavera, idSigep)
    '            If Not printResponse Then
    '                Throw New Exception("insertAllWbsInSigep" & " " & _AbortMsg)
    '            End If

    '            ' La fine dell'update viene identificata dal flag status=2  sulla riga corispondente all IdSigep e idPrimavera
    '            sql = " update pm6_log set status =2 ,note = 'OK' where  id = " & _IdLog & " and  IDSIGEP = " & idSigep & " and  IDPM6 = " & projId
    '            DbUtility.QueryExec(sql, Primavera6.datasource.sigep)

    '            WriteLog("updateProjectDates:" & printResponse)
    '            WriteLog("Processing end at :" & Now & vbCrLf & vbCrLf & vbCrLf)
    '        ElseIf direction = 0 Then
    '            '///// 0: adesso uso questa direzione per aggiornare SIGEP (importando un progetto PM6 in sigep) da SIGEP
    '            '// il parametro di ingresso è  il CODICE PROGETTO appena creato in SIGEP col nome del progetto primavera6 
    '            'IMPORT PROJECT
    '            '///////////////////////////////////
    '            idPrimavera = 0
    '            idSigep = CInt(projId)
    '            idPrimavera = CInt(DbUtility.QueryTable("select pm6_projid  from prog_t056 where c_prog  = " & projId, Primavera6.datasource.sigep).Rows(0)(0))
    '            WriteLog("Processing project : " & idPrimavera & "-" & idSigep)
    '            'prima di procedere all'importazione verifico che non ci sia un update del progetto da importare
    '            Dim isUpdating As Integer = DbUtility.QueryValue("select count(*) from pm6_log where direction=1 and status=1 and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera, Primavera6.datasource.sigep)
    '            If isUpdating > 0 Then
    '                sql = " insert into pm6_log (ID, DATETIME, DIRECTION, IDSIGEP, IDPM6, OPERATION, STATUS, NOTE) values    (" & _IdLog & " , sysdate, 0 , " & idSigep & " , " & idPrimavera & ",'import',4,'Aborted For Simultaneous Upgrade')"
    '                DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '                _TransSigep.Rollback()
    '                Return False
    '            End If

    '            ' L'inizio dell'importazione viene identificato dal flag status=1  sulla riga corispondente all IdSigep e idPrimavera
    '            sql = " insert into pm6_log (ID, DATETIME, DIRECTION, IDSIGEP, IDPM6, OPERATION, STATUS, NOTE) values    (" & _IdLog & " , sysdate, 0 , " & idSigep & " , " & idPrimavera & ",'import',1,'')"
    '            DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '            Dim p5 As New Primavera6(idPrimavera, idSigep, Me)
    '            WriteLog(_InitLabel, idSigep, idPrimavera)

    '            'printResponse = p5.ImportWBSInSigep()
    '            'WriteLog(importWBSLabel, idSigep, idprimavera)
    '            'If printResponse = False Then
    '            '    Throw New Exception("ImportWBSInSigep")
    '            'End If

    '            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////
    '            '' ''NON ELIMINO PIU I TASK MA LI AGGIORNO; INSERISCO O CANCELLO' 9 sett 09
    '            ' ''printResponse = p5.deleteTaskInSigep()
    '            ' ''WriteLog("deleteTaskInSigep:" & printResponse)
    '            ' ''If printResponse = False Then
    '            ' ''    Throw New Exception("deleteTaskInSigep" & " " & AbortMsg)
    '            ' ''End If
    '            ' ''checkAbort(idSigep, idprimavera)
    '            '' ''/////////////////////////////////////////////////////////////////////////////////////////////////


    '            isUpdating = DbUtility.QueryValue("select count(*) from pm6_log where direction=1 and id > " & _IdLog & " and status=1 and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera, Primavera6.datasource.sigep)
    '            If isUpdating > 0 Then
    '                sql = " update pm6_log set note = 'Aborted For Simultaneous Upgrade' , status = 4 where  id = " & _IdLog & " and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera
    '                DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '                _TransSigep.Rollback()
    '                Return False
    '            End If

    '            printResponse = p5.insertAllWbsInSigep()
    '            WriteLog(_InsertAllWBSLabel, idSigep, idPrimavera)
    '            If printResponse = False Then
    '                Throw New Exception("insertAllWbsInSigep" & " " & _AbortMsg)
    '            End If
    '            CheckAbort(idSigep, idPrimavera)

    '            ' MP 7/7/2011: le milestones si devono "fermare" nella msp_tasks
    '            'printResponse = p5.insertAllMilestoneInSigep()
    '            'WriteLog(importMilestoneLabel, idSigep, idprimavera)
    '            'If printResponse = False Then
    '            '    Throw New Exception("insertAllMilestoneInSigep" & " " & AbortMsg)
    '            'End If
    '            'checkAbort(idSigep, idprimavera)
    '            ' MP 7/7/2011 END [le milestones si devono "fermare" nella msp_tasks]

    '            Dim lvlMax As Integer = 0
    '            lvlMax = DbUtility.QueryValue("select nvl(max(p_lvl),0)  from DEFN_LVL_STRU_T014   a where a.C_PROG= " & idSigep & "  and c_stru='WBSP6'", Primavera6.datasource.sigep)
    '            If lvlMax = 0 Then
    '                printResponse = p5.ImportWBSInSigep()
    '                WriteLog(_ImportWBSLabel, idSigep, idPrimavera)
    '                If printResponse = False Then
    '                    Throw New Exception("ImportWBSInSigep" & " " & _AbortMsg)
    '                End If
    '            End If

    '            ' update timestamp
    '            DbUtility.QueryExec("UPDATE PROG_T056 SET LAST_P6_UPDATE = " & PO.Base.DBUtilities.ServerDate(p5.SigepDBType) & " where C_PROG = " & idSigep, Primavera6.datasource.sigep)

    '            ' La fine dell'importazione viene identificata dal flag status=2  sulla riga corispondente all IdSigep e idPrimavera
    '            sql = " update pm6_log set status =2, note = 'OK'  where   id = " & _IdLog & " and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera
    '            DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '            WriteLog("Processing end at :" & Now & vbCrLf & vbCrLf & vbCrLf)

    '            'prima di chiudere verifico che nel frattempo nn ci sia stata una UPDATE del progetto importato
    '            isUpdating = DbUtility.QueryValue("select count(*) from pm6_log where direction=1 and id > " & _IdLog & " and status=1 and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera, Primavera6.datasource.sigep)
    '            If isUpdating > 0 Then
    '                sql = " update pm6_log set note = 'Old Version Imported' , status = 5 where  id = " & _IdLog & " and  IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera
    '                DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '                _TransSigep.Rollback()
    '                Return False
    '            End If
    '        End If
    '        _TransSigep.Commit()
    '        Return True
    '    Catch ex As Exception
    '        _TransSigep.Rollback()
    '        If direction = 1 Then
    '            sql = " update pm6_log set note = '" & Left(ex.Message.Replace("'", ""), 150) & "', status=3 where id = " & _IdLog & " and IDSIGEP = " & idSigep & " and  IDPM6 = " & projId
    '        ElseIf direction = 0 Then
    '            sql = " update pm6_log set note = '" & Left(ex.Message.Replace("'", ""), 150) & "', status=3 where id = " & _IdLog & " and IDSIGEP = " & idSigep & " and  IDPM6 = " & idPrimavera
    '        End If
    '        Try
    '            DbUtility.QueryExec(sql, Primavera6.datasource.sigep)
    '        Catch
    '        End Try
    '        WriteLog(ex.Message)
    '        Return False
    '    Finally
    '        If Not _CnSigep Is Nothing Then
    '            If _CnSigep.State = ConnectionState.Open Then _CnSigep.Close()
    '            _CnSigep.Dispose()
    '        End If
    '        If Not _DmlSigepCmd Is Nothing Then _DmlSigepCmd.Dispose()
    '        If Not _TransSigep Is Nothing Then _TransSigep.Dispose()
    '    End Try

    'End Function

    ''' <summary>
    ''' Added by Anubrij (POINDIA)
    ''' </summary>
    ''' <param name="idPM6">PM6 project ID</param>
    ''' <param name="idSigep">Sigep project ID</param>
    ''' <returns>Return boolean</returns>
    ''' <remarks>Update the Primavera sigep project association before performing publish as</remarks>
    Public Function UpdateProjectAssociation(poConn As IDataClient, ByVal idPM6 As Integer, ByVal idSigep As Integer) As Boolean

        Dim externalTransaction As Boolean = poConn.TransactionActive
        Dim primaveraConn As IDataClient = PO.DBInterface.IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
        Dim pm6TimeNow As DateTime = DateTime.MinValue
        Dim dbsysdate As String = "sysdate"
        If IBaseDataMgr.GetDBType(DataSource.PO) = PO.Base.DBType.SQLServer Then dbsysdate = "getdate()"

        Try
            ' reading P6 time-now
            If Not primaveraConn.ExecuteScalar("select LAST_RECALC_DATE from PROJECT where PROJ_ID = " & idPM6.ToString(), pm6TimeNow) Then pm6TimeNow = DateTime.MinValue

            Dim timeNowUpdateString As String = " NULL "
            If DateTime.Compare(pm6TimeNow, DateTime.MinValue) <> 0 Then
                timeNowUpdateString = PO.Base.DBUtilities.DBDateFormula(IBaseDataMgr.GetDBType(DataSource.PO), pm6TimeNow.ToString("yyyyMMdd"), "YYYYMMDD")
            End If

            ' start transaction if necessary
            If Not externalTransaction Then poConn.BeginTransaction(IsolationLevel.ReadCommitted)

            poConn.ExecuteNonQuery("DELETE FROM MSP_TASKS WHERE EXISTS (" &
                                                    "SELECT C_PROG FROM PROG_T056 " &
                                                    " WHERE PROG_T056.C_PROG = MSP_TASKS.PROJ_ID " &
                                                    " AND PROG_T056.PM6_PROJID = " & idPM6 &
                                                    " AND PROG_T056.C_PROG <> " & idSigep & ")")
            If Not poConn.LastException Is Nothing Then Throw poConn.LastException

            poConn.ExecuteNonQuery("UPDATE PROG_T056 SET PM6 = null, pm6_projid = null, d_cor_external = null where pm6_projid = " & idPM6)
            If Not poConn.LastException Is Nothing Then Throw poConn.LastException

            poConn.ExecuteNonQuery("UPDATE PROG_T056 SET PM6 = null, pm6_projid = null, d_cor_external = null where C_PROG = " & idSigep)
            If Not poConn.LastException Is Nothing Then Throw poConn.LastException

            poConn.ExecuteNonQuery("UPDATE PROG_T056 SET PM6 = 1, pm6_projid = " & idPM6 & ", d_cor_external = " & timeNowUpdateString & ", LAST_P6_UPDATE = " & dbsysdate & " where C_PROG = " & idSigep)
            If Not poConn.LastException Is Nothing Then Throw poConn.LastException

            If Not externalTransaction Then poConn.CommitTransaction()

            Return True

        Catch ex As Exception
            If Not externalTransaction AndAlso poConn.TransactionActive Then poConn.RollbackTransaction()
            Me.WriteLog("Error in updateProjectAssociation : " & ex.Message)
            Return False
        End Try
    End Function
    Private Function UpdatePublishTimestamp(idSigep As Integer) As Boolean

        ' update timestamp
        Try
            _DMLDBConn.ExecuteNonQuery("UPDATE PROG_T056 SET LAST_P6_UPDATE = " & PO.Base.DBUtilities.ServerDate(PO.DBInterface.IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.PO)) & " where C_PROG = " & idSigep)
            If Not _DMLDBConn.LastException Is Nothing Then Throw _DMLDBConn.LastException
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function





End Class
