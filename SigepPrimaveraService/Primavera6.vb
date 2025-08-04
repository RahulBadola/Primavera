Imports System.Data.OleDb
Imports System.Linq
Imports PO.Base
Imports PO.DataClient
Imports PO.DBInterface

Public Class Primavera6
    Private Class MyComparer
        Implements IComparer

        Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim a As New task
            a = CType(x, task)
            Dim b As New task
            b = CType(y, task)
            Return a.w.number.CompareTo(b.w.number)
        End Function

    End Class

    Public Enum datasource
        sigep
        pm5
    End Enum

    Public Enum taskType
        TT_Task = 0
        TT_Mile = 1
        TT_FinMile = 2
    End Enum

    Private _LastException As Exception
    'Private _connectionstring As String
    Private _projIdPM6 As String
    Private _projIdSigep As String

    Private _task As New List(Of task)
    Private _wbs As New List(Of wbs)
    'Private _taskLinks As New List(Of TaskLinks)

    Private _Counter As Integer

    Private _wbsPrefix As String
    Private _IdToFind As Integer
    'Private groupingMode As Boolean = False

    Private _Caller As SigepPrimaveraApp
    Private _WbsHash As New Hashtable
    'Private _wbsDataTable As New DataTable

    Private _SigepProjStatus As Integer = 0
    Private _SigepProjCompany As Integer = 0
    Private _GroupByActivityCode As Integer = 0
    Private _SigepProjNameAndCompany As String = String.Empty
    Private _SigepProjLaborQtyFlag As Integer = -1

    ' Private cmdSigep As New OleDbCommand
    'Private _skylineFlagCol As String
    'Private _slipChartFlagCol As String
    'Private _skylineAttrCol As String
    'Private _pptFlagCol As String
    Private _CustomAttributes As Dictionary(Of Integer, Dictionary(Of P6CustomAttributes.AttributeCode, String)) = Nothing

    Private _LaborQtyFlag As Boolean = True
    Private _ChainNames As New SortedList(Of String, String) ' used to store WBS codes and perform a quicker search

    Private _P6OpenProjWithSigepDraftProj As Boolean = False

    Private _ProjectCalendar As P6Calendar = Nothing
    Private _TaskCalendars As New Dictionary(Of Integer, P6Calendar)

    Public Sub New(ByVal pm6Id As String, ByVal sigepId As String, ByVal caller As SigepPrimaveraApp)

        Dim objLock As Object = New Object()
        Dim p6Conn As PO.DataClient.IDataClient = IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
        Dim pm6ProgressedTasks As Integer = 0
        Dim p6Null As String = DBUtilities.IsNull(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.Primavera))
        _LastException = Nothing

        SyncLock objLock

            _projIdPM6 = pm6Id
            _projIdSigep = sigepId
            _Caller = caller

            '' P6 project general information
            'Dim projDataTable As DataTable = p6Conn.ExecuteQuery("Select proj_short_name, plan_start_date, plan_end_date from project  where proj_id=" & _projIdPM6).Tables(0)
            'If projDataTable.Rows.Count > 0 Then
            '    Dim projRow As DataRow
            '    projRow = projDataTable.Rows(0)
            '    _P6ProjectName = projRow(0)
            '    _ProjectPlannedStart = IIf(IsDBNull(projRow(1)), Nothing, projRow(1))
            '    _ProjectPlannedEnd = IIf(IsDBNull(projRow(2)), Nothing, projRow(1))
            'End If

            ' reading main SIGEP project information
            ReadSigepProjectInfo()

            ' counting active\completed tasks for the Primavera project
            p6Conn.ExecuteScalar("select " & p6Null & "(sum(" & p6Null & "(complete_cnt, 0) + " & p6Null & "(active_cnt, 0)), 0) from sumtask where proj_id = " & _projIdPM6, pm6ProgressedTasks)
            _P6OpenProjWithSigepDraftProj = (pm6ProgressedTasks > 0) AndAlso (_SigepProjStatus = 0)

            Using x As New P6Calendar()
                _ProjectCalendar = x.LoadProjectCalendar(_projIdPM6)
            End Using

            'imposto i task
            Me.setTask()

            'imposto la wbs
            Me.setWbs()

            ' recupero di attributi aggiuntivi milestones
            'LoadMlstnAttr()
            LoadCustomAttributes()

            '///////////////////////   controllo che la wbs non parta da zero ///////////////////////////////////
            'cerco il livello minimo
            Dim m = GetMinWbs()
            'se livello min=0 then aggiungo a tutti i livelli 1
            Dim gap As Integer = m - 1
            If m <> 1 Then
                For Each tt As task In _task
                    tt.w.level -= gap
                Next
            End If
            '////////////////////////////////////////////////////////////////////////////////////////////////////

            'ora devo aggiungere i task fittizi della wbs
            Dim j As Integer = 0
            'Dim exists As Boolean = False
            'Dim cont As Integer = 0
            'Dim t As New task()
            For j = 1 To _task.Count
                _task(j - 1).w.number = j
            Next

            'For Each ww As wbs In _wbs
            '    setSummaryDateAndPercComplete(ww)
            'Next
            ' setSummaryDateAndPercComplete
            '      setSummaryPercComp()


        End SyncLock

        '_Caller.CheckAbort(_projIdSigep, _projIdPM6)

    End Sub

    Private Function ReadSigepProjectInfo() As Boolean

        Dim isnull As String = DBUtilities.IsNull(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.PO))
        Dim data As DataSet = _Caller._DMLDBConn.ExecuteQuery("select C_AZD, " & isnull & "(F_STA, 0) F_STA, " & isnull & "(PM6_GROUPINGBY,0) PM6_GROUPINGBY, S_PROG_NOM, " & isnull & "(USE_RES_BDG_LABOR_VALUES, -1) USE_RES_BDG_LABOR_VALUES from prog_t056 where c_prog=" & Me._projIdSigep)
        Try
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            If data.Tables(0).Rows.Count <> 1 Then Throw New Exception("Could not retrieve information for SIGEP project " & _projIdSigep.ToString())
            Dim row As DataRow = data.Tables(0).Rows(0)
            _SigepProjStatus = Convert.ToInt32(row("F_STA"))
            _SigepProjCompany = Convert.ToInt32(row("C_AZD"))
            _GroupByActivityCode = Convert.ToInt32(row("PM6_GROUPINGBY"))
            If Not IsDBNull(row("S_PROG_NOM")) Then _SigepProjNameAndCompany = DirectCast(row("S_PROG_NOM"), String)
            _SigepProjLaborQtyFlag = Convert.ToInt32(row("USE_RES_BDG_LABOR_VALUES"))

            Return True
        Catch ex As Exception
            _LastException = ex
            Return False
        End Try

    End Function
    Public ReadOnly Property LastException As Exception
        Get
            Return _LastException
        End Get
    End Property
    Private Function SetStartDate(ByVal currDate As DateTime, ByVal newDate As DateTime) As DateTime
        If DateTime.Compare(newDate, DateTime.MinValue) = 0 Then Return currDate
        If DateTime.Compare(newDate, DateTime.MaxValue) = 0 Then Return currDate

        If DateTime.Compare(currDate, DateTime.MinValue) = 0 Then Return newDate
        If DateTime.Compare(currDate, DateTime.MaxValue) = 0 Then Return newDate

        If currDate > newDate Then
            Return newDate
        Else
            Return currDate
        End If

    End Function
    Private Function UpdateEndDate(ByVal currDate As DateTime, ByVal newDate As DateTime) As DateTime
        If DateTime.Compare(newDate, DateTime.MinValue) = 0 Then Return currDate
        If DateTime.Compare(newDate, DateTime.MaxValue) = 0 Then Return currDate

        If DateTime.Compare(currDate, DateTime.MinValue) = 0 Then Return newDate
        If DateTime.Compare(currDate, DateTime.MaxValue) = 0 Then Return newDate

        If currDate < newDate Then
            Return newDate
        Else
            Return currDate
        End If

    End Function
    Private Function SetSummaryDateAndPercComplete(ByVal w As wbs)

        Dim actEndDate As DateTime = DateTime.MinValue
        Dim actStartDate As DateTime = DateTime.MinValue
        Dim fcastEndDate As DateTime = DateTime.MinValue
        Dim fcastStartDate As DateTime = DateTime.MinValue
        Dim blineEndDate As DateTime = DateTime.MinValue
        Dim blineStartDate As DateTime = DateTime.MinValue
        Dim startDates As List(Of DateTime)
        Dim endDates As List(Of DateTime)
        Dim durationTot As Decimal = 0
        Dim actualDurationTot As Decimal = 0
        Dim remainDurationTot As Decimal = 0
        Dim percComplTot As Decimal = 0

        If Not (w.nodeType AndAlso w.visible) Then Return True
        Try

            Dim children As List(Of wbs) = (From item In _wbs Where item.parentId = w.id Select item Order By item.level Descending).ToList()
            For Each child As wbs In children
                SetSummaryDateAndPercComplete(child)
                '' actual dates
                'If Not (child.actualEndDate = Nothing OrElse IsDBNull(child.actualEndDate) OrElse DateTime.Compare(child.actualEndDate, DateTime.MinValue) = 0) Then
                '    If actualEndDate < child.actualEndDate Then actualEndDate = child.actualEndDate
                'End If
                '' forecast dates
                'If Not (child.forecastEndDate = Nothing OrElse IsDBNull(child.forecastEndDate) OrElse DateTime.Compare(child.forecastEndDate, DateTime.MinValue) = 0) Then
                '    If forecastEndDate < child.forecastEndDate Then forecastEndDate = child.forecastEndDate
                'End If
                '' baseline dates
                'If Not (child.baselineEndDate = Nothing OrElse IsDBNull(child.baselineEndDate) OrElse DateTime.Compare(child.baselineEndDate, DateTime.MinValue) = 0) Then
                '    If baselineEndDate < child.baselineEndDate Then baselineEndDate = child.baselineEndDate
                'End If

                '' durations
                'actualDurationTot += child.actualDuration
                'remainDurationTot += child.remainDuration
                'If durationTot = 0 Then
                '    percComplTot = 0
                'Else
                '    percComplTot += child.percentComplete * (child.remainDuration + child.actualDuration) / durationTot
                'End If
            Next

            ' forecast dates
            startDates = Nothing
            startDates = (From item In _wbs Where item.parentId = w.id AndAlso DateTime.Compare(item.forecastStartDate, DateTime.MinValue) <> 0 Select item.forecastStartDate).ToList()
            If startDates.Count > 0 Then
                fcastStartDate = startDates.Min()
                w.forecastStartDate = SetStartDate(w.forecastStartDate, fcastStartDate)
            End If

            endDates = Nothing
            endDates = (From item In _wbs Where item.parentId = w.id Select item.forecastEndDate).ToList()
            If endDates.Count > 0 AndAlso Not endDates.Contains(DateTime.MinValue) Then
                fcastEndDate = endDates.Max()
                w.forecastEndDate = UpdateEndDate(w.forecastEndDate, fcastEndDate)
            End If

            ' baseline dates
            startDates = (From item In _wbs Where item.parentId = w.id AndAlso DateTime.Compare(item.baselineStartDate, DateTime.MinValue) <> 0 Select item.baselineStartDate).ToList()
            If startDates.Count > 0 Then
                blineStartDate = startDates.Min()
                w.baselineStartDate = SetStartDate(w.baselineStartDate, blineStartDate)
            End If

            endDates = Nothing
            endDates = (From item In _wbs Where item.parentId = w.id AndAlso DateTime.Compare(item.baselineEndDate, DateTime.MinValue) <> 0 Select item.baselineEndDate).ToList()
            If endDates.Count > 0 AndAlso Not endDates.Contains(DateTime.MinValue) Then
                blineEndDate = endDates.Max()
                w.baselineEndDate = UpdateEndDate(w.baselineEndDate, blineEndDate)
            End If
            ' if there's no baseline information, let's take the forecast...
            If DateTime.Compare(DateTime.MinValue, w.baselineStartDate) = 0 Then w.baselineStartDate = w.forecastStartDate
            If DateTime.Compare(DateTime.MinValue, w.baselineEndDate) = 0 Then w.baselineEndDate = w.forecastEndDate

            ' actual dates
            startDates = Nothing
            startDates = (From item In _wbs Where item.parentId = w.id AndAlso DateTime.Compare(item.actualStartDate, DateTime.MinValue) <> 0 Select item.actualStartDate).ToList()
            If startDates.Count > 0 Then
                actStartDate = startDates.Min()
                w.actualStartDate = SetStartDate(w.actualStartDate, actStartDate)
            End If
            endDates = Nothing
            endDates = (From item In _wbs Where item.parentId = w.id Select item.actualEndDate).ToList()
            If endDates.Count > 0 AndAlso Not endDates.Contains(DateTime.MinValue) Then
                actEndDate = endDates.Max()
                w.actualEndDate = UpdateEndDate(w.actualEndDate, actEndDate)
            End If

            ' durations
            durationTot = (From item In _wbs Where item.parentId = w.id Select item.actualDuration + item.remainDuration).Sum()
            actualDurationTot = (From item In _wbs Where item.parentId = w.id Select item.actualDuration).Sum()
            remainDurationTot = (From item In _wbs Where item.parentId = w.id Select item.remainDuration).Sum()
            If durationTot = 0 Then : percComplTot = 0
            Else
                percComplTot = (From item In _wbs Where item.parentId = w.id Select item.percentComplete * ((item.actualDuration + item.remainDuration) / durationTot)).Sum()
            End If

            'children = (From item In _wbs Where item.parentId = w.id Select item).ToList()
            'If durationTot = 0 Then : percComplTot = 0
            'Else
            '    percComplTot = 0
            '    For Each child As wbs In children
            '        percComplTot += child.percentComplete * (child.remainDuration + child.actualDuration) / durationTot
            '    Next
            'End If

            w.percentComplete = percComplTot
            w.actualDuration = actualDurationTot
            w.remainDuration = remainDurationTot

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    'Public Sub setSummaryPercComp()
    '    Dim cont As Integer
    '    Dim maxLevel As Integer
    '    Dim childs As ArrayList
    '    maxLevel = getMaxWbsLevel()
    '    For cont = maxLevel To 0 Step -1
    '        For Each ww As wbs In _wbs
    '            If ww.level = cont Then
    '                childs = getChild(ww)
    '                If ww.nodeType = True And ww.visible = True Then  'se  un nodo di wbs
    '                    Dim percComplTmp As New ArrayList()
    '                    Dim percComplTot As Decimal = 0
    '                    Dim taskInWbsNode As New ArrayList
    '                    Dim durationTot As Decimal = 0
    '                    Dim ActualDurationTot As Decimal = 0
    '                    Dim RemainDurationTot As Decimal = 0
    '                    For Each www As wbs In childs
    '                        durationTot += www.remainDuration + www.actualDuration
    '                    Next
    '                    'una volta calcolata la duration totale del nodo posso calcoloare la per_compl del nodo facendo la somma pesata dei task che lo compongono
    '                    For Each www As wbs In childs
    '                        ActualDurationTot += www.actualDuration
    '                        RemainDurationTot += www.remainDuration
    '                        If durationTot = 0 Then
    '                            percComplTot = 0
    '                        Else
    '                            percComplTot += www.percentComplete * (www.remainDuration + www.actualDuration) / durationTot
    '                        End If
    '                    Next
    '                    ww.percentComplete = percComplTot
    '                    ww.actualDuration = ActualDurationTot
    '                    ww.remainDuration = RemainDurationTot
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub

    Private Sub setTask(Optional ByVal loadMilestones As Boolean = False)
        Try

            _Caller.WriteLog("Creating Tasks", _projIdSigep, _projIdPM6)

            Dim tasksStmt As String = String.Empty
            Dim calendarId As Integer
            Dim taskId As Long
            Dim calendar As P6Calendar = Nothing
            Dim emptyCalendars As New List(Of Integer)
            Dim hoursPerDay As Decimal
            'Dim projLaborQtyFlag As Integer
            Dim isnull As String = DBUtilities.IsNull(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.Primavera))
            Dim sigepIsNullFunction As String = DBUtilities.IsNull(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.PO))
            Dim p6Concat As String = DBUtilities.StrConcatOperator(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.Primavera))
            Dim sigepConcat As String = DBUtilities.StrConcatOperator(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.PO))
            Dim p6Conn As IDataClient = IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
            'vedo dalle impostazioni di progetto in sigep se il livello di dettaglio è activities o groupBY
            'se è scelto groupBY leggo in base a quale activitycode.
            'da ricordare che nel caso di attività raggruppate vanno sostituite con la risultante avente come data minima la minima delle singole e come massima la massima delle singole mentre il nome è preso dall'activity code più il nome della wbs di appartenenza (per eviare doppioni)
            '_Caller._DMLDBConn.ExecuteScalar("select nvl(PM6_GROUPINGBY, 0) from prog_t056 where c_prog = " & _projIdSigep, _GroupByActivityCode)

            If loadMilestones Then
                tasksStmt = "Select * from task a where a.PROJ_ID=" & _projIdPM6 & "  And task_type <>'TT_Task'"
            Else
                If _GroupByActivityCode = 0 Then
                    tasksStmt = "select * from task where delete_date is null and proj_id = " & _projIdPM6 & " order by wbs_id, task_id "

                    'prendo anche il campo TOTAL_DRTN_HR_CNT da tasksum per inserire nei nodi summary la task_duration
                    'sql = "select task.*,tasksum.TOTAL_DRTN_HR_CNT from task,tasksum where task.delete_date is null and task.proj_id = " & _projId & " and task.wbs_id=tasksum.wbs_id"
                Else
                    'groupingMode = True
                    '''' commented out on 15/10/2014 as 'sql' variable is overwritten below - MP
                    'sql = " select a.WBS_ID,  a.actv_code_id, a.actv_code_name as actv_code_name , start_act,finish_act,start_target,finish_target from "
                    'sql = sql & "  ("
                    'sql = sql & "  SELECT   "
                    'sql = sql & "  t1.WBS_ID, "
                    'sql = sql & "  t2.actv_code_id, "
                    '' sql = sql & "  t3.actv_code_name || '_' || t1.WBS_ID as actv_code_name "
                    'sql = sql & "  t3.actv_code_name || '_' || t4.WBS_short_name as actv_code_name "
                    'sql = sql & "  FROM    "
                    'sql = sql & "  task t1 inner join  taskactv t2  on  t1.TASK_id = t2.task_id  inner join  actvcode t3 on  t3.ACTV_CODE_ID= t2.actv_code_id "
                    'sql = sql & "  inner join projwbs t4 on t4.wbs_id = t1.wbs_id  "
                    'sql = sql & "  WHERE  (actv_code_type_id= {0} or actv_code_type_id is null )and   t1.delete_date is null and   t1.proj_id = {1}    "
                    'sql = sql & "  group by   t2.actv_code_id ,t1.WBS_ID,t2.actv_code_id,actv_code_name  "
                    'sql = sql & " , t4.WBS_short_name"
                    'sql = sql & "  )a"
                    'sql = sql & "  inner join "
                    'sql = sql & "  ( "
                    'sql = sql & "  select  "
                    'sql = sql & "  wbs_id,actv_code_id, "
                    ''modifica 21 luglio
                    ''sql = sql & "  to_char(min(target_start_date),'dd/mm/yyyy') as start_target,"
                    ''sql = sql & "  to_char(min(target_end_date),'dd/mm/yyyy') as finish_target ,"
                    'sql = sql & "  to_char(min(EARLY_START_DATE),'dd/mm/yyyy') as start_target,"
                    'sql = sql & "  to_char(max(EARLY_END_DATE),'dd/mm/yyyy') as finish_target ," ' 22 luglio prima era to_char(min(EARLY_END_DATE),'dd/mm/yyyy') as finish_target
                    'sql = sql & "  to_char(min(act_start_date),'dd/mm/yyyy') as start_act, "
                    'sql = sql & "  to_char(max(act_end_date),'dd/mm/yyyy') as finish_act  " ' 22 luglio prima era to_char(min(act_end_date),'dd/mm/yyyy') as finish_act  
                    'sql = sql & "  from task inner join  taskactv on task.TASK_ID=taskactv.TASK_ID  "
                    'sql = sql & "  where  task.PROJ_ID={1}  "
                    'sql = sql & "  and actv_code_type_id={0} "
                    'sql = sql & "  group by actv_code_id,wbs_id"
                    'sql = sql & "  )b on b.wbs_id=a.wbs_id and a.actv_code_id=b.actv_code_id order by a.wbs_id"
                    '''' END commented out on 15/10/2014 as 'sql' variable is overwritten below - MP

                    tasksStmt = "  select distinct aa.wbs_id,aa.actv_code_id, aa.actv_code_name AS actv_code_name, aa.start_act, aa.finish_act, aa.start_target,aa. finish_target , baseline_start,baseline_end, CLNDR_ID from   "
                    tasksStmt = tasksStmt & "          (  "
                    tasksStmt = tasksStmt & "          SELECT   a.wbs_id, a.actv_code_id, a.actv_code_name AS actv_code_name, start_act, finish_act, start_target, finish_target, CLNDR_ID "
                    tasksStmt = tasksStmt & "            FROM (SELECT   t1.wbs_id, t2.actv_code_id, t3.actv_code_name  " & p6Concat & " '_'  " & p6Concat & " t4.wbs_short_name AS actv_code_name, t1.CLNDR_ID       "
                    tasksStmt = tasksStmt & "                           FROM task t1  "
                    tasksStmt = tasksStmt & "                           INNER JOIN taskactv t2 ON t1.task_id = t2.task_id  "
                    tasksStmt = tasksStmt & "                           INNER JOIN actvcode t3 ON t3.actv_code_id = t2.actv_code_id  "
                    tasksStmt = tasksStmt & "                           INNER JOIN projwbs t4 ON t4.wbs_id = t1.wbs_id  "
                    tasksStmt = tasksStmt & "                     WHERE (t3.actv_code_type_id = {0} OR t3.actv_code_type_id IS NULL)  "
                    tasksStmt = tasksStmt & "               AND t1.delete_date IS NULL  "
                    tasksStmt = tasksStmt & "               AND t1.proj_id = {1} "
                    tasksStmt = tasksStmt & "          GROUP BY t2.actv_code_id,  t1.wbs_id,  t2.actv_code_id, t3.actv_code_name, t4.wbs_short_name, t1.CLNDR_ID) a  "
                    tasksStmt = tasksStmt & "         INNER JOIN  "
                    tasksStmt = tasksStmt & "         (SELECT   wbs_id, actv_code_id , "
                    tasksStmt = tasksStmt & "                     TO_CHAR (MIN (case when EARLY_START_DATE is not null then EARLY_START_DATE else  target_start_date end ), 'dd/mm/yyyy' ) AS start_target,  "
                    tasksStmt = tasksStmt & "                     TO_CHAR (MAX (case when EARLY_END_DATE is not null then EARLY_END_DATE else target_END_date end), 'dd/mm/yyyy' ) AS finish_target,                     "
                    tasksStmt = tasksStmt & "                             TO_CHAR (MIN (act_start_date), 'dd/mm/yyyy') AS start_act, TO_CHAR (MAX (act_end_date), 'dd/mm/yyyy') AS finish_act  "
                    tasksStmt = tasksStmt & "                             FROM task INNER JOIN taskactv ON task.task_id = taskactv.task_id  "
                    tasksStmt = tasksStmt & "                     WHERE task.proj_id ={1}  AND taskactv.actv_code_type_id =  {0} "
                    tasksStmt = tasksStmt & "                  GROUP BY actv_code_id, wbs_id) b  "
                    tasksStmt = tasksStmt & "                 ON b.wbs_id = a.wbs_id AND a.actv_code_id = b.actv_code_id  "
                    tasksStmt = tasksStmt & "        ORDER BY a.wbs_id  "
                    tasksStmt = tasksStmt & "        ) AA  "
                    tasksStmt = tasksStmt & "         inner join   "
                    tasksStmt = tasksStmt & "        (    "
                    tasksStmt = tasksStmt & "        select a. wbs_id,a.actv_code_id,  "
                    tasksStmt = tasksStmt & "        min (case when EARLY_START_DATE is not null then  EARLY_START_DATE else  target_start_date end ) as baseline_start,  "
                    tasksStmt = tasksStmt & "        max(case when  EARLY_END_DATE is not null then EARLY_END_DATE else target_END_date end ) as baseline_end  from   "
                    tasksStmt = tasksStmt & "          (  "
                    tasksStmt = tasksStmt & "          SELECT   t1.wbs_id, t2.actv_code_id, t3.actv_code_name  " & p6Concat & " '_'  " & p6Concat & " t4.wbs_short_name AS actv_code_name     , t1.TASK_ID ,target_start_date, target_end_date "
                    tasksStmt = tasksStmt & "                   FROM task t1  "
                    tasksStmt = tasksStmt & "                   INNER JOIN taskactv t2 ON t1.task_id = t2.task_id  "
                    tasksStmt = tasksStmt & "                   INNER JOIN actvcode t3 ON t3.actv_code_id = t2.actv_code_id  "
                    tasksStmt = tasksStmt & "                   INNER JOIN projwbs t4 ON t4.wbs_id = t1.wbs_id  "
                    tasksStmt = tasksStmt & "                   WHERE (t3.actv_code_type_id = {0} OR t3.actv_code_type_id IS NULL)  "
                    tasksStmt = tasksStmt & "                  AND t1.delete_date IS NULL  "
                    tasksStmt = tasksStmt & "                  AND t1.proj_id ={1} order by wbs_id,actv_code_id  "
                    tasksStmt = tasksStmt & "                  )a  "
                    tasksStmt = tasksStmt & "                  inner join  "
                    tasksStmt = tasksStmt & "                   (  "
                    tasksStmt = tasksStmt & "                     SELECT case when T2.EARLY_START_DATE is not null then T2.EARLY_START_DATE else  t2.target_start_date end EARLY_START_DATE  ,  "
                    tasksStmt = tasksStmt & "                                  case when T2.EARLY_END_DATE is not null then T2.EARLY_END_DATE else  t2.target_END_date end EARLY_END_DATE, "
                    tasksStmt = tasksStmt & "                                   t1.WBS_ID,t1.TASK_NAME,t1.task_id  "
                    tasksStmt = tasksStmt & "                       FROM TASK  T1   "
                    tasksStmt = tasksStmt & "                       INNER JOIN PROJECT  P ON T1.PROJ_ID = P.PROJ_ID   "
                    tasksStmt = tasksStmt & "                       INNER JOIN TASK  T2 ON P.SUM_BASE_PROJ_ID = T2.PROJ_ID AND T1.TASK_CODE = T2.TASK_CODE  "
                    tasksStmt = tasksStmt & "                        WHERE P.PROJ_ID ={1} "
                    tasksStmt = tasksStmt & "                   )b  "
                    tasksStmt = tasksStmt & "                        on a.task_id=b.task_id  "
                    tasksStmt = tasksStmt & "                        group by a. wbs_id,a.actv_code_id  "
                    tasksStmt = tasksStmt & "             )BB   "
                    tasksStmt = tasksStmt & "           on aa.wbs_id=bb.wbs_id  and aa.actv_code_id=bb.actv_code_id  "
                    '///////////////////////////////////////////////////////////////////////////////////////////// fine ///////////////////////////////////////////////////////////////////////////////////
                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                    tasksStmt = String.Format(tasksStmt, _GroupByActivityCode, _projIdPM6)

                End If
            End If

            ' TODO: add appropriate field
            '_Caller._DMLDBConn.ExecuteScalar("select " & sigepIsNullFunction & "(USE_RES_BDG_LABOR_VALUES, -1) from prog_t056 where c_prog=" & Me._projIdSigep, projLaborQtyFlag)
            If _SigepProjLaborQtyFlag < 0 Then
                _LaborQtyFlag = (ConfigurationManager.AppSettings("labor_qty") = "1")
            Else
                _LaborQtyFlag = (_SigepProjLaborQtyFlag <> 0)
            End If

            'per evitare n letture al db mi carico in un datatable le date di baseline dei task
            'sql = "SELECT T2.EARLY_START_DATE, T2.EARLY_END_DATE, t1.task_id , t2.target_start_date, t2.target_end_date,  t1.task_id  FROM TASK  T1 INNER JOIN PROJECT  P ON T1.PROJ_ID = P.PROJ_ID INNER JOIN TASK  T2 ON P.SUM_BASE_PROJ_ID = T2.PROJ_ID AND T1.TASK_CODE = T2.TASK_CODE WHERE P.PROJ_ID =" & _projId
            'modifica 27 ago 09
            'Dim sql1 = "SELECT case when T2.EARLY_START_DATE is not null then T2.EARLY_START_DATE else  t2.target_start_date end EARLY_START_DATE , case when T2.EARLY_END_DATE is not null then T2.EARLY_END_DATE else  t2.target_END_date end EARLY_END_DATE,  t1.task_id  FROM TASK  T1 INNER JOIN PROJECT  P ON T1.PROJ_ID = P.PROJ_ID INNER JOIN TASK  T2 ON P.SUM_BASE_PROJ_ID = T2.PROJ_ID AND T1.TASK_CODE = T2.TASK_CODE WHERE P.PROJ_ID =" & _projId
            'modifica 15 novembre 2011
            Dim baselineStmt = "SELECT case when not T2.act_start_date is null then T2.act_start_date " &
                    " when not T2.EARLY_START_DATE is null then T2.EARLY_START_DATE " &
                    " else  t2.target_start_date end EARLY_START_DATE " &
                    " , case when not T2.act_END_date is null then T2.act_END_date " &
                    " when not T2.EARLY_END_DATE is null then T2.EARLY_END_DATE " &
                    " else  t2.target_END_date end EARLY_END_DATE,  t1.task_id  " &
                    " FROM TASK  T1 INNER JOIN PROJECT  P ON T1.PROJ_ID = P.PROJ_ID " &
                    " INNER JOIN TASK  T2 ON P.SUM_BASE_PROJ_ID = T2.PROJ_ID AND T1.TASK_CODE = T2.TASK_CODE " &
                    " WHERE P.PROJ_ID =" & _projIdPM6
            'fine
            Dim dtBaseline As DataTable = p6Conn.ExecuteQuery(baselineStmt).Tables(0)
            Dim dtTasks As DataTable = Nothing
            Try
                dtTasks = p6Conn.ExecuteQuery(tasksStmt).Tables(0)
            Catch ex As Exception
            End Try

            tasksStmt = ""
            Dim dtTaskLinks As DataTable = Nothing
            tasksStmt = " SELECT B.task_pred_id,B.TASK_ID, B.pred_task_id,B.proj_id,B.pred_proj_id,CASE  B.pred_type WHEN 'PR_SS' THEN 3 WHEN 'PR_SF' THEN 2 " &
                        " WHEN 'PR_FS' THEN 1 WHEN 'PR_FF' THEN 0 END pred_type, B.lag_hr_cnt,B.update_date,B.update_user,B.create_date,B.create_user,B.delete_session_id,B.delete_date " &
                        " FROM TASK A INNER JOIN TASKPRED B" &
                        " ON A.TASK_ID=B.TASK_ID WHERE B.DELETE_SESSION_ID IS NULL AND A.PROJ_ID= " & _projIdPM6 & " and B.PRED_PROJ_ID = " & _projIdPM6 & " ORDER BY B.TASK_ID "

            'Try
            '    dtTaskLinks = p6Conn.ExecuteQuery(tasksStmt).Tables(0)
            'Catch ex As Exception

            'End Try 
            'EXCLUDED AS WORK IS NOT COMPLETE YET == END
            For Each row As DataRow In dtTasks.Rows

                taskId = Convert.ToInt64(row("task_id"))
                If Not IsDBNull(row("CLNDR_ID")) Then
                    calendarId = Convert.ToInt32(row("CLNDR_ID"))
                    If Not _TaskCalendars.TryGetValue(calendarId, calendar) AndAlso Not emptyCalendars.Contains(calendarId) Then
                        Using x As New P6Calendar()
                            calendar = x.LoadCalendar(calendarId)
                            If Not calendar Is Nothing Then : _TaskCalendars.Add(calendar.CalendarId, calendar)
                            Else : emptyCalendars.Add(calendarId)
                            End If
                        End Using
                    End If
                Else
                    calendar = _ProjectCalendar
                End If
                hoursPerDay = 8
                If Not calendar Is Nothing Then
                    If calendar.HoursPerDay > 0 Then hoursPerDay = calendar.HoursPerDay
                End If

                Dim t As New task()

                If _GroupByActivityCode = 0 Or loadMilestones Then

                    t.id = taskId
                    t.name = IIf(IsDBNull(row("task_name")), String.Empty, row("task_name"))
                    t.description = t.name
                    'modifica 21 luglio
                    't.forecaststartDate = DbUtility.replaceNull(r("target_start_date"), DBNull.Value)
                    't.forecastEndDate = DbUtility.replaceNull(r("target_end_date"), DBNull.Value)
                    t.Calendar = calendar

                    'fine
                    t.forecaststartDate = DateTime.MinValue
                    t.forecastEndDate = DateTime.MinValue
                    t.actualStartDate = DateTime.MinValue
                    t.actualEndDate = DateTime.MinValue

                    If Not IsDBNull(row("EARLY_START_DATE")) Then
                        t.forecaststartDate = row("EARLY_START_DATE")
                    End If
                    If Not IsDBNull(row("EARLY_END_DATE")) Then
                        t.forecastEndDate = row("EARLY_END_DATE")
                    End If

                    ' if SIGEP project is in draft and P6 project has progress, actuals are ignored
                    If Not IsDBNull(row("act_start_date")) AndAlso Not _P6OpenProjWithSigepDraftProj Then
                        t.actualStartDate = row("act_start_date")
                    End If
                    If Not IsDBNull(row("act_end_date")) AndAlso Not _P6OpenProjWithSigepDraftProj Then
                        t.actualEndDate = row("act_end_date")
                    End If

                    'MODIFICA 23 LUGLIO SE LE ACTUAL ESISTONO METTERE LE FORECAST=ACTUAL
                    If _SigepProjStatus = 2 Then
                        If t.actualStartDate <> Nothing OrElse DateTime.Compare(t.actualStartDate, DateTime.MinValue) <> 0 Then
                            t.forecaststartDate = t.actualStartDate
                            If _P6OpenProjWithSigepDraftProj Then t.forecaststartDate = t.baselineStartDate
                        End If
                        If t.actualEndDate <> Nothing AndAlso DateTime.Compare(t.actualEndDate, DateTime.MinValue) <> 0 Then
                            t.forecastEndDate = t.actualEndDate
                            If _P6OpenProjWithSigepDraftProj Then t.forecastEndDate = t.baselineEndDate
                        End If
                    End If
                    '__________________________________________________________________

                    'new version
                    'Dim tmpDt As New DataTable
                    'tmpDt = getBaselineDate(t.id)
                    'If tmpDt.Rows.Count > 0 Then
                    Dim drBaseline As DataRow() = dtBaseline.Select("  TASK_ID = " & t.id)
                    t.baselineStartDate = DateTime.MinValue
                    t.baselineEndDate = DateTime.MinValue
                    If drBaseline.Length > 0 Then
                        If Not IsDBNull(drBaseline(0)(0)) Then
                            t.baselineStartDate = drBaseline(0)(0)
                        End If
                        If Not IsDBNull(drBaseline(0)(1)) Then
                            t.baselineEndDate = drBaseline(0)(1)
                        End If
                    End If
                    'end
                    If IsDBNull(row("phys_complete_pct")) Then
                        t.physicalPercentComplete = 0
                    Else
                        t.physicalPercentComplete = row("phys_complete_pct")
                    End If

                    't.percentComplete = SIGEPPRIMAVERA.DbUtility.replaceNull(r("phys_complete_pct"), DBNull.Value)
                    ' t.wbsId = r("wbs_id")
                    t.code = row("task_code")
                    t.activityType = row("task_type")
                    t.status = row("status_code")
                    t.w.id = row("wbs_id")
                    t.remainDurationP6 = row("remain_drtn_hr_cnt") / hoursPerDay
                    t.remainDuration = t.remainDurationP6 '/ hoursPerDay
                    t.totalDurationP6 = row("target_drtn_hr_cnt") / hoursPerDay
                    If Not _P6OpenProjWithSigepDraftProj Then
                        t.actualDuration = t.totalDurationP6 '/ hoursPerDay
                        t.actualDuration = t.actualDuration - t.remainDuration
                    End If

                    If DateTime.Compare(t.actualEndDate, DateTime.MinValue) <> 0 Then
                        t.percentComplete = 100
                    Else
                        'If (t.actualDuration + t.remainDuration) > 0 Then
                        '    t.percentComplete = 100 * t.actualDuration / (t.actualDuration + t.remainDuration)
                        'Else
                        '    t.percentComplete = 0
                        'End If
                        If t.totalDurationP6 <> 0 Then
                            t.percentComplete = 100 * ((t.totalDurationP6 - t.remainDurationP6) / t.totalDurationP6)
                        Else
                            t.percentComplete = 0
                        End If
                    End If

                    If Not IsDBNull(row("ACT_WORK_QTY")) AndAlso Not _P6OpenProjWithSigepDraftProj Then : t.HrsActualWork = Convert.ToDecimal(row("ACT_WORK_QTY")) : t.ActualWork = t.HrsActualWork / hoursPerDay : End If
                    If Not IsDBNull(row("REMAIN_WORK_QTY")) Then : t.HrsRemainingWork = Convert.ToDecimal(row("REMAIN_WORK_QTY")) : t.RemainingWork = t.HrsRemainingWork / hoursPerDay : End If
                    If Not IsDBNull(row("TARGET_WORK_QTY")) Then : t.HrsTargetWork = Convert.ToDecimal(row("TARGET_WORK_QTY")) : t.TargetWork = t.HrsTargetWork / hoursPerDay : End If
                    If Not IsDBNull(row("ACT_EQUIP_QTY")) AndAlso Not _P6OpenProjWithSigepDraftProj Then : t.HrsActualEqpmQty = Convert.ToDecimal(row("ACT_EQUIP_QTY")) : t.ActualEqpmQty = t.HrsActualEqpmQty / hoursPerDay : End If
                    If Not IsDBNull(row("REMAIN_EQUIP_QTY")) Then : t.HrsRemainingEqpmQty = Convert.ToDecimal(row("REMAIN_EQUIP_QTY")) : t.RemainingEqpmQty = t.HrsRemainingEqpmQty / hoursPerDay : End If
                    If Not IsDBNull(row("TARGET_EQUIP_QTY")) Then : t.HrsTargetEqpmQty = Convert.ToDecimal(row("TARGET_EQUIP_QTY")) : t.TargetEqpmQty = t.HrsTargetEqpmQty / hoursPerDay : End If
                    'If t.TargetWork - t.RemainingWork <> 0 Then t.WorkPercComplete = (t.ActualWork / (t.TargetWork - t.RemainingWork)) * 100.0
                    t.WorkPercComplete = t.percentComplete
                    'If t.TargetEqpmQty - t.RemainingEqpmQty <> 0 Then t.EqpmQtyPercComplete = (t.ActualEqpmQty / (t.TargetEqpmQty - t.RemainingEqpmQty)) * 100.0
                    t.EqpmQtyPercComplete = t.percentComplete
                Else

                    '///////////////// in questo caso non ho più la corrispondenza 1 a 1 con i task di primavera perchè ho fatto un raggruppamento per codice attività
                    '//////////////// quindi manca il t.id
                    t.status = "nullStatus"
                    t.w.id = row("wbs_id")
                    t.activityCode = 1
                    t.activityType = "nullType"
                    t.name = IIf(IsDBNull(row("actv_code_name")), String.Empty, row("actv_code_name"))
                    t.description = IIf(IsDBNull(row("actv_code_name")), String.Empty, row("actv_code_name"))
                    t.Calendar = calendar

                    t.forecaststartDate = DateTime.MinValue
                    t.forecastEndDate = DateTime.MinValue
                    t.actualStartDate = DateTime.MinValue
                    t.actualEndDate = DateTime.MinValue
                    t.baselineStartDate = DateTime.MinValue
                    t.baselineEndDate = DateTime.MinValue

                    If Not IsDBNull(row("start_target")) Then
                        t.forecaststartDate = row("start_target")
                    End If
                    If Not IsDBNull(row("finish_target")) Then
                        t.forecastEndDate = row("finish_target")
                    End If

                    If Not IsDBNull(row("start_act")) AndAlso Not _P6OpenProjWithSigepDraftProj Then
                        t.actualStartDate = row("start_act")
                    End If
                    If Not IsDBNull(row("finish_act")) AndAlso Not _P6OpenProjWithSigepDraftProj Then
                        t.actualEndDate = row("finish_act")
                    End If

                    If Not IsDBNull(row("baseline_start")) Then
                        t.baselineStartDate = row("baseline_start")
                        If _P6OpenProjWithSigepDraftProj Then t.forecaststartDate = t.baselineStartDate
                    End If
                    If Not IsDBNull(row("baseline_end")) Then
                        t.baselineEndDate = row("baseline_end")
                        If _P6OpenProjWithSigepDraftProj Then t.forecastEndDate = t.baselineEndDate

                    End If

                    t.id = row("wbs_id") & row("actv_code_id")
                    ' If Not t.actualEndDate = Nothing Then
                    't.percentComplete = 100
                    ' End If

                    'MODIFICA 23 LUGLIO SE LE ACTUAL ESISTONO METTERE LE FORECAST=ACTUAL
                    If _SigepProjStatus = 2 Then
                        If t.actualStartDate <> Nothing AndAlso DateTime.Compare(t.actualStartDate, DateTime.MinValue) <> 0 Then
                            t.forecaststartDate = t.actualStartDate
                        End If
                        If t.actualEndDate <> Nothing AndAlso DateTime.Compare(t.actualEndDate, DateTime.MinValue) <> 0 Then
                            t.forecastEndDate = t.actualEndDate
                        End If
                    End If
                    '__________________________________________________________________


                    '///////// aggiungiamo la perc_complete dei task risultanti /////////////////7
                    tasksStmt = " select sum (peso*pct_complete), sum(target_drtn_hr_cnt), (sum (peso*pct_complete)/100)*sum(target_drtn_hr_cnt) " &
                    " ,  (sum(phys_complete_pct)/count(phys_complete_pct*100))  " &
                    " ,  sum(ACT_WORK_QTY) ACT_WORK_QTY,  sum(REMAIN_WORK_QTY) REMAIN_WORK_QTY,  sum(TARGET_WORK_QTY) TARGET_WORK_QTY  " &
                    " ,  sum(ACT_EQUIP_QTY) ACT_EQUIP_QTY,  sum(REMAIN_EQUIP_QTY) REMAIN_EQUIP_QTY,  sum(TARGET_EQUIP_QTY) TARGET_EQUIP_QTY  " &
                    " from ( " &
                    "       select a. *  " &
                    "           , case when (select sum( target_drtn_hr_cnt) from task inner join  taskactv on task.TASK_ID=taskactv.TASK_ID where task.PROJ_ID=" & _projIdPM6 & " and actv_code_id=a.actv_code_id and wbs_id=a.wbs_id)>0 " &
                    "               then target_drtn_hr_cnt/ ( select sum( target_drtn_hr_cnt) from task inner join  taskactv on task.TASK_ID=taskactv.TASK_ID where  task.PROJ_ID=" & _projIdPM6 & " and actv_code_id=a.actv_code_id and wbs_id=a.wbs_id) " &
                    "               else 0 " &
                    "           end as peso  " &
                    "       from  ( " &
                    "               select  phys_complete_pct, actv_code_id, wbs_id, task_name, target_drtn_hr_cnt " &
                    "               , ACT_WORK_QTY, REMAIN_WORK_QTY, TARGET_WORK_QTY, ACT_EQUIP_QTY, REMAIN_EQUIP_QTY,TARGET_EQUIP_QTY " &
                    "                   , case when TARGET_DRTN_HR_CNT>0 then (100*(task.TARGET_DRTN_HR_CNT-task.REMAIN_DRTN_HR_CNT) / (task.TARGET_DRTN_HR_CNT)) else 0 end as pct_complete  " &
                    "               from  task inner join  taskactv on task.TASK_ID=taskactv.TASK_ID " &
                    "               where task.PROJ_ID= " & _projIdPM6 & " and actv_code_id= " & row("actv_code_id") & "  and wbs_id= " & t.w.id & " " &
                    "           ) a )  "

                    Try
                        Dim dtTasksCompletionInfo As DataTable = p6Conn.ExecuteQuery(tasksStmt).Tables(0)

                        If Not IsDBNull(row("ACT_WORK_QTY")) AndAlso Not _P6OpenProjWithSigepDraftProj Then : t.HrsActualWork = Convert.ToDecimal(row("ACT_WORK_QTY")) : t.ActualWork = t.HrsActualWork / hoursPerDay : End If
                        If Not IsDBNull(row("REMAIN_WORK_QTY")) Then : t.HrsRemainingWork = Convert.ToDecimal(row("REMAIN_WORK_QTY")) : t.RemainingWork = t.HrsRemainingWork / hoursPerDay : End If
                        If Not IsDBNull(row("TARGET_WORK_QTY")) Then : t.HrsTargetWork = Convert.ToDecimal(row("TARGET_WORK_QTY")) : t.TargetWork = t.HrsTargetWork / hoursPerDay : End If
                        If Not IsDBNull(row("ACT_EQUIP_QTY")) AndAlso Not _P6OpenProjWithSigepDraftProj Then : t.HrsActualEqpmQty = Convert.ToDecimal(row("ACT_EQUIP_QTY")) : t.ActualEqpmQty = t.HrsActualEqpmQty / hoursPerDay : End If
                        If Not IsDBNull(row("REMAIN_EQUIP_QTY")) Then : t.HrsRemainingEqpmQty = Convert.ToDecimal(row("REMAIN_EQUIP_QTY")) : t.RemainingEqpmQty = t.HrsRemainingEqpmQty / hoursPerDay : End If
                        If Not IsDBNull(row("TARGET_EQUIP_QTY")) Then : t.HrsTargetEqpmQty = Convert.ToDecimal(row("TARGET_EQUIP_QTY")) : t.TargetEqpmQty = t.HrsTargetEqpmQty / hoursPerDay : End If

                        If Not _P6OpenProjWithSigepDraftProj Then : t.percentComplete = dtTasksCompletionInfo.Rows(0)(0)
                        Else : t.percentComplete = 0 : End If

                        If Not _P6OpenProjWithSigepDraftProj Then : t.WorkPercComplete = t.percentComplete
                        Else : t.WorkPercComplete = 0 : End If

                        t.EqpmQtyPercComplete = t.percentComplete
                        If Not _P6OpenProjWithSigepDraftProj Then
                            t.actualDuration = CInt(dtTasksCompletionInfo.Rows(0)(2)) / hoursPerDay
                            t.remainDuration = CInt(dtTasksCompletionInfo.Rows(0)(1) - dtTasksCompletionInfo.Rows(0)(2)) / hoursPerDay
                        Else
                            t.actualDuration = 0
                            t.remainDuration = CInt(dtTasksCompletionInfo.Rows(0)(1)) / hoursPerDay
                        End If
                        't.physicalPercentComplete = CInt(dt2.Rows(0)(3))
                        If Not _P6OpenProjWithSigepDraftProj Then : t.physicalPercentComplete = dtTasksCompletionInfo.Rows(0)(3)
                        Else : t.physicalPercentComplete = 0 : End If

                        If t.percentComplete = 0 Or IsDBNull(t.percentComplete) Then t.actualEndDate = DateTime.MinValue
                    Catch ex As Exception
                        t.percentComplete = 0
                        t.actualDuration = 0
                        t.remainDuration = 0
                    End Try
                End If

                ' adding task to collection
                Me._task.Add(t)
                'If _task.Count Mod 10 = 0 Then
                '    _Caller.WriteLog("Creating Task " & _task.Count & "/" & dtTasks.Rows.Count, _projIdSigep, _projIdPM6)
                'End If
            Next
            'If Not dtTaskLinks Is Nothing Then
            '    For Each r As DataRow In dtTaskLinks.Rows
            '        Dim tlnk As New TaskLinks()
            '        tlnk.Dcid = r("lag_hr_cnt")
            '        tlnk.DcPreid = r("task_id")
            '        tlnk.DcSuccid = r("pred_task_id")
            '        tlnk.DcType = r("pred_type")
            '        tlnk.TaskUid = r("task_id")
            '        tlnk.Dclnklagfmt = 7
            '        Me._taskLinks.Add(tlnk)
            '    Next
            'End If
        Catch ex As Exception
            System.Diagnostics.Debug.Print(ex.Message)
            _LastException = ex
        Finally

        End Try
    End Sub

    Private Function CalcWbsCode(itemCode As String, itemParent As Integer, wbsRelations As Dictionary(Of Integer, Integer), wbsCodes As Dictionary(Of Integer, String)) As String

        Dim result As String = itemCode
        Dim parentCode As String = String.Empty
        Dim newParentId As Integer
        If Not wbsCodes.TryGetValue(itemParent, parentCode) Then Return result
        result = parentCode & "." & result
        If Not wbsRelations.TryGetValue(itemParent, newParentId) Then Return result
        Return CalcWbsCode(result, newParentId, wbsRelations, wbsCodes)

    End Function

    Private Sub setWbs()

        Dim wbsCodes As Dictionary(Of Integer, String)
        Dim wbsRelations As Dictionary(Of Integer, Integer)
        Dim sql As String
        Dim w As wbs
        Dim conn As PO.DataClient.IDataClient = IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)
        Dim isnull As String = DBUtilities.IsNull(IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.Primavera))
        _Caller.WriteLog("Creating WBS ", _projIdSigep, _projIdPM6)

        Try

            sql = " select proj.*, " & isnull & "(sumt.TOTAL_DRTN_HR_CNT, 0) TOTAL_DRTN_HR_CNT, " & isnull & "(sumt.REMAIN_DRTN_HR_CNT, 0) REMAIN_DRTN_HR_CNT " _
                    & " from projwbs  proj " _
                    & " left outer join sumtask sumt on proj.proj_id = sumt.proj_id and proj.wbs_id = sumt.wbs_id " _
                    & " where proj.proj_id = " & _projIdPM6 & " " _
                    & " and proj.delete_date Is null " _
                    & " order by proj.parent_wbs_id, proj.seq_num asc, proj.wbs_id "
            Dim data As DataTable = conn.ExecuteQuery(sql).Tables(0)

            ' building a dictionary that will be used to calculate the complete code for each WBS item
            wbsCodes = (From item In data.AsEnumerable()
                        Select New With
                           {
                               Key .Key = Convert.ToInt32(item("WBS_ID")),
                               Key .Val = DirectCast(item("WBS_SHORT_NAME"), String)
                           }).Distinct().AsEnumerable().ToDictionary(Function(k) k.Key, Function(v) v.Val)
            wbsRelations = (From item In data.AsEnumerable()
                            Select New With
                           {
                               Key .Key = Convert.ToInt32(item("WBS_ID")),
                               Key .Val = Convert.ToInt32(item("PARENT_WBS_ID"))
                           }).Distinct().AsEnumerable().ToDictionary(Function(k) k.Key, Function(v) v.Val)

            For Each r As DataRow In data.Rows
                w = New wbs()
                w.name = r("wbs_name")
                w.seqNumber = r("seq_num")
                If r("proj_node_flag") = "Y" Then
                    w.isRoot = True
                    w.level = 0
                    w.number = 0
                    w.parentId = Nothing
                Else
                    w.parentId = r("parent_wbs_id")
                End If
                w.id = r("wbs_id")
                w.chainName = r("wbs_short_name")
                w.remainDurationP6 = r("REMAIN_DRTN_HR_CNT")
                w.totalDurationP6 = r("TOTAL_DRTN_HR_CNT")
                w.code = CalcWbsCode(DirectCast(r("WBS_SHORT_NAME"), String), w.parentId, wbsRelations, wbsCodes)
                Me._wbs.Add(w)
            Next
            AddWbsToTasks()

            Dim cont As Integer = 0
            For Each t As task In _task
                '  nn la uso, recupero i dati grazie al ws
                '  calculateActualDurationOrPercentComplete(t)
                w = New wbs()
                w.id = t.id
                w.parentId = t.w.id
                w.nodeType = False
                w.name = t.name
                w.chainName = GetElementByID(w.parentId).level + 1
                w.forecastStartDate = t.forecaststartDate
                w.forecastEndDate = t.forecastEndDate
                w.actualStartDate = t.actualStartDate
                w.actualEndDate = t.actualEndDate
                w.baselineStartDate = t.baselineStartDate
                w.baselineEndDate = t.baselineEndDate
                w.status = t.status
                w.activityType = t.activityType
                w.code = t.code
                w.physicalPercentComplete = t.physicalPercentComplete
                w.percentComplete = t.percentComplete
                w.actualDuration = t.actualDuration
                w.remainDuration = t.remainDuration

                w.ActualEqpmQty = t.ActualEqpmQty
                w.ActualWork = t.ActualWork
                w.TargetEqpmQty = t.TargetEqpmQty
                w.TargetWork = t.TargetWork
                w.RemainingEqpmQty = t.RemainingEqpmQty
                w.RemainingWork = t.RemainingWork

                w.HrsActualEqpmQty = t.HrsActualEqpmQty
                w.HrsActualWork = t.HrsActualWork
                w.HrsTargetEqpmQty = t.HrsTargetEqpmQty
                w.HrsTargetWork = t.HrsTargetWork
                w.HrsRemainingEqpmQty = t.HrsRemainingEqpmQty
                w.HrsRemainingWork = t.HrsRemainingWork

                w.EqpmQtyPercComplete = t.EqpmQtyPercComplete
                w.WorkPercComplete = t.WorkPercComplete

                w.totalDurationP6 = t.totalDurationP6
                w.remainDurationP6 = t.remainDurationP6

                _wbs.Add(w)
                cont += 1
            Next

            Dim i As Integer
            For i = 0 To _wbs.Count - 1
                _wbs(i).level = GetLevel(_wbs(i))
            Next

            _wbs.Sort(AddressOf CompareWbs)

            If _wbs.Count > 0 Then
                SetWbsNumber(GetRoot(), 0)
            End If

            'segnaposto settaggio _wbsDatatable
            '///////////////////////////// mi creo un datatable per _wbs così da usare una ricerca indicizzata e non più sequenziale
            For Each item As wbs In _wbs
                If item.id <> 0 Then
                    '_wbsHash.Add(item.id, w)
                    _WbsHash.Add(item.id, item)
                End If
            Next
            '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            _wbs.Sort(AddressOf CompareWbsByNumber)


            SetChainName()
            SetItemPath()
            ' infine tolgo tutti i nodi di wbs che nn hanno figli di tipo task -perchè in primavera nn li fa vedere
            Dim _wbsToBeRemoved As New ArrayList
            For Each ww As wbs In _wbs
                If Not ww.hasChild(ww, _wbs) AndAlso ww.nodeType Then
                    ww.visible = False
                End If
            Next
            'fine

        Catch ex As Exception
            _LastException = ex
        End Try

    End Sub

    Private Sub LoadCustomAttributes()

        Dim conn As PO.DataClient.IDataClient = IBaseDataMgr.GetDBConnection(PO.DBInterface.DataSource.Primavera)

        Dim attrCode As P6CustomAttributes.AttributeCode
        Dim attrList As Dictionary(Of String, P6CustomAttributes.AttributeCode) = P6CustomAttributes.LoadCustomAttributesList()
        Dim taskId As Integer
        Dim attrType As String
        Dim attrValue As String
        Dim attrPair As Dictionary(Of P6CustomAttributes.AttributeCode, String)
        Dim data As DataTable
        Dim sql As String = String.Empty
        Dim customAttrFilter As String = String.Join(",", attrList.Keys.ToList().Select(Function(x) DBUtilities.QS(x.ToLower())).ToArray())

        _CustomAttributes = New Dictionary(Of Integer, Dictionary(Of P6CustomAttributes.AttributeCode, String))

        If attrList.Count = 0 Then Return


        sql = " select val.FK_ID as TASK_ID, typ.UDF_TYPE_LABEL, val.UDF_TEXT as PARAM_VALUE "
        sql &= " from UDFVALUE val, UDFTYPE typ "
        sql &= " where Val.PROJ_ID = " & _projIdPM6.ToString() & " "
        sql &= " and val.UDF_TYPE_ID = typ.UDF_TYPE_ID "
        sql &= " and typ.UDF_TYPE_LABEL in (" & customAttrFilter & ") "
        sql &= " and val.DELETE_DATE is null and typ.DELETE_DATE is null "

        data = conn.ExecuteQuery(sql).Tables(0)

        For Each row As DataRow In data.Rows
            taskId = Convert.ToInt32(row("TASK_ID"))
            attrType = DirectCast(row("UDF_TYPE_LABEL"), String)
            If IsDBNull(row("PARAM_VALUE")) Then
                attrValue = String.Empty
            Else
                attrValue = DirectCast(row("PARAM_VALUE"), String)
            End If
            attrCode = attrList(DirectCast(row("UDF_TYPE_LABEL"), String).ToLower()) ' P6CustomAttributes.EncodeAttribute(attrList(attrValue))
            If Not _CustomAttributes.TryGetValue(taskId, attrPair) Then
                _CustomAttributes(taskId) = New Dictionary(Of P6CustomAttributes.AttributeCode, String)
            End If
            _CustomAttributes(taskId)(attrCode) = P6CustomAttributes.ParseAttrValue(attrCode, attrValue)
        Next

    End Sub

    'Private Sub LoadMlstnAttr()
    '    Try
    '        _sigepprimavera.WriteLog("Retrieving milestones additional infos", _projIdSigep, _projId)

    '        Dim data As New DataTable
    '        Dim sql As String = String.Empty
    '        Dim mlstnAttr As MlstnAttributes
    '        Dim taskId As Integer
    '        Dim attrType As String
    '        Dim attrValue As String

    '        sql = " select val.FK_ID as TASK_ID, typ.UDF_TYPE_LABEL, val.UDF_TEXT as PARAM_VALUE "
    '        sql &= " from UDFVALUE val, UDFTYPE typ "
    '        sql &= " where Val.PROJ_ID = " & _projId.ToString() & " "
    '        sql &= " and val.UDF_TYPE_ID = typ.UDF_TYPE_ID "
    '        sql &= " and typ.UDF_TYPE_LABEL in ('" & _skylineFlagCol & "', '" & _slipChartFlagCol & "', '" & _skylineAttrCol & "', '" & _pptFlagCol & "') "
    '        sql &= " and val.DELETE_DATE is null and typ.DELETE_DATE is null "

    '        data = DbUtility.QueryTable(sql, datasource.pm5)

    '        For Each row As DataRow In data.Rows
    '            taskId = Convert.ToInt32(row("TASK_ID"))
    '            attrType = CStr(row("UDF_TYPE_LABEL"))
    '            If IsDBNull(row("PARAM_VALUE")) Then
    '                attrValue = String.Empty
    '            Else
    '                attrValue = CStr(row("PARAM_VALUE"))
    '            End If

    '            If Not _mlstnAdditionalAttr.TryGetValue(taskId, mlstnAttr) Then
    '                mlstnAttr = New MlstnAttributes()
    '            End If

    '            If (String.Compare(attrType, _skylineFlagCol, True) = 0) Then
    '                mlstnAttr.SetSkylineFlag(attrValue)
    '            ElseIf (String.Compare(attrType, _slipChartFlagCol, True) = 0) Then
    '                mlstnAttr.SetSlipChartFlag(attrValue)
    '            ElseIf (String.Compare(attrType, _skylineAttrCol, True) = 0) Then
    '                mlstnAttr.SetSkylineAttrFlags(attrValue)
    '            ElseIf (String.Compare(attrType, _pptFlagCol, True) = 0) Then
    '                mlstnAttr.SetPPTFlag(attrValue)
    '            End If

    '            _mlstnAdditionalAttr(taskId) = mlstnAttr
    '        Next

    '    Catch ex As Exception
    '        Console.WriteLine(ex.Message)
    '    End Try
    'End Sub

    'Public Function GetTaskInWbsNode(ByVal w As wbs) As ArrayList
    '    Dim ret As New ArrayList
    '    For Each t As task In _task
    '        If t.w.id = w.id Then
    '            ret.Add(t)
    '        End If
    '    Next
    '    Return ret
    'End Function

    Public Sub SetChainName()
        For Each w As wbs In _wbs
            Dim tmp As String = String.Empty
            If Not GetWBSNodeName(w.parentId) Is Nothing Then
                Dim cont4 As Integer = 1
                While ExistName(GetWBSNodeName(w.parentId) & "." & cont4)
                    cont4 += 1
                End While
                w.chainName = GetWBSNodeName(w.parentId) & "." & cont4
            Else
                w.chainName = w.level
            End If
            _ChainNames.Add(w.chainName, String.Empty)
        Next
    End Sub
    Public Sub SetItemPath()

        For Each w As wbs In _wbs
            Dim result As List(Of String) = Nothing
            w.ItemPath.AddRange(GetParentIDs(w, result))
        Next

    End Sub
    Private Function GetParentIDs(obj As wbs, ByRef result As List(Of String)) As List(Of String)

        If result Is Nothing Then result = New List(Of String)
        If obj.isRoot Then
            result.Add(_projIdPM6)
            Return result
        Else
            result.Add(obj.id)
            Dim parentObj As wbs = _wbs.FirstOrDefault(Function(x) x.id = obj.parentId)
            If Not parentObj Is Nothing Then result = GetParentIDs(parentObj, result)
        End If
        Return result

    End Function
    Public Function ExistName(ByVal name As String) As Boolean
        Return _ChainNames.ContainsKey(name)
        'For Each ww As wbs In _wbs
        '    If ww.chainName = name Then
        '        Return True
        '    End If
        'Next
        'Return False
    End Function

    'Public Sub setSummaryDateAndPercComplete()
    '    Dim cont As Integer
    '    Dim maxLevel As Integer
    '    Dim childs As ArrayList
    '    maxLevel = getMaxWbsLevel()
    '    For cont = maxLevel To 0 Step -1
    '        For Each ww As wbs In _wbs
    '            If ww.level = cont Then
    '                childs = getChild(ww)
    '                If ww.nodeType = True And ww.visible = True Then  'se  un nodo di wbs
    '                    ww.forecastStartDate = getForecastDate(childs, 0)
    '                    ww.forecastEndDate = getForecastDate(childs, 1)
    '                    ww.actualStartDate = getActualDate(childs, 0)
    '                    ww.actualEndDate = getActualDate(childs, 1)
    '                    ww.baselineStartDate = getBaselineDate(childs, 0)
    '                    ww.baselineEndDate = getBaselineDate(childs, 1)
    '                    '///////// questo per il calcolo delle perc_complete dei nodi wbs
    '                    Dim percComplTmp As New ArrayList()
    '                    Dim percComplTot As Decimal = 0
    '                    Dim taskInWbsNode As New ArrayList
    '                    Dim durationTot As Decimal = 0
    '                    Dim ActualDurationTot As Decimal = 0
    '                    Dim RemainDurationTot As Decimal = 0
    '                    For Each www As wbs In childs
    '                        durationTot += www.remainDuration + www.actualDuration
    '                    Next
    '                    'una volta calcolata la duration totale del nodo posso calcoloare la per_compl del nodo facendo la somma pesata dei task che lo compongono
    '                    For Each www As wbs In childs
    '                        ActualDurationTot += www.actualDuration
    '                        RemainDurationTot += www.remainDuration
    '                        If durationTot = 0 Then
    '                            percComplTot = 0
    '                        Else
    '                            percComplTot += www.percentComplete * (www.remainDuration + www.actualDuration) / durationTot
    '                        End If
    '                    Next
    '                    ww.percentComplete = percComplTot
    '                    ww.actualDuration = ActualDurationTot
    '                    ww.remainDuration = RemainDurationTot
    '                    '///////////////////////////////////////////////////////////////////////////////////////////////

    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub

    'Public Sub groupByActivityCode()
    '    Dim cont As Integer
    '    Dim maxLevel As Integer
    '    Dim childs As ArrayList
    '    maxLevel = getMaxWbsLevel()
    '    For cont = maxLevel To 0 Step -1
    '        For Each ww As wbs In _wbs
    '            If ww.level = cont Then
    '                childs = getChild(ww)
    '                If ww.nodeType = False Then   'se  un nodo  task
    '                    ww.forecastStartDate = getForecastDate(childs, 0)
    '                    ww.forecastEndDate = getForecastDate(childs, 1)
    '                End If
    '            End If
    '        Next
    '    Next
    'End Sub

    'Private Function getMax(ByVal i As Integer)
    '    Dim j As Integer = 0
    '    Dim max = 0
    '    For j = 0 To i
    '        If _task(j).w.number > max Then
    '            max = _task(j).w.number + 1
    '        End If
    '    Next
    '    Return max
    'End Function
    Private Function GetWBSNodeName(ByVal idToFind As Integer) As String
        If Not CType(_WbsHash.Item(idToFind), wbs) Is Nothing Then
            Return CType(_WbsHash.Item(idToFind), wbs).chainName
        Else
            Return Nothing
        End If

        'Return getWBSNodeName2(idToFind)
        'Dim i As Integer = 0
        'For i = 0 To _wbs.Count - 1
        '    If _wbs(i).id = idToFind Then
        '        Return _wbs(i).chainName
        '    End If
        'Next
        'Return Nothing
    End Function

    'Private Function getWBSNodeName2(ByVal idToFind As Integer) As String
    '    'Dim dr() As DataRow
    '    'dr = _wbsDataTable.Select("id = " & idToFind)
    '    'If dr.Length > 0 Then
    '    '    Return dr(0)("chainname")
    '    'Else
    '    '    Return Nothing
    '    'End If
    '    If Not CType(_wbsHash.Item(idToFind), wbs) Is Nothing Then
    '        Return CType(_wbsHash.Item(idToFind), wbs).chainName
    '    Else
    '        Return Nothing
    '    End If
    'End Function

    Public Function DeleteTaskInSigep() As Boolean
        Try
            _Caller._DMLDBConn.ExecuteNonQuery("delete from msp_tasks where  proj_id= " & _projIdSigep)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            Return True
        Catch ex As Exception
            _LastException = ex
            Return False
        End Try
    End Function

    'Public Function deleteMspTask(ByVal projId As Integer, ByVal taskUid As Integer) As Boolean
    '    Dim stmt As String
    '    Dim transaction As OleDbTransaction = Nothing

    '    Try
    '        transaction = _Caller._CnSigep.BeginTransaction

    '        stmt = "delete from PROG_TRACK_T091 where i_act_id = " & taskUid.ToString() & " and C_PROG = " & projId.ToString() & " "
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)

    '        stmt = " delete from tab_cost_work_phasing where ass_id in (select ass_id from cost_rev_asso_task where proj_id = " & projId & " and task_uid = " & taskUid & ")"
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)
    '        stmt = " delete from tab_cost_baseline_phasing where ass_id in (select ass_id from cost_rev_asso_task where proj_id = " & projId & " and task_uid = " & taskUid & ")"
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)
    '        stmt = " delete from tab_cost_actual_phasing where ass_id in (select ass_id from cost_rev_asso_task where proj_id = " & projId & " and task_uid = " & taskUid & ")"
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)

    '        stmt = " delete from cost_rev_asso_task where proj_id = " & projId & " and task_uid = " & taskUid & " "
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)
    '        stmt = " delete from po_task_pbs_codes where proj_id = " & projId & " and task_uid = " & taskUid & " "
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)
    '        stmt = " delete from msp_task_assignment where proj_id = " & projId & " and task_uid = " & taskUid & " "
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)

    '        stmt = " delete from msp_tasks where proj_id = " & projId & " and task_uid = " & taskUid & " "
    '        DbUtility.QueryExec(stmt, datasource.sigep, _Caller._DmlSigepCmd)

    '        transaction.Commit()
    '        Return True
    '    Catch ex As Exception
    '        transaction.Rollback()
    '        Return False
    '    End Try
    'End Function
    Public Function InsertAllWbsInSigep() As Boolean

        Try

            Dim cont As Integer = 0
            Dim i As Integer = 0
            Dim sql As String = String.Empty
            Dim taskExists As Boolean
            Dim existingTasks As New SortedDictionary(Of Integer, Integer)
            'Dim ExistingTasksList As DataTable, P6DeletedTasks As DataTable
            Dim deletedTasksList As String = String.Empty

            'inserisco la task_dur dei nodi di tipo summary
            Dim countAssign As Integer = 0

            'verifico se posso usare la tabella tasksum 
            Dim tasksum As Boolean = True
            'non voglio usare la tasksum
            tasksum = False

            _Caller.WriteLog("Checking Changed Tasks ", _projIdSigep, _projIdPM6)

            ' recreate the special tasks for MSP
            _Caller._DMLDBConn.ExecuteNonQuery("delete from msp_tasks where task_uid in (0, -65534, -65535, -65536) and proj_id=" & _projIdSigep)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            'Dim projname As String = String.Empty 'DbUtility.QueryValue("select s_prog_nom from prog_t056 where c_prog = " & _projIdSigep, datasource.sigep)
            '_Caller._DMLDBConn.ExecuteScalar("select s_prog_nom from prog_t056 where c_prog = " & _projIdSigep, projname)
            'If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            sql = "Insert into MSP_TASKS (reserved_data,proj_id,task_uid) values (0," & _projIdSigep & ",-65534)"
            _Caller._DMLDBConn.ExecuteNonQuery(sql)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            sql = "Insert into MSP_TASKS (reserved_data,proj_id,task_uid) values (0," & _projIdSigep & ",-65535)"
            _Caller._DMLDBConn.ExecuteNonQuery(sql)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            sql = "Insert into MSP_TASKS (reserved_data,proj_id,task_uid) values (0," & _projIdSigep & ",-65536)"
            _Caller._DMLDBConn.ExecuteNonQuery(sql)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            sql = "Insert into MSP_TASKS (task_name,task_id,task_uid,proj_id) values ('" & _SigepProjNameAndCompany & "',0,0," & _projIdSigep & " )"
            _Caller._DMLDBConn.ExecuteNonQuery(sql)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            Dim allWbsItems As List(Of wbs) = (From item In _wbs Select item Order By item.level Descending).ToList()
            For Each item As wbs In allWbsItems
                SetSummaryDateAndPercComplete(item)
            Next

            For i = 0 To _wbs.Count - 1
                'If i Mod 10 = 0 Then
                '    _Caller.WriteLog("Loading Task " & i & "/" & _wbs.Count - 1, _projIdSigep, _projIdPM6)
                'End If
                _Counter = i
                If _wbs(i).visible Then 'And (_wbs(i).activityType <> [Enum].GetName(GetType(taskType), 1) And _wbs(i).activityType <> [Enum].GetName(GetType(taskType), 2)) Then

                    Try
                        'SetSummaryDateAndPercComplete(_wbs(i))
                        taskExists = existingTasks.ContainsKey(_wbs(i).id)
                        If Not InsertWbsInSigep(i, taskExists, _SigepProjStatus) Then Return False
                    Catch
                    End Try

                End If
            Next
            _Caller._DMLDBConn.ExecuteNonQuery("delete from msp_links where  proj_id=" & _projIdSigep)
            'SaveTaskLink(_taskLinks)
            _Counter = -1
            Return True
        Catch ex3 As Exception
            _LastException = ex3
            Return False
        Finally
        End Try
    End Function


    'Public Function insertAllMilestoneInSigep() As Boolean
    '    Try
    '        Dim owner As String = String.Empty
    '        Try
    '            owner = DbUtility.QueryTable("select s_prog_mgr  from prog_t056 where c_prog=" & _projIdSigep, datasource.sigep).Rows(0)(0)
    '            If owner = "" Then owner = "0"
    '        Catch ex As Exception
    '            owner = "0"
    '        End Try
    '        '///////////////////////// caso di aggregazione //////////////////////////////
    '        If groupingMode = True Then
    '            _task.Clear()
    '            Dim cont As Integer
    '            cont = 0
    '            setTask(True) '"select * from task a where a.PROJ_ID=" & _projId & "  and task_type <>'TT_Task'"
    '            For Each t As task In _task
    '                Dim w As New wbs()
    '                w.id = 9000 + cont
    '                w.parentId = t.w.id
    '                w.nodeType = False
    '                w.name = t.name
    '                w.chainName = getElementByID2(w.parentId).level + 1
    '                w.forecastStartDate = t.forecaststartDate
    '                w.forecastEndDate = t.forecastEndDate
    '                w.actualStartDate = t.actualStartDate
    '                w.actualEndDate = t.actualEndDate
    '                w.baselineStartDate = t.baselineStartDate
    '                w.baselineEndDate = t.baselineEndDate
    '                w.status = t.status
    '                w.activityType = t.activityType
    '                w.code = t.code
    '                w.remainDurationP6 = t.remainDurationP6
    '                w.totalDurationP6 = t.totalDurationP6
    '                _wbs.Add(w)
    '                cont += 1
    '            Next
    '        End If
    '        '///////////////////////// fine caso di aggregazione //////////////////////////////


    '        Dim m As New milestone(_sigepprimavera._CmdSigep)
    '        Dim attr As Dictionary(Of P6CustomAttributes.AttributeCode, String)

    '        m.deleteAll(_projIdSigep)

    '        For Each w As wbs In _wbs
    '            If w.nodeType = False AndAlso w.activityType = [Enum].GetName(GetType(taskType), 1) Or w.activityType = [Enum].GetName(GetType(taskType), 2) Then

    '                If _CustomAttributes.TryGetValue(w.id, attr) Then
    '                    Dim value As String
    '                    If attr.TryGetValue(P6CustomAttributes.AttributeCode.InProgressFlag, value) Then m.FLG_ONGOING = IIf(value = "1", "Y", "N")
    '                    If attr.TryGetValue(P6CustomAttributes.AttributeCode.ToBeWatchedFlag, value) Then m.FLG_TOBEWATCHED = IIf(value = "1", "Y", "N")
    '                    If attr.TryGetValue(P6CustomAttributes.AttributeCode.SlipChartFlag, value) Then m.FLG_SLIP_CHART = IIf(value = "1", "Y", "N")
    '                    If attr.TryGetValue(P6CustomAttributes.AttributeCode.SkylineFlag, value) Then m.FLG_SKYLINE = IIf(value = "1", "Y", "N")
    '                    If attr.TryGetValue(P6CustomAttributes.AttributeCode.PPTFlag, value) Then m.FLG_POWERPT = IIf(value = "1", "Y", "N")
    '                End If

    '                m.C_PROG = _projIdSigep
    '                m.C_MLSTN = w.code
    '                m.S_MLSTN = w.name
    '                m.C_TIP_MLSTN = "PM6"
    '                m.D_PLANNED = w.baselineStartDate
    '                m.D_ACTUAL = w.actualStartDate
    '                m.D_FORECAST = w.forecastStartDate
    '                m.C_MSP_TASK = w.getIndex(w.id, _wbs) + 1
    '                m.F_LINK = "F"
    '                m.V_LAG = 0
    '                m.F_MSP_TASK_STATUS = IIf(w.status.ToLower.IndexOf("complete") >= 0, "C", "P") 'in primavera usano 2 stati per le milestone TK_Complete e TK_NotStart in sigep ce ne sono 3 "c" "p" e "i"!!!
    '                m.F_CRITICALY = String.Empty
    '                m.V_DELAY = DateDiff(DateInterval.Day, w.forecastStartDate, w.actualStartDate)
    '                If w.actualStartDate.Year = 1 Then m.V_DELAY = 0
    '                m.C_OWNER = CDec(owner)
    '                m.DESCRIPTION = String.Empty
    '                m.BENEFIT = String.Empty
    '                m.create()
    '            End If
    '        Next
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function

    Private Function InsertWbsInSigep(ByVal index As Integer, ByVal taskExists As Boolean, ByVal projStatus As Integer) As Boolean

        Dim actDuration As Integer
        Dim stmt As String
        Dim t As wbs = _wbs(index)
        '22 luglio 2009
        'se il progetto sigep non è in stato progress non inserisco le date di actual la  phy complete e la % complete
        If projStatus <> 2 Then
            t.actualDuration = 0
            t.actualStartDate = Nothing
            t.actualEndDate = Nothing
            t.percentComplete = 0
            t.EqpmQtyPercComplete = 0
            t.WorkPercComplete = 0
            t.physicalPercentComplete = 0
        End If

        If _GroupByActivityCode > 0 Then t.physicalPercentComplete = 0
        'fine-----------------

        actDuration = t.actualDuration

        Dim forecaststartDate = "NULL"
        Dim forecastEndDate = "NULL"
        Dim actualStartDate = "NULL"
        Dim actualEndDate = "NULL"
        Dim task_is_milestone As Integer = 0

        'Dim attr As MlstnAttributes = Nothing
        Dim attr As Dictionary(Of P6CustomAttributes.AttributeCode, String) = Nothing

        If t.nodeType = False AndAlso t.activityType = [Enum].GetName(GetType(taskType), 1) Or t.activityType = [Enum].GetName(GetType(taskType), 2) Then
            task_is_milestone = 1
            '_mlstnAdditionalAttr.TryGetValue(t.id, attr)
            '_CustomAttributes.TryGetValue(t.id, attr)
        End If
        _CustomAttributes.TryGetValue(t.id, attr)

        Try
            If Not taskExists Then
                stmt = InsertTaskStmt(t, task_is_milestone, attr, index, actDuration)
            Else
                stmt = UpdateTaskStmt(t, task_is_milestone, attr, index, actDuration)
            End If
            _Caller._DMLDBConn.ExecuteNonQuery(stmt)

            'Dim dc As List(Of TaskLinks) = _taskLinks.FindAll(Function(x) x.DcPreid = t.id)
            'If Not IsNothing(dc) Then
            '    For Each prId As TaskLinks In dc
            '        prId.DcPreid = t.number
            '        prId.TaskUid = t.id
            '    Next
            'End If
            'Dim dcs As List(Of TaskLinks) = _taskLinks.FindAll(Function(x) x.DcSuccid = t.id)
            'If Not IsNothing(dcs) Then
            '    For Each pdId As TaskLinks In dcs
            '        pdId.DcSuccid = t.number
            '        pdId.TaskUid = t.id
            '    Next
            'End If

            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            ' update PT tasks if any
            stmt = UpdatePTTaskStmt(t, attr)
            _Caller._DMLDBConn.ExecuteNonQuery(stmt)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            ' Threading.Thread.Sleep(2000)
            'If index Mod 10 = 0 Then
            '    _Caller.CheckAbort(_projIdSigep, _projIdPM6)
            'End If

            Return True
        Catch ex As Exception
            _LastException = ex
            Return False
        End Try
    End Function
    'EXCLUDED AS WORK IS NOT COMPLETE YET 
    'Save link on Sigep Table
    'Private Sub SaveTaskLink(_taskLinks As List(Of TaskLinks))
    '    Dim stmt As String = ""
    '    Dim i As Integer
    '    Dim _data As String = "0"
    '    For i = 0 To _taskLinks.Count - 1
    '        stmt = " insert into msp_links(RESERVED_DATA,PROJ_ID,EXT_EDIT_REF_DATA,LINK_UID,LINK_IS_CROSS_PROJ,LINK_PRED_UID,LINK_SUCC_UID,"
    '        stmt &= "LINK_TYPE,LINK_LAG_FMT,LINK_LAG)values("
    '        stmt &= "0," & _projIdSigep & ",NULL,"
    '        stmt &= _taskLinks(i).TaskUid & ",0," & _taskLinks(i).DcPreid & ","
    '        stmt &= _taskLinks(i).DcSuccid & "," & _taskLinks(i).DcType & "," & _taskLinks(i).Dclnklagfmt & ",'" & _taskLinks(i).DcLag
    '        stmt &= "')"
    '        _Caller._DMLDBConn.ExecuteNonQuery(stmt)
    '    Next
    'End Sub
    '''' <summary>
    '''' just for test
    '''' </summary>
    '''' <param name="w"></param>
    '''' <returns></returns>
    'Public Sub insertWbsInSigep(cmd As OleDbCommand, ByVal index As Integer, ByVal taskExists As Boolean, ByVal projStatus As Integer)
    '    Dim SqlInsert As String = ""
    '    Dim t As wbs = _wbs(index)

    '    Dim actDuration As Integer

    '    '22 luglio 2009
    '    'se il progetto sigep non è in stato progress non inserisco le date di actual la  phy complete e la % complete
    '    If projStatus <> 2 Then
    '        t.actualDuration = 0
    '        t.actualStartDate = Nothing
    '        t.actualEndDate = Nothing
    '        t.percentComplete = 0
    '        t.EqpmQtyPercComplete = 0
    '        t.WorkPercComplete = 0
    '        t.physicalPercentComplete = 0
    '    End If

    '    If groupingMode Then
    '        t.physicalPercentComplete = 0
    '    End If
    '    'fine-----------------

    '    actDuration = t.actualDuration

    '    Dim forecaststartDate = "NULL"
    '    Dim forecastEndDate = "NULL"
    '    Dim actualStartDate = "NULL"
    '    Dim actualEndDate = "NULL"
    '    Dim task_is_milestone As Integer = 0
    '    Dim attr As Dictionary(Of P6CustomAttributes.AttributeCode, String) = Nothing

    '    If t.nodeType = False AndAlso t.activityType = [Enum].GetName(GetType(taskType), 1) Or t.activityType = [Enum].GetName(GetType(taskType), 2) Then
    '        task_is_milestone = 1
    '    End If
    '    _CustomAttributes.TryGetValue(t.id, attr)

    '    If Not taskExists Then
    '        SqlInsert = InsertTask(t, task_is_milestone, attr, index, actDuration)
    '        DbUtility.QueryExec(SqlInsert, cmd)
    '    Else
    '        'altrimenti lo aggiorno
    '        SqlInsert = updateTask(t, task_is_milestone, attr, index, actDuration)
    '        DbUtility.QueryExec(SqlInsert, cmd)
    '    End If
    '    ' update PT tasks if any
    '    SqlInsert = updatePTTask(t, attr)
    '    DbUtility.QueryExec(SqlInsert, cmd)

    '    ' Threading.Thread.Sleep(2000)
    '    If index Mod 10 = 0 Then
    '        _sigepprimavera.checkAbort(_projIdSigep, _projId)
    '    End If
    'End Sub
    Public Function GetLevel(ByVal w As wbs) As Integer
        Dim ret As Integer
        If w.isRoot = True Then
            ret = 1
        Else
            ret = GetLevel(GetParent(w)) + 1
        End If
        Return ret
    End Function

#Region "utility"
    Public Function GetParent(ByVal w As wbs) As wbs
        Dim cont As Integer
        For cont = 0 To _wbs.Count - 1
            If w.parentId = _wbs(cont).id Then
                Return _wbs(cont)
            End If
        Next
        Return GetRoot()
    End Function
    Public Function GetRoot() As wbs
        Dim cont As Integer
        For cont = 0 To _wbs.Count - 1
            If _wbs(cont).isRoot = True Then
                Return _wbs(cont)
            End If
        Next
        Return Nothing
    End Function

    'Public Sub setWbsNumber()
    '    Dim cont As Integer = 0
    '    Dim i As Integer
    '    Dim min As Integer = 0
    '    Dim max As Integer = 0
    '    For i = 1 To getMaxWbsLevel()
    '        Dim j As Integer = 1
    '        Dim l = wbsOfLevel("=", i)
    '        Dim lPrec = wbsOfLevel("<", i)
    '        min = lPrec
    '        '  max = lPrec + l
    '        For Each ww As wbs In _wbs
    '            If ww.level = i Then
    '                ww.number = min + j
    '                j += 1
    '            End If
    '        Next
    '    Next
    'End Sub
    Public Sub AddWbsToTasks()
        For Each t As task In _task
            For Each w As wbs In _wbs
                If w.id = t.w.id Then
                    t.w.level = w.level
                    t.w.number = w.number
                    t.w.chainName = w.chainName
                    t.w.name = w.name
                    t.w.parentId = w.parentId
                    Exit For
                End If
            Next
        Next
        Dim shift As Integer = 0
        Dim taskArray(_task.Count - 1) As task
        Dim i As Integer
        For i = 0 To _task.Count - 1
            taskArray(i) = _task(i)
        Next
        Array.Sort(taskArray, New MyComparer)
        For i = 0 To _task.Count - 1
            _task(i) = taskArray(i)
        Next
        ' Dim _wbsToAdd As New ArrayList()
    End Sub

    'Public Function wbsOfLevel(ByVal op As String, ByVal l As Integer) As Integer
    '    Dim cont As Integer = 0
    '    For Each w As wbs In _wbs
    '        If op = "=" Then
    '            If w.level = l Then
    '                cont += 1
    '            End If
    '        ElseIf op = "<" Then
    '            If w.level < l Then
    '                cont += 1
    '            End If
    '        End If
    '    Next
    '    Return cont
    'End Function

    Public Function GetChainName(ByVal w As wbs) As String

        Dim ret As String
        If w.isRoot = True Then
            ret = _wbsPrefix
        Else
            '    getLevel(getParent(w))
            Dim tmp = _wbsPrefix & "."
            ret = GetChainName(GetParent(w)) & "." & w.chainName

        End If
        Return ret

    End Function

    Public Sub SetWbsNumber(ByVal ww As wbs, ByRef contt As Integer)
        If ww.number = 0 Then
            contt = contt + 1
            ww.number = contt
            Dim i As Integer
            For i = 0 To GetChild(ww).Count - 1
                SetWbsNumber(GetChild(ww)(i), contt)
            Next
        End If
    End Sub

    Public Function GetMaxWbsLevel() As Integer
        Dim max As Integer = 0
        For Each w As wbs In _wbs
            If w.level > max Then
                max = w.level
            End If
        Next
        Return max
    End Function

    Public Function GetChild(ByVal w As wbs) As ArrayList
        'SIGEPPRIMAVERA.SigepPrimaveraApp.WriteLog("getChildNode of  :" & w.id)
        _IdToFind = w.id
        Return New ArrayList(_wbs.FindAll(AddressOf IsChildPredicate))
    End Function

    'Public Function getBaselineDate(ByVal wbs As ArrayList, ByVal mode As Integer) As Date
    '    Dim i As Integer
    '    Dim min As Date
    '    Dim max As Date
    '    Dim j As Integer = 0
    '    If mode = 0 Then 'ritorna la start date
    '        While min = "#12:00:00 AM#" And j <= wbs.Count - 1
    '            min = CType(wbs.Item(j), wbs).baselineStartDate
    '            j += 1
    '        End While
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).baselineStartDate < min And CType(wbs.Item(i), wbs).baselineStartDate <> "#12:00:00 AM#" Then '#12:00:00 AM#
    '                    min = CType(wbs.Item(i), wbs).baselineStartDate
    '                End If
    '            End If
    '        Next
    '        Return min

    '    Else 'ritorna la end date
    '        max = CType(wbs.Item(0), wbs).baselineEndDate
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).baselineEndDate > max Then
    '                    max = CType(wbs.Item(i), wbs).baselineEndDate
    '                End If
    '            End If
    '        Next
    '        Return max
    '    End If
    'End Function

    'Public Function getForecastDate(ByVal wbs As ArrayList, ByVal mode As Integer) As Date
    '    Dim i As Integer
    '    Dim min As Date
    '    Dim max As Date
    '    Dim j As Integer = 0
    '    If mode = 0 Then 'ritorna la start date
    '        While min = "#12:00:00 AM#" And j <= wbs.Count - 1
    '            min = CType(wbs.Item(j), wbs).forecastStartDate
    '            j += 1
    '        End While
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).forecastStartDate < min And CType(wbs.Item(i), wbs).forecastStartDate <> "#12:00:00 AM#" Then '#12:00:00 AM#
    '                    min = CType(wbs.Item(i), wbs).forecastStartDate
    '                End If
    '            End If
    '        Next
    '        Return min

    '    Else 'ritorna la end date
    '        max = CType(wbs.Item(0), wbs).forecastEndDate
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).forecastEndDate > max Then
    '                    max = CType(wbs.Item(i), wbs).forecastEndDate
    '                End If
    '            End If
    '        Next
    '        Return max
    '    End If
    'End Function

    'Public Function getActualDate(ByVal wbs As ArrayList, ByVal mode As Integer) As Date
    '    Dim i As Integer
    '    Dim min As Date
    '    Dim max As Date
    '    Dim j As Integer = 0
    '    If mode = 0 Then 'ritorna la start date
    '        While min = "#12:00:00 AM#" And j <= wbs.Count - 1
    '            min = CType(wbs.Item(j), wbs).actualStartDate
    '            j += 1
    '        End While
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).actualStartDate < min And CType(wbs.Item(i), wbs).actualStartDate <> "#12:00:00 AM#" Then '#12:00:00 AM#
    '                    min = CType(wbs.Item(i), wbs).actualStartDate
    '                End If
    '            End If
    '        Next
    '        Return min

    '    Else 'ritorna la end date
    '        max = CType(wbs.Item(0), wbs).actualEndDate
    '        For i = 0 To wbs.Count - 1
    '            If CType(wbs.Item(i), wbs).visible = True Then
    '                If CType(wbs.Item(i), wbs).actualEndDate > max Then
    '                    max = CType(wbs.Item(i), wbs).actualEndDate
    '                End If
    '            End If
    '        Next
    '        Return max
    '    End If
    'End Function

    Public Function GetElementByID(ByVal idval As Integer) As wbs
        For Each ww As wbs In _wbs
            If ww.id = idval Then
                Return ww
            End If
        Next
        Return Nothing
    End Function

    'Public Function getElementByID2(ByVal idval As Integer) As wbs
    '    idToFind = idval
    '    Dim retId As Integer
    '    retId = _wbs.FindIndex(AddressOf EqualIdPredicate)
    '    Return _wbs(retId)
    'End Function

    'Private Function EqualIdPredicate(ByVal w As wbs) As Boolean
    '    If (w.id = idToFind) Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

    Private Function IsChildPredicate(ByVal w As wbs) As Boolean
        If (w.parentId = _IdToFind) Then
            Return True
        Else
            Return False
        End If
    End Function

    'Private Function IsChildPredicateTask(ByVal t As task) As Boolean
    '    If (t.w.id = idToFind) Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

    Public Function GetMinWbs() As Integer
        Dim minBoolean As Integer
        If _task.Count > 0 Then
            minBoolean = _task(0).w.level
            For Each t As task In _task
                If t.w.level < minBoolean Then
                    minBoolean = t.w.level
                    Exit For
                End If
            Next
        Else
            minBoolean = 0
        End If
        Return minBoolean
    End Function
#End Region
    Public Function ImportWBSInSigep() As Boolean

        Try
            Dim Sql As String
            Dim cstru As String = "WBSP6"

            'elimino la wbs esistente
            _Caller._DMLDBConn.ExecuteNonQuery(" delete from DEFN_STRU_T012 where c_prog = " & _projIdSigep & " and c_stru=" & DBUtilities.QS(cstru))
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            _Caller._DMLDBConn.ExecuteNonQuery(" delete from defn_lvl_stru_t014 where c_prog = " & _projIdSigep & " and c_stru=" & DBUtilities.QS(cstru))
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
            _Caller._DMLDBConn.ExecuteNonQuery(" delete from stru_t087 where c_prog = " & _projIdSigep & " and c_stru=" & DBUtilities.QS(cstru))
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            'creo anagrafica della WBS
            Sql = String.Format("insert into DEFN_STRU_T012 ( C_PROG, C_STRU, S_STRU, C_STRU_SIST, T_STRU, C_AZD) values ('{0}','{1}','{2}','','W','{3}')", _projIdSigep, cstru, "WBS  Imported from Primavera", _SigepProjCompany)
            _Caller._DMLDBConn.ExecuteNonQuery(Sql)
            If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException

            'creo i livelli della struttura nuova
            For i As Integer = 0 To GetMaxWbsLevel()
                'If w.nodeType = True Then
                Sql = String.Format("insert into DEFN_LVL_STRU_T014  (C_PROG, C_STRU, P_LVL, S_LVL,  C_AZD) values ('{0}','{1}','{2}','{3}','{4}')", _projIdSigep, cstru, i, i, _SigepProjCompany)
                'FileIO.FileSystem.WriteAllText("C:\test.txt", Sql, True)
                _Caller._DMLDBConn.ExecuteNonQuery(Sql)
                If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
                'End If
            Next

            'inserisco i task, dopo averli ordinati (level\parent\id)
            _wbs.Sort(AddressOf CompareWbsByLevel)
            For Each w As wbs In _wbs
                If w.nodeType = True Then
                    Sql = String.Format("insert into STRU_T087 (C_PROG, P_ELMNT, P_ELMNT_PDR, C_STRU, C_ELMNT, P_LVL, S_ELMNT,  I_PESO,I_AV,S_WEIGHT_CRITERIA,  C_AZD) " &
                                            " values ({0},{1},{2},'{3}','{4}','{5}','{6}',{7},{8},{9},{10})", _projIdSigep, w.id _
                                            , IIf(w.isRoot, "NULL", w.parentId.ToString.Trim), cstru.Replace("'", "''").Trim() _
                                            , w.chainName.Trim(), w.level, w.name.Replace("'", "''").Trim(), 0, 0, 2, _SigepProjCompany)
                    _Caller._DMLDBConn.ExecuteNonQuery(Sql)
                    If Not _Caller._DMLDBConn.LastException Is Nothing Then Throw _Caller._DMLDBConn.LastException
                End If
            Next

            Return True
        Catch ex As Exception

            Return False
        End Try

    End Function
    ' getBaselineDate restituisce la start e end baseline date di un task
    'Public Function getBaselineDate(ByVal idtask As Integer) As DataTable
    '    Dim sql As String = "SELECT T2.EARLY_START_DATE, T2.EARLY_END_DATE  FROM TASK  T1 INNER JOIN PROJECT  P ON T1.PROJ_ID = P.PROJ_ID INNER JOIN TASK  T2 ON P.SUM_BASE_PROJ_ID = T2.PROJ_ID AND T1.TASK_CODE = T2.TASK_CODE WHERE T1.TASK_ID =" & idtask
    '    Dim dt As New DataTable
    '    dt = DbUtility.QueryTable(sql, datasource.pm5)
    '    Return dt
    'End Function

    'Public Function updateProjectDates() As Boolean
    '    Try
    '        Dim sql As String
    '        Dim plannedStart, plannedEnd, forecastStart, forecastEnd As Date
    '        sql = " select plan_start_date, plan_end_date, fcst_start_date,scd_end_date from project where proj_id = " & _projIdPM6
    '        Dim dt As New DataTable
    '        dt = DbUtility.QueryTable(sql, datasource.pm5)
    '        If Not IsDBNull(dt.Rows(0)(0)) Then
    '            plannedStart = dt.Rows(0)(0)
    '            sql = " update prog_t056 set  d_ini_pnf = " & DbUtility.QD(plannedStart) & "   where c_prog = " & _projIdSigep
    '            DbUtility.QueryExec(sql, datasource.sigep)
    '        End If
    '        If Not IsDBNull(dt.Rows(0)(1)) Then
    '            plannedEnd = dt.Rows(0)(1)
    '            sql = " update prog_t056 set  d_fin_pnf = " & DbUtility.QD(plannedEnd) & "   where c_prog = " & _projIdSigep
    '            DbUtility.QueryExec(sql, datasource.sigep)
    '        End If
    '        If Not IsDBNull(dt.Rows(0)(2)) Then
    '            forecastStart = dt.Rows(0)(2)
    '            sql = " update prog_t056 set  d_ini_frc = " & DbUtility.QD(forecastStart) & "   where c_prog = " & _projIdSigep
    '        End If
    '        If Not IsDBNull(dt.Rows(0)(3)) Then
    '            forecastEnd = dt.Rows(0)(3)
    '            sql = " update prog_t056 set  d_fin_rev = " & DbUtility.QD(forecastEnd) & "   where c_prog = " & _projIdSigep
    '        End If
    '        '   sql = " update prog_t056 set  d_ini_pnf = " & DbUtility.QD(plannedStart) & " , d_fin_pnf = " & DbUtility.QD(plannedEnd) & " , d_ini_frc= " & DbUtility.QD(forecastStart) & " , d_fin_frc = " & DbUtility.QD(forecastEnd) & " ,  d_ini_rev = " & DbUtility.QD(plannedStart) & " , d_fin_rev = " & DbUtility.QD(plannedEnd) & "  where c_prog = " & _projIdSigep
    '        Dim i As Integer
    '        i = DbUtility.QueryExec(sql, datasource.sigep)
    '        dt = Nothing
    '        dt = DbUtility.QueryTable("select  WBS_SHORT_NAME, WBS_NAME  from projwbs where proj_id = " & _projIdPM6 & " and wbs_id= (select min (wbs_id) from projwbs where proj_id =" & _projIdPM6 & ")", datasource.pm5)
    '        Dim projname, projdesc As String
    '        projname = dt.Rows(0)(0)
    '        projdesc = dt.Rows(0)(1)
    '        sql = "update prog_t056 set  S_PROG_AZD_NOM = '" & projname.ToUpper & "' ,  s_prog = '" & projdesc & "' where c_prog= " & _projIdSigep
    '        DbUtility.QueryExec(sql, datasource.sigep)
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try

    'End Function
    Private Shared Function CompareWbsByLevel(ByVal x As wbs, ByVal y As wbs) As Integer
        Dim result As Integer = x.level.CompareTo(y.level)
        If result = 0 Then
            result = x.parentId.CompareTo(y.parentId)
            If result = 0 Then
                Return x.id.CompareTo(y.id)
            Else
                Return result
            End If
        Else
            Return result
        End If

    End Function
    Private Shared Function CompareWbs(ByVal x As wbs, ByVal y As wbs) As Integer
        Dim result As Integer = x.parentId.CompareTo(y.parentId)
        If result = 0 Then
            result = x.seqNumber.CompareTo(y.seqNumber)
            If result = 0 Then
                Return x.id.CompareTo(y.id)
            Else
                Return result
            End If
        Else
            Return result
        End If
        'Return (x.parentId & x.seqNumber & x.id).CompareTo(y.parentId & y.seqNumber & y.id)
    End Function

    Private Shared Function CompareWbsByNumber(ByVal x As wbs, ByVal y As wbs) As Integer
        Return (x.number).CompareTo(y.number)
    End Function

    Public Function UpdateTaskStmt(ByVal t As wbs, ByVal task_is_milestone As Integer, attributesList As Dictionary(Of P6CustomAttributes.AttributeCode, String), ByVal index As Integer, ByVal actDuration As Integer) As String

        Dim Sql As String = String.Empty
        Dim connType As IDataClient.eConnectionType = _Caller._DMLDBConn.ConnectionType

        Try

            Sql = " update MSP_TASKS set "
            Sql &= " TASK_IS_EXTERNAL = 0"
            Sql &= ", TASK_IS_MILESTONE = " & task_is_milestone.ToString()
            Sql &= ", RESERVED_DATA = 0"
            Sql &= ", TASK_ID = " & t.number.ToString()
            Sql &= ", TASK_NAME = " & DBUtilities.QS(t.name)
            Sql &= ", TASK_DUR = " & RWUtilities.NrToString(t.duration)
            Sql &= ", TASK_OUTLINE_LEVEL = " & DBUtilities.QS(t.level)
            Sql &= ", TASK_OUTLINE_NUM = " & DBUtilities.QS(t.chainName)
            Sql &= ", TASK_START_DATE = " & IBaseDataMgr.DatetimeToDBFormula(t.forecastStartDate, "8:00", connType)
            Sql &= ", TASK_FINISH_DATE = " & IBaseDataMgr.DatetimeToDBFormula(t.forecastEndDate, "17:00", connType)
            Sql &= ", TASK_BASE_START = " & IBaseDataMgr.DatetimeToDBFormula(t.baselineStartDate, "8:00", connType)
            Sql &= ", TASK_BASE_FINISH = " & IBaseDataMgr.DatetimeToDBFormula(t.baselineEndDate, "17:00", connType)
            Sql &= ", TASK_WBS = " & DBUtilities.QS(t.chainName)
            Sql &= ", TASK_WBS_RIGHTMOST_LEVEL = " & DBUtilities.QS(t.chainName.Split(".")(t.chainName.Split(".").Length - 1))
            Sql &= ", TASK_CONSTRAINT_DATE = " & IBaseDataMgr.DatetimeToDBFormula(t.forecastStartDate, "8:00", connType)
            Sql &= ", TASK_CONSTRAINT_TYPE = 4"
            Sql &= ", EXT_EDIT_REF_DATA = '1'"
            Sql &= ", TASK_ACT_START = " & IBaseDataMgr.DatetimeToDBFormula(t.actualStartDate, "8:00", connType)
            Sql &= ", TASK_ACT_FINISH = " & IBaseDataMgr.DatetimeToDBFormula(t.actualEndDate, "17:00", connType)
            ''updated below Sql &= ", TASK_PCT_COMP = " & RWUtilities.NrToString(t.percentComplete, 2)
            Sql &= ", TASK_PHY_PCT_COMP = " & RWUtilities.NrToString(t.physicalPercentComplete)
            Sql &= ", TASK_ACT_DUR = " & RWUtilities.NrToString(actDuration)
            Sql &= ", TASK_REM_DUR = " & RWUtilities.NrToString(t.remainDuration)
            Sql &= ", TASK_BASE_DUR = " & RWUtilities.NrToString(t.totMinForecast)
            Sql &= ", TASK_IS_SUMMARY = " & IIf(t.hasChild(t, _wbs), 1, 0)
            Sql &= ", TASK_CODE = '" & t.code & "' "
            Sql &= ", ITEM_PATH = " & DBUtilities.QS(GetWbsItemPath(t))

            If Not attributesList Is Nothing Then
                Dim value As String = String.Empty
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.InProgressFlag, value) Then
                    Sql &= ", FLG_ONGOING = '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    Sql &= ", FLG_ONGOING = NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.ToBeWatchedFlag, value) Then
                    Sql &= ", FLG_TOBEWATCHED = '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    Sql &= ", FLG_TOBEWATCHED = NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.SlipChartFlag, value) Then
                    Sql &= ", FLG_SLIP_CHART = '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    Sql &= ", FLG_SLIP_CHART = NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.SkylineFlag, value) Then
                    Sql &= ", FLG_SKYLINE = '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    Sql &= ", FLG_SKYLINE = NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.PPTFlag, value) Then
                    Sql &= ", FLG_POWERPT = '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    Sql &= ", FLG_POWERPT = NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.DeliveryMgrCode, value) Then
                    Sql &= ", DELIVERY_MGR_CODE = " & DBUtilities.QS(value) & " "
                Else
                    Sql &= ", DELIVERY_MGR_CODE = NULL "
                End If
            Else
                Sql &= ", FLG_ONGOING = NULL "
                Sql &= ", FLG_TOBEWATCHED = NULL "
                Sql &= ", FLG_SLIP_CHART = NULL "
                Sql &= ", FLG_SKYLINE = NULL "
                Sql &= ", FLG_POWERPT = NULL "
                Sql &= ", DELIVERY_MGR_CODE = NULL "
            End If


            If _LaborQtyFlag Then
                Sql &= ", TASK_ACT_WORK =  " & RWUtilities.NrToString(t.ActualWork) & " "
                Sql &= ", TASK_REM_WORK = " & RWUtilities.NrToString(t.RemainingWork) & " "
                Sql &= ", TASK_WORK = " & RWUtilities.NrToString(t.TargetWork) & " "
                If t.WorkPercComplete.HasValue Then : Sql &= ", TASK_PCT_COMP = " & RWUtilities.NrToString(t.WorkPercComplete.Value, 2) & " "
                Else : Sql &= ", TASK_PCT_COMP = NULL " : End If
                'Sql &= ", TASK_PCT_COMP = " & IIf(Not t.WorkPercComplete.HasValue, "NULL", RWUtilities.NrToString(Convert.ToDecimal(t.WorkPercComplete), 2)) & " "
            Else
                Sql &= ", TASK_ACT_WORK =  " & RWUtilities.NrToString(t.ActualEqpmQty) & " "
                Sql &= ", TASK_REM_WORK = " & RWUtilities.NrToString(t.RemainingEqpmQty) & " "
                Sql &= ", TASK_WORK = " & RWUtilities.NrToString(t.TargetEqpmQty) & " "
                If t.EqpmQtyPercComplete.HasValue Then : Sql &= ", TASK_PCT_COMP = " & RWUtilities.NrToString(t.EqpmQtyPercComplete.Value, 2) & " "
                Else : Sql &= ", TASK_PCT_COMP = NULL " : End If
                'Sql &= ", TASK_PCT_COMP = " & IIf(Not t.EqpmQtyPercComplete.HasValue, "NULL", RWUtilities.NrToString(Convert.ToDecimal(t.EqpmQtyPercComplete), 2)) & " "
            End If


            Sql &= " where  task_uid =" & t.id & " and proj_id= " & _projIdSigep.ToString() & " "
        Catch ex As Exception
            'WriteLog("updateTask err: " & ex.Message)
        End Try
        Return Sql
    End Function
    Public Function UpdatePTTaskStmt(ByVal t As wbs, attributesList As Dictionary(Of P6CustomAttributes.AttributeCode, String)) As String

        Dim Sql As String = String.Empty
        Dim connType As IDataClient.eConnectionType = _Caller._DMLDBConn.ConnectionType
        Dim poconnType As PO.Base.DBType = IBaseDataMgr.GetDBType(PO.DBInterface.DataSource.PO)

        Try
            Dim forecastStartDate As String = IBaseDataMgr.DatetimeToDBFormula(t.forecastStartDate, "8:00", connType)
            Dim forecastEndDate As String = IBaseDataMgr.DatetimeToDBFormula(t.forecastEndDate, "17:00", connType)
            Dim plannedStartDate As String = IBaseDataMgr.DatetimeToDBFormula(t.baselineStartDate, "8:00", connType)
            Dim plannedEndDate As String = IBaseDataMgr.DatetimeToDBFormula(t.baselineEndDate, "17:00", connType)
            Dim actualStartDate As String = IBaseDataMgr.DatetimeToDBFormula(t.actualStartDate, "8:00", connType)
            Dim actualEndDate As String = IBaseDataMgr.DatetimeToDBFormula(t.actualEndDate, "17:00", connType)

            Select Case poconnType

                Case DBType.Oracle
                    If (String.Compare(forecastStartDate, "NULL", True) <> 0) Then forecastStartDate = "trunc(" & forecastStartDate & ", 'DDD')"
                    If (String.Compare(forecastEndDate, "NULL", True) <> 0) Then forecastEndDate = "trunc(" & forecastEndDate & ", 'DDD')"
                    If (String.Compare(plannedStartDate, "NULL", True) <> 0) Then plannedStartDate = "trunc(" & plannedStartDate & ", 'DDD')"
                    If (String.Compare(plannedEndDate, "NULL", True) <> 0) Then plannedEndDate = "trunc(" & plannedEndDate & ", 'DDD')"
                    If (String.Compare(actualStartDate, "NULL", True) <> 0) Then actualStartDate = "trunc(" & actualStartDate & ", 'DDD')"
                    If (String.Compare(actualEndDate, "NULL", True) <> 0) Then actualEndDate = "trunc(" & actualEndDate & ", 'DDD')"
                Case DBType.SQLServer
                    If (String.Compare(forecastStartDate, "NULL", True) <> 0) Then forecastStartDate = "convert(varchar, convert(datetime, " & forecastStartDate & "), 112)" '"trunc(" & forecastStartDate & ", 'DDD')"
                    If (String.Compare(forecastEndDate, "NULL", True) <> 0) Then forecastEndDate = "convert(varchar, convert(datetime, " & forecastEndDate & "), 112)" '"trunc(" & forecastEndDate & ", 'DDD')"
                    If (String.Compare(plannedStartDate, "NULL", True) <> 0) Then plannedStartDate = "convert(varchar, convert(datetime, " & plannedStartDate & "), 112)" '"trunc(" & plannedStartDate & ", 'DDD')"
                    If (String.Compare(plannedEndDate, "NULL", True) <> 0) Then plannedEndDate = "convert(varchar, convert(datetime, " & plannedEndDate & "), 112)" ' "trunc(" & plannedEndDate & ", 'DDD')"
                    If (String.Compare(actualStartDate, "NULL", True) <> 0) Then actualStartDate = "convert(varchar, convert(datetime, " & actualStartDate & "), 112)" '"trunc(" & actualStartDate & ", 'DDD')"
                    If (String.Compare(actualEndDate, "NULL", True) <> 0) Then actualEndDate = "convert(varchar, convert(datetime, " & actualEndDate & "), 112)" '"trunc(" & actualEndDate & ", 'DDD')"

                Case Else : Throw New Exception("Connection of type " & poconnType.ToString() & " not managed")
            End Select

            Sql = " update PROG_TRACK_T091 set "
            Sql &= " D_INI_FORE = " & forecastStartDate
            Sql &= ", D_FIN_FORE = " & forecastEndDate
            Sql &= ", D_INI_PLAN = " & plannedStartDate
            Sql &= ", D_FIN_PLAN = " & plannedEndDate
            Sql &= ", D_INI_ACT = " & actualStartDate
            Sql &= ", D_FIN_ACT = " & actualEndDate

            ' work
            Sql &= " , S_DIRECT_RES_UM = 'GG' "
            If _LaborQtyFlag Then
                Sql &= " , I_DIRECT_RES_INI = " & RWUtilities.NrToString(t.HrsTargetWork) & " "
                Sql &= " , I_DIRECT_RES_REV = " & RWUtilities.NrToString(t.HrsTargetWork) & " "
                Sql &= " , I_ACTUAL_RES_CUM = " & RWUtilities.NrToString(t.HrsActualWork) & " "
                Sql &= " , I_ESTIMATE_TO = " & RWUtilities.NrToString(t.HrsRemainingWork) & " "
                Sql &= " , I_ESTIMATE_AT = " & RWUtilities.NrToString(t.HrsRemainingWork + t.HrsActualWork, 2) & " "
            Else
                Sql &= " , I_DIRECT_RES_INI = " & RWUtilities.NrToString(t.HrsTargetEqpmQty) & " "
                Sql &= " , I_DIRECT_RES_REV = " & RWUtilities.NrToString(t.HrsTargetEqpmQty) & " "
                Sql &= " , I_ACTUAL_RES_CUM = " & RWUtilities.NrToString(t.HrsActualEqpmQty) & " "
                Sql &= " , I_ESTIMATE_TO = " & RWUtilities.NrToString(t.HrsRemainingEqpmQty) & " "
                Sql &= " , I_ESTIMATE_AT = " & RWUtilities.NrToString(t.HrsRemainingEqpmQty + t.HrsActualEqpmQty, 2) & " "
            End If

            If Not attributesList Is Nothing Then
                Dim deliveryMgrCode As String = String.Empty
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.DeliveryMgrCode, deliveryMgrCode) Then
                    Sql &= ", DELIVERY_MGR_CODE = " & DBUtilities.QS(deliveryMgrCode) & " "
                Else
                    Sql &= ", DELIVERY_MGR_CODE = NULL "
                End If
            Else
                Sql &= ", DELIVERY_MGR_CODE = NULL "
            End If

            Sql &= " where i_act_id = " & t.id.ToString() & " and C_PROG = " & _projIdSigep.ToString() & " "
        Catch ex As Exception
            'WriteLog("updateTask err: " & ex.Message)
            Return Nothing
        End Try
        Return Sql
    End Function
    Public Function InsertTaskStmt(ByVal t As wbs, ByVal task_is_milestone As Integer, attributesList As Dictionary(Of P6CustomAttributes.AttributeCode, String), ByVal index As Integer, ByVal actDuration As Integer) As String

        Dim SqlInsert As String = String.Empty
        Dim connType As IDataClient.eConnectionType = _Caller._DMLDBConn.ConnectionType

        Try

            SqlInsert = "Insert into MSP_TASKS (del_task_uid,TASK_IS_EXTERNAL,TASK_IS_MILESTONE,RESERVED_DATA,"
            SqlInsert &= " PROJ_ID,TASK_UID,TASK_ID,TASK_NAME,TASK_DUR,TASK_OUTLINE_LEVEL,"
            SqlInsert &= " TASK_OUTLINE_NUM,TASK_START_DATE,TASK_FINISH_DATE,TASK_BASE_START,TASK_BASE_FINISH,TASK_WBS,TASK_WBS_RIGHTMOST_LEVEL,TASK_CONSTRAINT_DATE,"
            '' TASK_PCT_COMP: inserted as last column SqlInsert &= " TASK_CONSTRAINT_TYPE,EXT_EDIT_REF_DATA,TASK_ACT_START,TASK_ACT_FINISH,TASK_PCT_COMP,TASK_PHY_PCT_COMP,TASK_ACT_DUR,TASK_REM_DUR,TASK_BASE_DUR,"
            SqlInsert &= " TASK_CONSTRAINT_TYPE,EXT_EDIT_REF_DATA,TASK_ACT_START,TASK_ACT_FINISH,TASK_PHY_PCT_COMP,TASK_ACT_DUR,TASK_REM_DUR,TASK_BASE_DUR,"
            SqlInsert &= " TASK_IS_SUMMARY,TASK_CODE "
            SqlInsert &= ", FLG_ONGOING, FLG_TOBEWATCHED, FLG_SLIP_CHART, FLG_SKYLINE, FLG_POWERPT, DELIVERY_MGR_CODE "
            SqlInsert &= ", TASK_ACT_WORK, TASK_REM_WORK, TASK_WORK, TASK_PCT_COMP "
            SqlInsert &= ", TASK_P6_REMAINING_DURATION, TASK_P6_TOTAL_DURATION, ITEM_PATH "
            SqlInsert &= " ) "

            SqlInsert &= " VALUES(" & t.id.ToString() & ",0," & task_is_milestone & ",0,"
            SqlInsert &= _projIdSigep.ToString() & "," 'PROJ_ID
            SqlInsert &= t.id.ToString() & ","   'RWUtilities.NrToString(index + 1) & ","            'TASK_UID
            SqlInsert &= t.number.ToString() & ","   'TASK_ID
            SqlInsert &= DBUtilities.QS(t.name) & ","          'TASK_NAME
            SqlInsert &= RWUtilities.NrToString(t.duration) & ","                  'TASK_DUR
            SqlInsert &= DBUtilities.QS(t.level) & "," '       TASK_OUTLINE_LEVEL indica il livello nell'albero
            SqlInsert &= DBUtilities.QS(t.chainName) & "," '    TASK_OUTLINE_NUM il nome tipo 1.1.1
            SqlInsert &= IBaseDataMgr.DatetimeToDBFormula(t.forecastStartDate, "8:00", connType) & "," 'data comprensiva di ora
            SqlInsert &= IBaseDataMgr.DatetimeToDBFormula(t.forecastEndDate, "17:00", connType) & ","  'data comprensiva di ora
            SqlInsert &= IBaseDataMgr.DatetimeToDBFormula(t.baselineStartDate, "8:00", connType) & "," 'data comprensiva di ora
            SqlInsert &= IBaseDataMgr.DatetimeToDBFormula(t.baselineEndDate, "17:00", connType) & "," 'data comprensiva di ora
            SqlInsert &= DBUtilities.QS(t.chainName) & "," ' TASK_WBS   --FK alla WBS da vedere bene mi sa che è solo il nome del nodo della WBS
            Dim WBS_RIGHTMOST_LEVEL As String = String.Empty
            WBS_RIGHTMOST_LEVEL = t.chainName.Split(".")(t.chainName.Split(".").Length - 1)
            SqlInsert &= DBUtilities.QS(WBS_RIGHTMOST_LEVEL) & "," 'da vedere
            SqlInsert &= IBaseDataMgr.DatetimeToDBFormula(t.forecastStartDate, "8:00", connType) & ","
            SqlInsert &= "4,'1'," & IBaseDataMgr.DatetimeToDBFormula(t.actualStartDate, "8:00", connType) & "," & IBaseDataMgr.DatetimeToDBFormula(t.actualEndDate, "17:00", connType) & ","
            '' TASK_PCT_COMP: inserted as last column SqlInsert &= RWUtilities.NrToString(t.percentComplete, 2) & "," & RWUtilities.NrToString(t.physicalPercentComplete) & "," & RWUtilities.NrToString(actDuration) & ","
            SqlInsert &= RWUtilities.NrToString(t.physicalPercentComplete) & "," & RWUtilities.NrToString(actDuration) & ","
            SqlInsert &= RWUtilities.NrToString(t.remainDuration) & "," & RWUtilities.NrToString(t.totMinForecast) & "," & IIf(t.hasChild(t, _wbs), 1, 0) & ","
            SqlInsert &= DBUtilities.QS(t.code) & " "

            If Not attributesList Is Nothing Then
                Dim value As String = String.Empty
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.InProgressFlag, value) Then
                    SqlInsert &= ", '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    SqlInsert &= ", NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.ToBeWatchedFlag, value) Then
                    SqlInsert &= ", '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    SqlInsert &= ", NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.SlipChartFlag, value) Then
                    SqlInsert &= ", '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    SqlInsert &= ", NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.SkylineFlag, value) Then
                    SqlInsert &= ", '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    SqlInsert &= ", NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.PPTFlag, value) Then
                    SqlInsert &= ", '" & IIf(value = "1", "Y", "N") & "' "
                Else
                    SqlInsert &= ", NULL "
                End If
                If attributesList.TryGetValue(P6CustomAttributes.AttributeCode.DeliveryMgrCode, value) Then
                    SqlInsert &= ", " & DBUtilities.QS(value) & " "
                Else
                    SqlInsert &= ", NULL "
                End If
            Else
                SqlInsert &= ", NULL "
                SqlInsert &= ", NULL "
                SqlInsert &= ", NULL "
                SqlInsert &= ", NULL "
                SqlInsert &= ", NULL "
                SqlInsert &= ", NULL "
            End If

            If _LaborQtyFlag Then
                SqlInsert &= ", " & t.ActualWork.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                SqlInsert &= ", " & t.RemainingWork.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                SqlInsert &= ", " & t.TargetWork.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                'SqlInsert &= ", " & IIf(Not t.WorkPercComplete.HasValue, "NULL", RWUtilities.NrToString(t.WorkPercComplete, 2)) & " "
                If t.WorkPercComplete.HasValue Then : SqlInsert &= ", " & RWUtilities.NrToString(Convert.ToDecimal(t.WorkPercComplete), 2) & " "
                Else : SqlInsert &= ", NULL " : End If
            Else
                SqlInsert &= ", " & t.ActualEqpmQty.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                SqlInsert &= ", " & t.RemainingEqpmQty.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                SqlInsert &= ", " & t.TargetEqpmQty.ToString(System.Globalization.CultureInfo.InvariantCulture) & " "
                'SqlInsert &= ", " & IIf(Not t.EqpmQtyPercComplete.HasValue, "NULL", RWUtilities.NrToString(t.EqpmQtyPercComplete, 2)) & " "
                If t.EqpmQtyPercComplete.HasValue Then : SqlInsert &= ", " & RWUtilities.NrToString(Convert.ToDecimal(t.EqpmQtyPercComplete), 2) & " "
                Else : SqlInsert &= ", NULL " : End If
            End If

            SqlInsert &= ", " & RWUtilities.NrToString(t.remainDurationP6) & ", " & RWUtilities.NrToString(t.totalDurationP6) & ", " & DBUtilities.QS(GetWbsItemPath(t))
            SqlInsert &= " ) "

            Return SqlInsert
        Catch ex As Exception
            'WriteLog("InsertTask err: " & ex.Message)
            Return Nothing
        End Try


    End Function
    Private Function GetWbsItemPath(t As wbs)
        Dim wbsItemPath As String = String.Empty
        If Not t.ItemPath Is Nothing AndAlso t.ItemPath.Count > 0 Then
            If t.ItemPath.Count = 1 Then : wbsItemPath = t.ItemPath(0)
            Else : wbsItemPath = String.Join(".", t.ItemPath.ToArray().Reverse().ToArray()) : End If
        End If
        Return wbsItemPath
    End Function


End Class
