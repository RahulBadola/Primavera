Imports PO.DataClient

Public Class milestone
    Public C_PROG, V_LAG, V_DELAY, C_OWNER As Integer
    Public C_MLSTN, S_MLSTN, C_TIP_MLSTN, C_MSP_TASK, F_LINK, F_MSP_TASK_STATUS, F_CRITICALY, DESCRIPTION, BENEFIT As String
    Public D_PLANNED, D_ACTUAL, D_FORECAST As Date
    Public FLG_SKYLINE, FLG_SLIP_CHART, FLG_POWERPT, FLG_ONGOING, FLG_TOBEWATCHED As String
    Private Shared cmdSigep As New OleDb.OleDbCommand

    'Private Sub WriteLog(ByVal text As String)
    '    Dim stmt As String = " insert into temp_log_table values (sysdate, '" & text.Replace("'", "''") & "')"
    '    DbUtility.QueryExec(stmt, Primavera6.datasource.sigep, cmdSigep)
    'End Sub

    Public Sub New(ByVal cmd As OleDb.OleDbCommand)
        cmdSigep = cmd
    End Sub

    Public Function create() As Boolean
        Try
            Dim stri As String = "insert into MLSTN_CONTR_T119 (C_PROG, C_MLSTN, S_MLSTN, C_TIP_MLSTN, D_PLANNED, D_ACTUAL, D_FORECAST, C_MSP_TASK, F_LINK, V_LAG, F_MSP_TASK_STATUS, F_CRITICALY, V_DELAY, C_OWNER, DESCRIPTION, BENEFIT, FLG_SKYLINE, FLG_SLIP_CHART, FLG_POWERPT, FLG_ONGOING, FLG_TOBEWATCHED) values " &
                                                                                                   "  ({0}, '{1}', '{2}','{3}',{4},{5},{6},'{7}','{8}',{9},'{10}','{11}',{12},{13},'{14}','{15}', '{16}', '{17}', '{18}', '{19}', '{20}') "
            stri = String.Format(stri, Me.C_PROG, Me.C_MLSTN.Replace("'", "''"), Me.S_MLSTN.Replace("'", "''"), Me.C_TIP_MLSTN, DbUtility.QD(Me.D_PLANNED), DbUtility.QD(Me.D_ACTUAL), DbUtility.QD(Me.D_FORECAST), Me.C_MSP_TASK, Me.F_LINK, Me.V_LAG, Me.F_MSP_TASK_STATUS, Me.F_CRITICALY, Me.V_DELAY, Me.C_OWNER, Me.DESCRIPTION.Replace("'", "''"), Me.BENEFIT.Replace("'", "''"))
            'WriteLog(stri)
            DbUtility.QueryExec(String.Format(stri, Me.C_PROG, Me.C_MLSTN, Me.S_MLSTN, Me.C_TIP_MLSTN, DbUtility.QD(Me.D_PLANNED), DbUtility.QD(Me.D_ACTUAL), DbUtility.QD(Me.D_FORECAST), Me.C_MSP_TASK, Me.F_LINK, Me.V_LAG, Me.F_MSP_TASK_STATUS, Me.F_CRITICALY, Me.V_DELAY, Me.C_OWNER, Me.DESCRIPTION, Me.BENEFIT, Me.FLG_SKYLINE, Me.FLG_SLIP_CHART, Me.FLG_POWERPT, Me.FLG_ONGOING, Me.FLG_TOBEWATCHED), Primavera6.datasource.sigep, cmdSigep)
        Catch ex As Exception
            'WriteLog("milestone.create err: " & ex.Message)
            Return False
        End Try
        Return True
    End Function


    Public Function delete()
        Try
            Dim stri As String = "delete from MLSTN_CONTR_T119 where c_prog= " & Me.C_PROG & " and c_mlstn = '" & Me.C_MLSTN & "'"
            DbUtility.QueryExec(stri, Primavera6.datasource.sigep, cmdSigep)
        Catch ex As Exception
            'WriteLog("milestone.delete err: " & ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function deleteAll(ByVal prog_id As Integer)
        Try
            Dim stri As String = "delete from MLSTN_CONTR_T119 where c_tip_mlstn = 'PM6' and c_prog= " & prog_id ' & " and c_mlstn = '" & Me.C_MLSTN & "'"
            DbUtility.QueryExec(stri, Primavera6.datasource.sigep, cmdSigep)
        Catch ex As Exception
            'WriteLog("milestone.deleteAll err: " & ex.Message)
            Return False
        End Try
        Return True
    End Function


End Class

'Public Class MlstnAttributes

'    Private _SkylineFlag As Boolean
'    Private _SlipChartFlag As Boolean
'    Private _InProgressFlag As Boolean
'    Private _ToBeWatchedFlag As Boolean
'    Private _PPTFlag As Boolean

'    Public Sub New()

'        _SkylineFlag = False
'        _SlipChartFlag = False
'        _InProgressFlag = False
'        _ToBeWatchedFlag = False
'        _PPTFlag = False

'    End Sub

'    Public Sub SetSkylineFlag(ByVal pm6Value As String)
'        _SkylineFlag = (String.Compare("X", pm6Value, True) = 0)
'    End Sub
'    Public Sub SetSlipChartFlag(ByVal pm6Value As String)
'        _SlipChartFlag = (String.Compare("X", pm6Value, True) = 0)
'    End Sub
'    Public Sub SetPPTFlag(ByVal pm6Value As String)
'        _PPTFlag = (String.Compare("X", pm6Value, True) = 0)
'    End Sub
'    Public Sub SetSkylineAttrFlags(ByVal pm6Value As String)
'        _InProgressFlag = (String.Compare("PR", pm6Value, True) = 0)
'        _ToBeWatchedFlag = (String.Compare("TBW", pm6Value, True) = 0)
'    End Sub

'    Public ReadOnly Property SkylineFlag() As Boolean
'        Get
'            Return _SkylineFlag
'        End Get
'    End Property
'    Public ReadOnly Property SlipChartFlag() As Boolean
'        Get
'            Return _SlipChartFlag
'        End Get
'    End Property
'    Public ReadOnly Property InProgressFlag() As Boolean
'        Get
'            Return _InProgressFlag
'        End Get
'    End Property
'    Public ReadOnly Property ToBeWatchedFlag() As Boolean
'        Get
'            Return _ToBeWatchedFlag
'        End Get
'    End Property
'    Public ReadOnly Property PPTFlag() As Boolean
'        Get
'            Return _PPTFlag
'        End Get
'    End Property

'End Class
