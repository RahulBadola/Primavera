Imports System.Data.SqlClient
Imports Primavera6
'Imports PwCrypt
'Imports SIGEPPRIMAVERA.Primavera5
'Imports SIGEPPRIMAVERA.PwCrypt


Public Class DbUtility

    Public Shared Function replaceNull(ByVal x As Object, Optional ByVal ReplaceWith As Object = "") As Object
        If IsNothing(x) Then Return ReplaceWith
        If IsDBNull(x) Then Return ReplaceWith
        Return x
    End Function

    Public Shared Function returnNull(ByVal s As String)
        If Trim(s) = "" Then
            Return "NULL"
        Else
            Return s
        End If
    End Function

    Public Shared Function QN(ByVal s As String) As String
        If Not IsNumeric(s) Then Return "NULL"
        Dim n As Decimal = CType(s, Decimal)
        If n = Decimal.MinValue Then Return "NULL"
        Return n.ToString.Replace(",", ".")
    End Function

    Public Shared Function QN(ByVal n As Decimal, Optional ByVal decimals As Integer = 0) As String
        If decimals > 0 Then
            Return n.ToString("0." & StrDup(decimals, "0"), New System.Globalization.CultureInfo(1033))
        End If
        Return n.ToString(New System.Globalization.CultureInfo(1033))
    End Function

    Public Shared Function QD(ByVal s As String) As String
        Dim USE_JOIN As Boolean = (System.Configuration.ConfigurationManager.AppSettings("USE_JOIN") = "YES")
        If Not IsDate(s) Then Return "NULL"
        Dim d As Date = CType(s, Date)
        If d = Date.MinValue Then Return "NULL"
        If USE_JOIN Then
            Return "'" & Format(d, "yyyyMMdd") & "'"
        Else
            Return "TO_DATE('" & Format(d, "dd-MM-yyyy") & "', 'DD-MM-YYYY')"
        End If
    End Function

    Public Shared Function QDT(ByVal s As String, ByVal t As String) As String
        If s Is Nothing Then Return "NULL"
        Dim USE_JOIN As Boolean = (System.Configuration.ConfigurationManager.AppSettings("sigep_USE_ORACLE") = "false")
        If Not IsDate(s) Then Return "NULL"
        Dim d As Date = CType(s & " ", Date)
        If d = Date.MinValue Then Return "NULL"
        If USE_JOIN Then
            Return "'" & Format(d, "yyyyMMdd HH\:mm") & "'"
        Else
            Return "TO_DATE('" & Format(d, "dd-MM-yyyy HH\:mm") & "', 'DD-MM-YYYY HH24:MI')"
        End If
    End Function
    Public Shared Function QDT(ByVal d As DateTime, ByVal t As String) As String
        If IsNothing(d) Then Return "NULL"
        If IsDBNull(d) Then Return "NULL"
        If DateTime.Compare(d, DateTime.MinValue) = 0 Then Return "NULL"
        If DateTime.Compare(d, DateTime.MaxValue) = 0 Then Return "NULL"
        Dim USE_JOIN As Boolean = (System.Configuration.ConfigurationManager.AppSettings("sigep_USE_ORACLE") = "false")

        If USE_JOIN Then
            Return "'" & d.ToString("yyyyMMdd HH:mm") & "'"
        Else
            Return "TO_DATE('" & d.ToString("yyyyMMdd HH:mm") & "', 'yyyyMMdd HH24:MI')"
        End If
    End Function

    Public Shared Function QS(ByVal s As String, Optional ByVal IncludeEqual As Boolean = False) As String
        If s = "" Then
            Return IIf(IncludeEqual, " IS NULL", "NULL")
        Else
            Return IIf(IncludeEqual, "=", "") & "'" & Replace(s, "'", "''") & "'"
        End If
    End Function

    Public Shared Function QueryTable(ByVal sql As String) As DataTable
        Dim cn As New OleDb.OleDbConnection(DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("sigep_DSN")))
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If
        Dim da As New OleDb.OleDbDataAdapter(sql, cn)
        Dim dt As DataTable = New DataTable
        da.Fill(dt)
        If cn.State = ConnectionState.Open Then
            cn.Close()
        End If
        dt.Dispose()
        Return dt
    End Function

    'Public Shared Function QueryValue(ByVal sql As String) As Object
    '    Dim cn As New OleDb.OleDbConnection(DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("sigep_DSN")))
    '    cn.Open()
    '    Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(sql, cn)
    '    Dim value As Object = cmd.ExecuteScalar
    '    cn.Close()
    '    Return value
    'End Function

    Public Shared Function QueryRow(ByVal sql As String) As DataRow
        Dim dt As DataTable = QueryTable(sql)
        If dt.Rows.Count = 0 Then Return Nothing
        Return dt.Rows(0)
    End Function

    Public Shared Function QueryExec(ByVal sql As String, ByVal cmd As OleDb.OleDbCommand) As Integer
        cmd.CommandText = sql
        Dim retVal As Object = cmd.ExecuteNonQuery
        Return retVal
    End Function

    Public Shared Function QueryTable(ByVal sql As String, ByVal source As datasource) As DataTable
        Dim dt As DataTable = New DataTable
        If source = datasource.sigep Then
            Dim cn = Nothing
            Dim cmd = Nothing
            Dim da = Nothing
            cn = New OleDb.OleDbConnection(DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("sigep_DSN")))
            cmd = New OleDb.OleDbCommand(sql, cn)
            da = New OleDb.OleDbDataAdapter(sql, CType(cn, OleDb.OleDbConnection))
            cn.Open()
            da.Fill(dt)
            cn.Close()
            cmd.Dispose()
            ' dt.Dispose()
        ElseIf source = datasource.pm5 Then
            Dim settings As New My.MySettings
            Dim cn = Nothing
            Dim cmd = Nothing
            Dim da = Nothing
            If System.Configuration.ConfigurationManager.AppSettings("pm5_USE_ORACLE").ToLower = "false" Then
                cn = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("pm5_DSN"))
                cmd = New SqlCommand(sql, cn)
                da = New SqlDataAdapter(sql, CType(cn, SqlConnection))
            Else
                cn = New OleDb.OleDbConnection(System.Configuration.ConfigurationManager.AppSettings("pm5_DSN"))
                cmd = New OleDb.OleDbCommand(sql, cn)
                da = New OleDb.OleDbDataAdapter(sql, CType(cn, OleDb.OleDbConnection))
            End If
            cn.Open()
            da.Fill(dt)
            cn.Close()
            cmd.Dispose()
            ' dt.Dispose()
        End If
        Return dt
    End Function

    Public Shared Function QueryExec(ByVal sql As String, ByVal source As datasource, Optional ByVal cmdTrans As OleDb.OleDbCommand = Nothing) As Integer
        Dim value As Integer
        Dim cn = Nothing
        Dim da = Nothing
        Dim cmd = Nothing

        '///////////////////////////////////////////////////////////// la transazione la uso solo per le scritture in sigep /////////////////////////////////////////////////////////////
        If source = datasource.sigep Then
            If cmdTrans Is Nothing Then
                cn = New OleDb.OleDbConnection(DeCryptDsnPassword(System.Configuration.ConfigurationManager.AppSettings("sigep_DSN")))
                da = New OleDb.OleDbDataAdapter(sql, CType(cn, OleDb.OleDbConnection))
                cn.Open()
                cmd = New OleDb.OleDbCommand(sql, cn)
                value = cmd.ExecuteNonQuery
                cn.Close()
                cmd.Dispose()
            Else
                cmd = cmdTrans
                CType(cmd, OleDb.OleDbCommand).CommandText = sql
                value = cmd.ExecuteNonQuery
            End If

            '///////////////////////////////////////////////////////////// primavera //////////////////////////////////////////////////////////////////////////////////////////////////
        ElseIf source = datasource.pm5 Then
            Dim settings As New My.MySettings
            If System.Configuration.ConfigurationManager.AppSettings("pm5_USE_ORACLE").ToLower = "true" Then
                cn = New OleDb.OleDbConnection(System.Configuration.ConfigurationManager.AppSettings("pm5_DSN"))
                If cmdTrans Is Nothing Then
                    cmd = New OleDb.OleDbCommand(sql, cn)
                Else
                    cmd = New OleDb.OleDbCommand(sql, cmdTrans.Connection)
                End If
            Else
                cn = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("pm5_DSN"))
                cmd = New SqlCommand(sql, cn)
            End If
            cn.Open()
            value = cmd.ExecuteNonQuery
            cn.Close()
            cmd.Dispose()
        End If
        Return value
    End Function

    Public Shared Function QueryValue(ByVal sql As String, ByVal source As datasource) As Object

        Dim dt As DataTable = QueryTable(sql, source)
        If dt.Rows.Count = 0 Then Return Nothing
        If Not IsDBNull(dt.Rows(0)(0)) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)(0)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function QueryValue(ByVal sql As String, ByVal source As datasource, ByVal cmdTrans As OleDb.OleDbCommand) As Object
        Dim cmd = Nothing
        '///////////////////////////////////////////////////////////// la transazione la uso solo per le scritture in sigep /////////////////////////////////////////////////////////////
        If source = datasource.sigep Then
            cmd = cmdTrans
            CType(cmd, OleDb.OleDbCommand).CommandText = sql
        End If
        Return CType(cmd, OleDb.OleDbCommand).ExecuteScalar
    End Function

    'Public Shared Function QueryTable(ByVal sql As String, ByVal source As datasource, ByVal cmdTrans As OleDb.OleDbCommand) As DataTable
    '    Dim da = New OleDb.OleDbDataAdapter(sql, cmdTrans.Connection)
    '    cmdTrans.ExecuteNonQuery()
    '    Dim dt As New DataTable
    '    da.Fill(dt)
    '    Return dt
    'End Function


End Class



