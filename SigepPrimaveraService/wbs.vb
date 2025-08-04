Imports System.Linq

Public Class wbs
    Public name As String
    Public level As Integer
    Public number As Integer
    Public isRoot As Boolean
    Public id As Integer
    Public parentId As Integer
    Public chainName As String
    Public nodeType As Boolean 'aggiunto per tenere traccia dei nodi veramente di tipo wbs o di tipo task: false indica i nodi riguardanti le attività e non la wbs
    Public visible As Boolean ' i nodi della wbs che nn hanno task associati nn li facio fadere sul gant esattamente come fa primavera
    Public forecastStartDate As DateTime
    Public forecastEndDate As DateTime
    Public actualStartDate As DateTime
    Public actualEndDate As DateTime
    Public baselineStartDate As DateTime
    Public baselineEndDate As DateTime
    Public activityType As String
    Public status As String
    Public code As String
    Public physicalPercentComplete As Integer
    Public percentComplete As Decimal
    Public actualDuration As Integer
    Public remainDuration As Integer
    Public totalDurationP6 As Decimal
    Public remainDurationP6 As Decimal
    Public totMinForecast As Integer
    Public duration As Integer
    Public seqNumber As Integer
    Public TargetWork As Decimal = 0
    Public ActualWork As Decimal = 0
    Public RemainingWork As Decimal = 0
    Public HrsTargetWork As Decimal = 0
    Public HrsActualWork As Decimal = 0
    Public HrsRemainingWork As Decimal = 0
    Public TargetEqpmQty As Decimal = 0
    Public ActualEqpmQty As Decimal = 0
    Public RemainingEqpmQty As Decimal = 0
    Public HrsTargetEqpmQty As Decimal = 0
    Public HrsActualEqpmQty As Decimal = 0
    Public HrsRemainingEqpmQty As Decimal = 0
    Public WorkPercComplete As Decimal? = Nothing
    Public EqpmQtyPercComplete As Decimal? = Nothing
    Public Calendar As P6Calendar = Nothing
    Public ItemPath As New List(Of String)

    Public Sub New()
        name = String.Empty
        level = 0
        number = 0
        isRoot = False
        nodeType = True
        visible = True
    End Sub

    Public Function hasChild(ByVal w As wbs, ByVal _wbs As List(Of wbs)) As Boolean

        Dim obj As wbs = _wbs.FirstOrDefault(Function(x) x.parentId = w.id)
        Return (Not obj Is Nothing)
        'Dim ret As Boolean = False
        'For Each ww As wbs In _wbs
        '    If ww.parentId = w.id Then
        '        ret = True
        '    End If
        'Next
        'Return ret

    End Function


    'Public Function getIndex(ByVal id As Integer, ByVal _wbs As List(Of wbs)) As Integer

    '    For i As Integer = 0 To _wbs.Count - 1
    '        If _wbs(i).id = id Then
    '            Return i
    '        End If
    '    Next

    'End Function
End Class
