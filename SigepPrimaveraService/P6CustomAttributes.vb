' This class must be updated whenever custom attributes references change
Public NotInheritable Class P6CustomAttributes

    Public Enum AttributeCode
        NotDefined = 0
        SkylineFlag = 1
        SlipChartFlag = 2
        InProgressFlag = 3
        ToBeWatchedFlag = 4
        PPTFlag = 5
        DeliveryMgrCode = 6
    End Enum

    Private Const SKYLINE_FLAG_COL_KEY As String = "skyline_flag_col"
    Private Const SLIPCHART_FLAG_COL_KEY As String = "slipchart_flag_col"
    Private Const TOBEWATCH_FLAG_COL_KEY As String = "tobewatched_flag_col"
    Private Const INPROGRESS_COL_KEY As String = "inprogress_col"
    Private Const PPT_FLAG_COL_KEY As String = "ppt_flag_col"
    Private Const DELIVMGRCODE_COL_KEY As String = "delivery_mgr_code_col"

    Private Shared Function EncodeAttribute(colName As String) As AttributeCode
        Select Case colName
            Case SKYLINE_FLAG_COL_KEY : Return AttributeCode.SkylineFlag
            Case TOBEWATCH_FLAG_COL_KEY : Return AttributeCode.ToBeWatchedFlag
            Case SLIPCHART_FLAG_COL_KEY : Return AttributeCode.SlipChartFlag
            Case INPROGRESS_COL_KEY : Return AttributeCode.InProgressFlag
            Case PPT_FLAG_COL_KEY : Return AttributeCode.PPTFlag
            Case DELIVMGRCODE_COL_KEY : Return AttributeCode.DeliveryMgrCode
            Case Else : Return AttributeCode.NotDefined
        End Select
    End Function

    ''' <summary>
    ''' loads custom attributes list from the configuration file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function LoadCustomAttributesList() As Dictionary(Of String, AttributeCode)

        Dim attrList As New Dictionary(Of String, AttributeCode)
        ' delivery_mgr_code_col
        Dim attrColumn As String = System.Configuration.ConfigurationManager.AppSettings(DELIVMGRCODE_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(DELIVMGRCODE_COL_KEY))
        ' keys to retrieve milestones additional infos
        ' skyline flag column
        attrColumn = System.Configuration.ConfigurationManager.AppSettings(SKYLINE_FLAG_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(SKYLINE_FLAG_COL_KEY))
        ' slip chart flag column
        attrColumn = System.Configuration.ConfigurationManager.AppSettings(SLIPCHART_FLAG_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(SLIPCHART_FLAG_COL_KEY))
        ' ppt flag column
        attrColumn = System.Configuration.ConfigurationManager.AppSettings(PPT_FLAG_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(PPT_FLAG_COL_KEY))
        ' in progress flag column
        attrColumn = System.Configuration.ConfigurationManager.AppSettings(INPROGRESS_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(INPROGRESS_COL_KEY))
        ' to be watched flag column
        attrColumn = System.Configuration.ConfigurationManager.AppSettings(TOBEWATCH_FLAG_COL_KEY)
        If Not String.IsNullOrEmpty(attrColumn) Then attrList.Add(attrColumn.ToLower(), EncodeAttribute(TOBEWATCH_FLAG_COL_KEY))

        Return attrList

    End Function
    Public Shared Function ParseAttrValue(attrCode As AttributeCode, attrValue As String)
        Select Case attrCode
            Case AttributeCode.SkylineFlag, AttributeCode.SlipChartFlag, AttributeCode.PPTFlag
                If String.Compare("X", attrValue, True) = 0 Then : Return "1"
                Else : Return "0" : End If
            Case AttributeCode.InProgressFlag
                If String.Compare("PR", attrValue, True) = 0 Then : Return "1"
                Else : Return "0" : End If
            Case AttributeCode.ToBeWatchedFlag
                If String.Compare("TBW", attrValue, True) = 0 Then : Return "1"
                Else : Return "0" : End If
            Case Else : Return attrValue
        End Select
    End Function

End Class

