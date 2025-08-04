Public Class task
    Private _id As String
    Private _parentId As DateTime
    Private _name As String
    Private _description As String
    Private _forecastStartDate As DateTime
    Private _forecastEndDate As DateTime
    Private _actualStartDate As DateTime
    Private _actualEndDate As DateTime
    Private _baselineStartDate As DateTime
    Private _baselineEndDate As DateTime
    Private _isSummary As Integer
    Private _percentComplete As Decimal
    Private _physicalPercentComplete As Integer
    Private _actualDuration As Integer
    Private _remainDuration As Integer
    Private _totalDurationP6 As Decimal
    Private _remainDurationP6 As Decimal
    Private _totMinForecast As Integer
    Private _activityType As String
    Private _status As String
    Private _w As New wbs()
    Private _code As String
    Private _activityCode As Integer
    Private _duration As Integer
    Private _TargetWork As Decimal = 0
    Private _ActualWork As Decimal = 0
    Private _RemainingWork As Decimal = 0
    Private _HrsTargetWork As Decimal = 0
    Private _HrsActualWork As Decimal = 0
    Private _HrsRemainingWork As Decimal = 0
    Private _TargetEqpmQty As Decimal = 0
    Private _ActualEqpmQty As Decimal = 0
    Private _RemainingEqpmQty As Decimal = 0
    Private _HrsTargetEqpmQty As Decimal = 0
    Private _HrsActualEqpmQty As Decimal = 0
    Private _HrsRemainingEqpmQty As Decimal = 0
    Private _WorkPercComplete As Decimal? = Nothing
    Private _EqpmQtyPercComplete As Decimal? = Nothing
    Private _Calendar As P6Calendar = Nothing
    Private _ItemPath As New List(Of String)

#Region "property"

    Public Property remainDuration() As Integer
        Get
            Return _remainDuration
        End Get
        Set(ByVal value As Integer)
            _remainDuration = value
        End Set
    End Property

    Public Property remainDurationP6() As Integer
        Get
            Return _remainDurationP6
        End Get
        Set(ByVal value As Integer)
            _remainDurationP6 = value
        End Set
    End Property

    Public Property duration() As Integer
        Get
            Return _duration
        End Get
        Set(ByVal value As Integer)
            _duration = value
        End Set
    End Property

    Public Property activityCode() As Integer
        Get
            Return _activityCode
        End Get
        Set(ByVal value As Integer)
            _activityCode = value
        End Set
    End Property

    Public Property status() As String
        Get
            Return _status
        End Get
        Set(ByVal value As String)
            _status = value
        End Set
    End Property

    Public Property code() As String
        Get
            Return _code
        End Get
        Set(ByVal value As String)
            _code = value
        End Set
    End Property

    Public Property activityType() As String
        Get
            Return _activityType
        End Get
        Set(ByVal value As String)
            _activityType = value
        End Set
    End Property

    Public Property w() As wbs
        Get
            Return (_w)
        End Get
        Set(ByVal value As wbs)
            _w = value
        End Set
    End Property

    Public Property isSummary() As Integer
        Get
            Return _isSummary
        End Get
        Set(ByVal value As Integer)
            _isSummary = value
        End Set
    End Property



    Public Property totMinForecast() As Integer
        Get
            Return _totMinForecast
        End Get
        Set(ByVal value As Integer)
            _totMinForecast = value
        End Set
    End Property

    Public Property actualDuration() As Integer
        Get
            Return _actualDuration
        End Get
        Set(ByVal value As Integer)
            _actualDuration = value
        End Set
    End Property

    Public Property totalDurationP6() As Integer
        Get
            Return _totalDurationP6
        End Get
        Set(ByVal value As Integer)
            _totalDurationP6 = value
        End Set
    End Property

    Public Property percentComplete() As Decimal
        Get
            Return _percentComplete
        End Get
        Set(ByVal value As Decimal)
            _percentComplete = value
        End Set
    End Property

    Public Property physicalPercentComplete() As Integer
        Get
            Return _physicalPercentComplete
        End Get
        Set(ByVal value As Integer)
            _physicalPercentComplete = value
        End Set
    End Property

    Public Property parentId() As String
        Get
            Return _parentId
        End Get
        Set(ByVal value As String)
            _parentId = value
        End Set
    End Property

    Public Property id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            _id = value
        End Set
    End Property

    Public Property name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Property description() As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

    Public Property forecaststartDate() As DateTime
        Get
            Return _forecastStartDate
        End Get
        Set(ByVal value As DateTime)
            _forecastStartDate = value
        End Set
    End Property

    Public Property forecastEndDate() As DateTime
        Get
            Return _forecastEndDate
        End Get
        Set(ByVal value As DateTime)
            _forecastEndDate = value
        End Set
    End Property

    Public Property baselineStartDate() As DateTime
        Get
            Return _baselineStartDate
        End Get
        Set(ByVal value As DateTime)
            _baselineStartDate = value
        End Set
    End Property

    Public Property baselineEndDate() As DateTime
        Get
            Return _baselineEndDate
        End Get
        Set(ByVal value As DateTime)
            _baselineEndDate = value
        End Set
    End Property

    Public Property actualEndDate() As DateTime
        Get
            Return _actualEndDate
        End Get
        Set(ByVal value As DateTime)
            _actualEndDate = value
        End Set
    End Property

    Public Property actualStartDate() As DateTime
        Get
            Return _actualStartDate
        End Get
        Set(ByVal value As DateTime)
            _actualStartDate = value
        End Set
    End Property

    Public Property TargetWork As Decimal
        Get
            Return _TargetWork
        End Get
        Set(value As Decimal)
            _TargetWork = value
        End Set
    End Property
    Public Property ActualWork As Decimal
        Get
            Return _ActualWork
        End Get
        Set(value As Decimal)
            _ActualWork = value
        End Set
    End Property
    Public Property RemainingWork As Decimal
        Get
            Return _RemainingWork
        End Get
        Set(value As Decimal)
            _RemainingWork = value
        End Set
    End Property
    Public Property HrsTargetWork As Decimal
        Get
            Return _HrsTargetWork
        End Get
        Set(value As Decimal)
            _HrsTargetWork = value
        End Set
    End Property
    Public Property HrsActualWork As Decimal
        Get
            Return _HrsActualWork
        End Get
        Set(value As Decimal)
            _HrsActualWork = value
        End Set
    End Property
    Public Property HrsRemainingWork As Decimal
        Get
            Return _HrsRemainingWork
        End Get
        Set(value As Decimal)
            _HrsRemainingWork = value
        End Set
    End Property
    Public Property TargetEqpmQty As Decimal
        Get
            Return _TargetEqpmQty
        End Get
        Set(value As Decimal)
            _TargetEqpmQty = value
        End Set
    End Property
    Public Property ActualEqpmQty As Decimal
        Get
            Return _ActualEqpmQty
        End Get
        Set(value As Decimal)
            _ActualEqpmQty = value
        End Set
    End Property
    Public Property RemainingEqpmQty As Decimal
        Get
            Return _RemainingEqpmQty
        End Get
        Set(value As Decimal)
            _RemainingEqpmQty = value
        End Set
    End Property
    Public Property HrsTargetEqpmQty As Decimal
        Get
            Return _HrsTargetEqpmQty
        End Get
        Set(value As Decimal)
            _HrsTargetEqpmQty = value
        End Set
    End Property
    Public Property HrsActualEqpmQty As Decimal
        Get
            Return _HrsActualEqpmQty
        End Get
        Set(value As Decimal)
            _HrsActualEqpmQty = value
        End Set
    End Property
    Public Property HrsRemainingEqpmQty As Decimal
        Get
            Return _HrsRemainingEqpmQty
        End Get
        Set(value As Decimal)
            _HrsRemainingEqpmQty = value
        End Set
    End Property
    Public Property WorkPercComplete As Decimal?
        Get
            Return _WorkPercComplete
        End Get
        Set(value As Decimal?)
            _WorkPercComplete = value
        End Set
    End Property
    Public Property EqpmQtyPercComplete As Decimal?
        Get
            Return _EqpmQtyPercComplete
        End Get
        Set(value As Decimal?)
            _EqpmQtyPercComplete = value
        End Set
    End Property

    Public Property Calendar As P6Calendar
        Get
            Return _Calendar
        End Get
        Set(value As P6Calendar)
            _Calendar = value
        End Set
    End Property

    Public ReadOnly Property ItemPath As List(Of String)
        Get
            Return _ItemPath
        End Get
    End Property
#End Region



End Class

