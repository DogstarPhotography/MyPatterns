''' <summary>
''' Used to pass SQL parameters for the dynamic SQL/stored procedure.
''' </summary>
''' <remarks></remarks>
Public Class DBParameter
    Private _Name As String
    ''' <summary>
    ''' The name of the parameter.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Private _DBType As SqlDbType
    ''' <summary>
    ''' The type of the parameter.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DBType() As SqlDbType
        Get
            Return _DBType
        End Get
        Set(ByVal value As SqlDbType)
            _DBType = DBType
        End Set
    End Property

    Private _Length As Integer
    ''' <summary>
    ''' The length or size of the parameter
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Length() As Integer
        Get
            Return _Length
        End Get
        Set(ByVal value As Integer)
            _Length = Length
        End Set
    End Property

    Private _ParameterValue As Object = Nothing
    ''' <summary>
    ''' The actual value of the parameter.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ParameterValue() As Object
        Get
            Return _ParameterValue
        End Get
        Set(ByVal value As Object)
            _ParameterValue = value
        End Set
    End Property

    Private _ForOutput As Boolean = False
    ''' <summary>
    ''' Determines if this parameter is an output.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ForOutput() As Boolean
        Get
            Return _ForOutput
        End Get
    End Property
    ''' <summary>
    ''' Default constructor with minimum parameters.
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="DBType"></param>
    ''' <param name="Length"></param>
    ''' <param name="ParameterValue"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String, ByVal DBType As SqlDbType, ByVal Length As Integer, ByVal ParameterValue As Object)
        _Name = Name
        _DBType = DBType
        _Length = Length
        _ParameterValue = ParameterValue
    End Sub
    ''' <summary>
    ''' Default constructor with full parameters.
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="DBType"></param>
    ''' <param name="Length"></param>
    ''' <param name="ParameterValue"></param>
    ''' <param name="forOutput"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Name As String, ByVal DBType As SqlDbType, ByVal Length As Integer, ByVal ParameterValue As Object, ByVal forOutput As Boolean)
        _Name = Name
        _DBType = DBType
        _Length = Length

        ' ADO.Net will not pass the value when the parameter is defined 'for output' even though
        ' SQL stored procedures are happy with OUTPUT parameters being used in both directions. 
        '
        ' _ParameterValue = ParameterValue

        If forOutput AndAlso ParameterValue IsNot Nothing Then
            Throw New Exception("Not allowed to set the parameter value for an output parameter." & ControlChars.CrLf & _
                                "Use two separate parameters - one for input and the other for output.")
        End If

        _ForOutput = forOutput
    End Sub
End Class
