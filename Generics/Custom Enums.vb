''' <summary>
''' 
''' </summary>
''' <typeparam name="T"></typeparam>
''' <remarks>See http://stackoverflow.com/questions/940002/getting-static-field-values-of-a-type-using-reflection/940340#940340</remarks>
Public Interface ICustomEnum(Of T)
    Function FromT(ByVal value As T) As ICustomEnum(Of T)
    ReadOnly Property Value() As T
    ''// Implement using a private constructor that accepts and sets the Value property, 
    ''// one shared readonly property for each desired value in the enum,
    ''// and widening conversions to and from T.
    ''// Then see this link to get intellisense support
    ''// that exactly matches a normal enum:
    ''// http://stackoverflow.com/questions/102084/hidden-features-of-vb-net/102217#102217
End Interface

''' <completionlist cref="ReasonCodeValue"/>
Public NotInheritable Class ReasonCodeValue
    Implements ICustomEnum(Of String)

    Private _value As String
    Public ReadOnly Property Value() As String Implements ICustomEnum(Of String).Value
        Get
            Return _value
        End Get
    End Property

    Private Sub New(ByVal value As String)
        _value = value
    End Sub

    Private Shared _ServiceNotCovered As New ReasonCodeValue("SNCV")
    Public Shared ReadOnly Property ServiceNotCovered() As ReasonCodeValue
        Get
            Return _ServiceNotCovered
        End Get
    End Property

    Private Shared _MemberNotEligible As New ReasonCodeValue("MNEL")
    Public Shared ReadOnly Property MemberNotEligible() As ReasonCodeValue
        Get
            Return _MemberNotEligible
        End Get
    End Property

    Public Shared Function FromString(ByVal value As String) As ICustomEnum(Of String)
        ''// use reflection or a dictionary here if you have a lot of values
        Select Case value
            Case "SNCV"
                Return _ServiceNotCovered
            Case "MNEL"
                Return _MemberNotEligible
            Case Else
                Return Nothing ''//or throw an exception
        End Select
    End Function

    Public Function FromT(ByVal value As String) As ICustomEnum(Of String) Implements ICustomEnum(Of String).FromT
        Return FromString(value)
    End Function

    Public Shared Widening Operator CType(ByVal item As ReasonCodeValue) As String
        Return item.Value
    End Operator

    Public Shared Widening Operator CType(ByVal value As String) As ReasonCodeValue
        Return DirectCast(FromString(value), ReasonCodeValue)
    End Operator
End Class