Namespace Opus
    Public Class EmployerDetail
        'ET.EmployerID, ET.SchemeID, ET.Employer_ID, ET.Employer_Desc, SD.SchemeDescription, SD.SchemeCaseFilter

#Region "Properties"
        Private _EmployerID As String
        Public Property EmployerID() As String
            Get
                Return _EmployerID
            End Get
            Set(ByVal value As String)
                _EmployerID = value
            End Set
        End Property

        Private _SchemeID As String
        Public Property SchemeID() As String
            Get
                Return _SchemeID
            End Get
            Set(ByVal value As String)
                _SchemeID = value
            End Set
        End Property

        Private _Employer_ID As String
        Public Property Employer_ID() As String
            Get
                Return _Employer_ID
            End Get
            Set(ByVal value As String)
                _Employer_ID = value
            End Set
        End Property

        Private _Employer_Desc As String
        Public Property Employer_Desc() As String
            Get
                Return _Employer_Desc
            End Get
            Set(ByVal value As String)
                _Employer_Desc = value
            End Set
        End Property

        Private _SchemeDescription As String
        Public Property SchemeDescription() As String
            Get
                Return _SchemeDescription
            End Get
            Set(ByVal value As String)
                _SchemeDescription = value
            End Set
        End Property

        Private _SchemeCaseFilter As String
        Public Property SchemeCaseFilter() As String
            Get
                Return _SchemeCaseFilter
            End Get
            Set(ByVal value As String)
                _SchemeCaseFilter = value
            End Set
        End Property
#End Region

#Region "Constructors"
        Public Sub New(ByVal EmployerID As String, _
                       ByVal SchemeID As String, _
                       ByVal Employer_ID As String, _
                       ByVal Employer_Desc As String, _
                       ByVal SchemeDescription As String, _
                       ByVal SchemeCaseFilter As String)
            _EmployerID = EmployerID
            _SchemeID = SchemeID
            _Employer_ID = Employer_ID
            _Employer_Desc = Employer_Desc
            _SchemeDescription = SchemeDescription
            _SchemeCaseFilter = SchemeCaseFilter
        End Sub
#End Region
    End Class
End Namespace
