
Public Class YFLogin

#Region "Value properties"
    Private _UserID As Integer
    Private _Username As String
    Private _PasswordHash As Byte()
    Private _Employer As String
    Private _EmployerID As Integer?
    Private _Firstname As String
    Private _Lastname As String
    Private _EmailAddress As String
    Private _SecondaryEmail As String
    Private _LastLogin As DateTime?
    Private _PwdExpiry As DateTime?
    Private _CurrentSession As String
    Private _Validated As Boolean
    Private _BasicAccess As Integer?
    Private _OnlineForms As Integer?
    Private _DataMatch As Integer?
    Private _MemberDetails As Integer?
    Private _Estimates As Integer?
    Private _LG221 As Integer?
    Private _Interface As Integer?
    Private _SiteAdmin As Integer?
    Private _AuthoriseUser As Integer?
    Private _EndofYear As Integer?
    Private _Disabled As Boolean?
    Private _EOYLoginID As Integer
    Private _LoginFails As Integer?
    Private _Deleted As Integer?
    Private _OnlyShowOwnForms As Boolean?
    Private _CaseCode As String
    Private _EmployerSchemePrefix As String
    ''' <summary>
    ''' User ID.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserID() As Integer
        Get
            Return _UserID
        End Get
    End Property
    ''' <summary>
    ''' User name.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Username() As String
        Get
            Return _Username
        End Get
    End Property
    ''' <summary>
    ''' Password hash.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PasswordHash() As Byte()
        Get
            Return _PasswordHash
        End Get
    End Property
    ''' <summary>
    ''' Employer.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Employer() As String
        Get
            Return _Employer
        End Get
    End Property
    ''' <summary>
    ''' Employer ID.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EmployerID() As Integer?
        Get
            Return _EmployerID
        End Get
    End Property
    ''' <summary>
    ''' First name.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Firstname() As String
        Get
            Return _Firstname
        End Get
    End Property
    ''' <summary>
    ''' Last name.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Lastname() As String
        Get
            Return _Lastname
        End Get
    End Property
    ''' <summary>
    ''' Primary email address.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EmailAddress() As String
        Get
            Return _EmailAddress
        End Get
    End Property
    ''' <summary>
    ''' Secondary email address.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SecondaryEmail() As String
        Get
            Return _SecondaryEmail
        End Get
    End Property
    ''' <summary>
    ''' Last login.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastLogin() As DateTime?
        Get
            Return _LastLogin
        End Get
    End Property
    ''' <summary>
    ''' Password expiry.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PwdExpiry() As DateTime?
        Get
            Return _PwdExpiry
        End Get
    End Property
    ''' <summary>
    ''' Current session.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentSession() As String
        Get
            Return _CurrentSession
        End Get
    End Property
    ''' <summary>
    ''' Validated.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Validated() As Boolean
        Get
            Return _Validated
        End Get
    End Property
    ''' <summary>
    ''' Basic access.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property BasicAccess() As Integer?
        Get
            Return _BasicAccess
        End Get
    End Property
    ''' <summary>
    ''' Online forms.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property OnlineForms() As Integer?
        Get
            Return _OnlineForms
        End Get
    End Property
    ''' <summary>
    ''' Datamatch.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DataMatch() As Integer?
        Get
            Return _DataMatch
        End Get
    End Property
    ''' <summary>
    ''' Member details.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MemberDetails() As Integer?
        Get
            Return _MemberDetails
        End Get
    End Property
    ''' <summary>
    ''' Estimates.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Estimates() As Integer?
        Get
            Return _Estimates
        End Get
    End Property
    ''' <summary>
    ''' LG221.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LG221() As Integer?
        Get
            Return _LG221
        End Get
    End Property
    ''' <summary>
    ''' Interface.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property [Interface]() As Integer?
        Get
            Return _Interface
        End Get
    End Property
    ''' <summary>
    ''' Site admin.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SiteAdmin() As Integer?
        Get
            Return _SiteAdmin
        End Get
    End Property
    ''' <summary>
    ''' Authorise user.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property AuthoriseUser() As Integer?
        Get
            Return _AuthoriseUser
        End Get
    End Property
    ''' <summary>
    ''' End of year flag.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EndofYear() As Integer?
        Get
            Return _EndofYear
        End Get
    End Property
    ''' <summary>
    ''' Disabled flag.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Disabled() As Boolean?
        Get
            Return _Disabled
        End Get
    End Property
    ''' <summary>
    ''' End of year login ID.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EOYLoginID() As Integer?
        Get
            Return _EOYLoginID
        End Get
    End Property
    ''' <summary>
    ''' Njumber of failed logins.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LoginFails() As Integer?
        Get
            Return _LoginFails
        End Get
    End Property
    ''' <summary>
    ''' Deleted flag.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Deleted() As Integer?
        Get
            Return _Deleted
        End Get
    End Property
    ''' <summary>
    ''' Only show own forms flag.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property OnlyShowOwnForms() As Boolean?
        Get
            Return _OnlyShowOwnForms
        End Get
    End Property
    ''' <summary>
    ''' Case code.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CaseCode() As String
        Get
            Return _CaseCode
        End Get
        Set(ByVal value As String)
            _CaseCode = value
        End Set
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EmployerSchemePrefix() As String
        Get
            Return _EmployerSchemePrefix
        End Get
        Set(ByVal value As String)
            _EmployerSchemePrefix = value
        End Set
    End Property

#End Region

#Region "Constructors"
    ''' <summary>
    ''' Standard constructor that takes all fields.
    ''' </summary>
    ''' <param name="pUserID"></param>
    ''' <param name="pUsername"></param>
    ''' <param name="pPasswordHash"></param>
    ''' <param name="pEmployer"></param>
    ''' <param name="pEmployerID"></param>
    ''' <param name="pFirstname"></param>
    ''' <param name="pLastname"></param>
    ''' <param name="pEmailAddress"></param>
    ''' <param name="pSecondaryEmail"></param>
    ''' <param name="pLastLogin"></param>
    ''' <param name="pPwdExpiry"></param>
    ''' <param name="pCurrentSession"></param>
    ''' <param name="pValidated"></param>
    ''' <param name="pBasicAccess"></param>
    ''' <param name="pOnlineForms"></param>
    ''' <param name="pDataMatch"></param>
    ''' <param name="pMemberDetails"></param>
    ''' <param name="pEstimates"></param>
    ''' <param name="pLG221"></param>
    ''' <param name="pInterface"></param>
    ''' <param name="pSiteAdmin"></param>
    ''' <param name="pAuthoriseUser"></param>
    ''' <param name="pEndofYear"></param>
    ''' <param name="pDisabled"></param>
    ''' <param name="pEOYLoginID"></param>
    ''' <param name="pLoginFails"></param>
    ''' <param name="pDeleted"></param>
    ''' <param name="pOnlyShowOwnForms"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pUserID As Integer, _
                   ByVal pUsername As String, _
                   ByVal pPasswordHash As Byte(), _
                   ByVal pEmployer As String, _
                   ByVal pEmployerID As Integer?, _
                   ByVal pFirstname As String, _
                   ByVal pLastname As String, _
                   ByVal pEmailAddress As String, _
                   ByVal pSecondaryEmail As String, _
                   ByVal pLastLogin As DateTime?, _
                   ByVal pPwdExpiry As DateTime?, _
                   ByVal pCurrentSession As String, _
                   ByVal pValidated As Boolean, _
                   ByVal pBasicAccess As Integer?, _
                   ByVal pOnlineForms As Integer?, _
                   ByVal pDataMatch As Integer?, _
                   ByVal pMemberDetails As Integer?, _
                   ByVal pEstimates As Integer?, _
                   ByVal pLG221 As Integer?, _
                   ByVal pInterface As Integer?, _
                   ByVal pSiteAdmin As Integer?, _
                   ByVal pAuthoriseUser As Integer?, _
                   ByVal pEndofYear As Integer?, _
                   ByVal pDisabled As Boolean?, _
                   ByVal pEOYLoginID As Integer, _
                   ByVal pLoginFails As Integer?, _
                   ByVal pDeleted As Integer?, _
                   ByVal pOnlyShowOwnForms As Boolean?)
        _UserID = pUserID
        _Username = pUsername
        _PasswordHash = pPasswordHash
        _Employer = pEmployer
        _EmployerID = pEmployerID
        _Firstname = pFirstname
        _Lastname = pLastname
        _EmailAddress = pEmailAddress
        _SecondaryEmail = pSecondaryEmail
        _LastLogin = pLastLogin
        _PwdExpiry = pPwdExpiry
        _CurrentSession = pCurrentSession
        _Validated = pValidated
        _BasicAccess = pBasicAccess
        _OnlineForms = pOnlineForms
        _DataMatch = pDataMatch
        _MemberDetails = pMemberDetails
        _Estimates = pEstimates
        _LG221 = pLG221
        _Interface = pInterface
        _SiteAdmin = pSiteAdmin
        _AuthoriseUser = pAuthoriseUser
        _EndofYear = pEndofYear
        _Disabled = pDisabled
        _EOYLoginID = pEOYLoginID
        _LoginFails = pLoginFails
        _Deleted = pDeleted
        _OnlyShowOwnForms = pOnlyShowOwnForms
    End Sub
    ''' <summary>
    ''' Constructor with limited set of fields to integrate with 'YF_UsersByEmployerV2' stored procedure.
    ''' </summary>
    ''' <param name="pUserID"></param>
    ''' <param name="pUsername"></param>
    ''' <param name="pEmployer"></param>
    ''' <param name="pEmployerID"></param>
    ''' <param name="pFirstname"></param>
    ''' <param name="pLastname"></param>
    ''' <param name="pEmailAddress"></param>
    ''' <param name="pLastLogin"></param>
    ''' <param name="pPwdExpiry"></param>
    ''' <param name="pCurrentSession"></param>
    ''' <param name="pValidated"></param>
    ''' <param name="pBasicAccess"></param>
    ''' <param name="pOnlineForms"></param>
    ''' <param name="pMemberDetails"></param>
    ''' <param name="pSiteAdmin"></param>
    ''' <param name="pAuthoriseUser"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pUserID As Integer, _
                   ByVal pUsername As String, _
                   ByVal pEmployer As String, _
                   ByVal pEmployerID As Integer?, _
                   ByVal pFirstname As String, _
                   ByVal pLastname As String, _
                   ByVal pEmailAddress As String, _
                   ByVal pLastLogin As DateTime?, _
                   ByVal pPwdExpiry As DateTime?, _
                   ByVal pCurrentSession As String, _
                   ByVal pValidated As Boolean, _
                   ByVal pBasicAccess As Integer?, _
                   ByVal pOnlineForms As Integer?, _
                   ByVal pMemberDetails As Integer?, _
                   ByVal pSiteAdmin As Integer?, _
                   ByVal pAuthoriseUser As Integer?)
        'YF_UsersByEmployerV2 fields:
        '[UserID],
        '[Username],
        '[Employer],
        '[EmployerID],
        '[Firstname],
        '[Lastname],
        '[EmailAddress],
        '[LastLogin],
        '[PwdExpiry],
        '[CurrentSession],
        'Convert(BIT, [Validated]) 'Validated',
        'Convert(BIT, [BasicAccess]) 'BasicAccess',
        'Convert(BIT, [OnlineForms]) 'OnlineForms',
        'Convert(BIT, [MemberDetails]) 'MemberDetails',
        'Convert(BIT, [SiteAdmin]) 'SiteAdmin',
        'Convert(BIT, [AuthoriseUser]) 'AuthoriseUser',
        '[Disabled]
        _UserID = pUserID
        _Username = pUsername
        _Employer = pEmployer
        _EmployerID = pEmployerID
        _Firstname = pFirstname
        _Lastname = pLastname
        _EmailAddress = pEmailAddress
        _LastLogin = pLastLogin
        _PwdExpiry = pPwdExpiry
        _CurrentSession = pCurrentSession
        _Validated = pValidated
        _BasicAccess = pBasicAccess
        _OnlineForms = pOnlineForms
        _MemberDetails = pMemberDetails
        _SiteAdmin = pSiteAdmin
        _AuthoriseUser = pAuthoriseUser
    End Sub
    ''' <summary>
    ''' Alternate version for use with YF_UserRightsCheck stored procedure
    ''' </summary>
    ''' <param name="pUserID"></param>
    ''' <param name="pCurrentSession"></param>
    ''' <param name="pBasicAccess"></param>
    ''' <param name="pOnlineForms"></param>
    ''' <param name="pMemberDetails"></param>
    ''' <param name="pSiteAdmin"></param>
    ''' <param name="pAuthoriseUser"></param>
    ''' <param name="pDataMatch"></param>
    ''' <param name="pEstimates"></param>
    ''' <param name="pLG221"></param>
    ''' <param name="pEndofYear"></param>
    ''' <param name="pCaseCode"></param>
    ''' <remarks>
    ''' YF_UserRightsCheck ignores 'HT' users - Hertfordshire
    ''' </remarks>
    Public Sub New(ByVal pUserID As Integer, _
                   ByVal pCurrentSession As String, _
                   ByVal pBasicAccess As Integer?, _
                   ByVal pOnlineForms As Integer?, _
                   ByVal pMemberDetails As Integer?, _
                   ByVal pSiteAdmin As Integer?, _
                   ByVal pAuthoriseUser As Integer?, _
                   ByVal pDataMatch As Integer?, _
                   ByVal pEstimates As Integer?, _
                   ByVal pLG221 As Integer?, _
                   ByVal pEndofYear As Integer?, _
                   ByVal pCaseCode As String)
        'SELECT LG.[UserID], LG.[CurrentSession], LG.[BasicAccess], LG.[OnlineForms], LG.[MemberDetails], LG.[SiteAdmin], 
        '    LG.[AuthoriseUser], LG.[DataMatch], LG.[Estimates], LG.[LG221], LG.[EndOfYear], ET.Case_Code
        'FROM [dbo].[tblYFLogin] LG	INNER JOIN [Sirius-lon-cms-02].[prows2000].dbo.Employer_table ET
        _UserID = pUserID
        _CurrentSession = pCurrentSession
        _BasicAccess = pBasicAccess
        _OnlineForms = pOnlineForms
        _MemberDetails = pMemberDetails
        _SiteAdmin = pSiteAdmin
        _AuthoriseUser = pAuthoriseUser
        _DataMatch = pDataMatch
        _Estimates = pEstimates
        _LG221 = pLG221
        _EndofYear = pEndofYear
        _CaseCode = pCaseCode
    End Sub
    ''' <summary>
    ''' Alternate version for use with YF_CheckLogin_Secure
    ''' </summary>
    ''' <param name="pUserID"></param>
    ''' <param name="pUsername"></param>
    ''' <param name="pEmployer"></param>
    ''' <param name="pFirstname"></param>
    ''' <param name="pLastname"></param>
    ''' <param name="pEmailAddress"></param>
    ''' <param name="pLastLogin"></param>
    ''' <param name="pCurrentSession"></param>
    ''' <param name="pEOYLoginID"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal pUserID As Integer, _
                   ByVal pUsername As String, _
                   ByVal pEmployer As String, _
                   ByVal pFirstname As String, _
                   ByVal pLastname As String, _
                   ByVal pEmailAddress As String, _
                   ByVal pLastLogin As DateTime?, _
                   ByVal pPwdExpiry As DateTime?, _
                   ByVal pCurrentSession As String, _
                   ByVal pValidated As Boolean, _
                   ByVal pEOYLoginID As Integer)
        'YF_CheckLogin_Secure
        'SELECT [UserID], [Username], [Employer], [Firstname], [Lastname], 
        '[EmailAddress], ISNULL([LastLogin], '2005-01-01') 'LastLogin', [PwdExpiry], ISNULL(CurrentSession, '') 'CurrentSession', [Validated], ISNULL([EoyLoginID], 0) 'EoyLoginID'
        'FROM [dbo].[tblYFLogin]
        _UserID = pUserID
        _Username = pUsername
        _Employer = pEmployer
        _Firstname = pFirstname
        _Lastname = pLastname
        _EmailAddress = pEmailAddress
        _LastLogin = pLastLogin
        _PwdExpiry = pPwdExpiry
        _CurrentSession = pCurrentSession
        _Validated = pValidated
        _EOYLoginID = pEOYLoginID
    End Sub
#End Region
End Class


