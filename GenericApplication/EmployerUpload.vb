Namespace Opus
    Public Class EmployerUpload

#Region "Value properties"
        Private _ID As Integer
        Private _SenderUserID As Integer
        Private _RecipientUserID As Integer
        Private _FromAddress As String
        Private _ReplyToAddress As String
        Private _ToAddress As String
        Private _Directory As String
        Private _Filename As String
        Private _RandomKey As String
        Private _Status As String
        Private _Updated As DateTime?
        ''' <summary>
        ''' Unique record ID.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ID() As Integer
            Get
                Return _ID
            End Get
        End Property
        ''' <summary>
        ''' UserID of sender.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SenderUserID() As Integer
            Get
                Return _SenderUserID
            End Get
        End Property
        ''' <summary>
        ''' UserId of recipient.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RecipientUserID() As Integer
            Get
                Return _RecipientUserID
            End Get
        End Property
        ''' <summary>
        ''' Email address of sender.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FromAddress() As String
            Get
                Return _FromAddress
            End Get
        End Property
        ''' <summary>
        ''' Emial address to send replies to.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>When blank indicates that replies should be sent to the From address.</remarks>
        Public ReadOnly Property ReplyToAddress() As String
            Get
                Return _ReplyToAddress
            End Get
        End Property
        ''' <summary>
        ''' Email address of recipient.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ToAddress() As String
            Get
                Return _ToAddress
            End Get
        End Property
        ''' <summary>
        ''' Sub directory.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Directory() As String
            Get
                Return _Directory
            End Get
        End Property

        ''' <summary>
        ''' Full filename.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Filename() As String
            Get
                Return _Filename
            End Get
        End Property
        ''' <summary>
        ''' Random key.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property RandomKey() As String
            Get
                Return _RandomKey
            End Get
        End Property

        ''' <summary>
        ''' Status.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Status() As String
            Get
                Return _Status
            End Get
        End Property
        ''' <summary>
        ''' Last update.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Updated() As DateTime?
            Get
                Return _Updated
            End Get
        End Property
#End Region
        ''' <summary>
        ''' Standard constructor.
        ''' </summary>
        ''' <param name="pID"></param>
        ''' <param name="pSenderUserID"></param>
        ''' <param name="pRecipientUserID"></param>
        ''' <param name="pFilename"></param>
        ''' <param name="pStatus"></param>
        ''' <param name="pUpdated"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pID As Integer, _
                       ByVal pSenderUserID As Integer, _
                       ByVal pRecipientUserID As Integer, _
                       ByVal pFromAddress As String, _
                       ByVal pReplyToAddress As String, _
                       ByVal pToAddress As String, _
                       ByVal pDirectory As String, _
                       ByVal pFilename As String, _
                       ByVal pRandomKey As String, _
                       ByVal pStatus As String, _
                       ByVal pUpdated As DateTime?)
            _ID = pID
            _SenderUserID = pSenderUserID
            _RecipientUserID = pRecipientUserID
            _FromAddress = pFromAddress
            _ReplyToAddress = pReplyToAddress
            _ToAddress = pToAddress
            _Directory = pDirectory
            _Filename = pFilename
            _RandomKey = pRandomKey
            _Status = pStatus
            _Updated = pUpdated
        End Sub
        ''' <summary>
        ''' Cut down constructor for use when creating new EmployerUpload objects to be inserted into the database.
        ''' </summary>
        ''' <param name="pSenderUserID"></param>
        ''' <param name="pRecipientUserID"></param>
        ''' <param name="pFromAddress"></param>
        ''' <param name="pReplyToAddress"></param>
        ''' <param name="pToAddress"></param>
        ''' <param name="pFilename"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pSenderUserID As Integer, _
                       ByVal pRecipientUserID As Integer, _
                       ByVal pFromAddress As String, _
                       ByVal pReplyToAddress As String, _
                       ByVal pToAddress As String, _
                       ByVal pDirectory As String, _
                       ByVal pFilename As String)
            _ID = -1
            _SenderUserID = pSenderUserID
            _RecipientUserID = pRecipientUserID
            _FromAddress = pFromAddress
            _ReplyToAddress = pReplyToAddress
            _ToAddress = pToAddress
            _Directory = pDirectory
            _Filename = pFilename
            _RandomKey = Guid.NewGuid.ToString.Replace("-"c, String.Empty).Substring(1, 12)
            _Status = "CREATED"
            _Updated = Now
        End Sub
    End Class

End Namespace

