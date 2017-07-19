Imports System.IO
Imports System.Text
Imports Sirius.Core

Namespace Opus
    ''' <summary>
    ''' Acts as an access layer between the application and the Opus database.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Access
        Dim DBAccess As SQLAccess

#Region " Construction "
        ''' <summary>
        ''' Default constructor.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            DBAccess = ConnectionManager.GetInstance.GetSQLAccess("Opus")
        End Sub
        ''' <summary>
        ''' Constructor that takes a database name.
        ''' </summary>
        ''' <param name="databaseInterface"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal databaseInterface As String)
            DBAccess = ConnectionManager.GetInstance.GetSQLAccess(databaseInterface)
        End Sub
#End Region

#Region "User and Employer Data"
        ''' <summary>
        ''' Get list of users for employer.
        ''' </summary>
        ''' <param name="employerCode"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLogins(ByVal employerCode As String) As List(Of YFLogin)
            Dim list As New List(Of YFLogin)
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "YF_UsersByEmployerV2", _
                                                             New DBParameter() _
                                                             {New DBParameter("@EmployerCode", SqlDbType.VarChar, 5, employerCode), _
                                                              New DBParameter("@DISABLED", SqlDbType.Int, 0, 0)} _
                                                             ).Rows
                list.Add(New YFLogin(NullInt(row("UserID")), NullStr(row("UserName")), _
                                     NullStr(row("Employer")), NullInt(row("EmployerID")), _
                                     NullStr(row("FirstName")), NullStr(row("LastName")), NullStr(row("EmailAddress")), _
                                     NullDate(row("LastLogin")), NullDate(row("PwdExpiry")), NullStr(row("CurrentSession")), NullBool(row("Validated")), _
                                     NullInt(row("BasicAccess")), NullInt(row("OnlineForms")), NullInt(row("MemberDetails")), _
                                     NullInt(row("SiteAdmin")), NullInt(row("AuthoriseUser"))))
            Next
            Return list
        End Function
        ''' <summary>
        ''' Get the team email address for the given user.
        ''' </summary>
        ''' <param name="userID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserSchemeAddress(ByVal userID As Integer) As String
            Return NullStr(DBAccess.GetScalar( _
                CommandType.StoredProcedure, "CMS_GetSchemeEmailAddress", _
                New DBParameter() {New DBParameter("@UserID", SqlDbType.Int, 0, userID)} _
                ))
        End Function
#End Region

#Region "CMSUpload CRUD"
        ''' <summary>
        ''' Fetch a list of employer uploads.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEmployerUploads() As List(Of EmployerUpload)
            Dim list As New List(Of EmployerUpload)
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "CMS_GetEmployerUploads").Rows
                list.Add(New EmployerUpload(NullInt(row("ID")), NullInt(row("SenderUserID")), NullInt(row("RecipientUserID")), _
                                            NullStr(row("FromAddress")), NullStr(row("ReplyToAddress")), NullStr(row("ToAddress")), _
                                            NullStr(row("Directory")), NullStr(row("Filename")), NullStr(row("RandomKey")), NullStr(row("Status")), _
                                            NullDate(row("Updated"))))
            Next
            Return list
        End Function
        ''' <summary>
        ''' Create a new employer upload.
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateEmployerUpload(ByRef item As EmployerUpload) As Boolean
            Dim DBParameterList As New List(Of DBParameter)
            DBParameterList.Add(New DBParameter("@SenderUserID", SqlDbType.Int, 0, item.SenderUserID))
            DBParameterList.Add(New DBParameter("@RecipientUserID", SqlDbType.Int, 0, item.RecipientUserID))
            DBParameterList.Add(New DBParameter("@FromAddress", SqlDbType.VarChar, 50, item.FromAddress))
            DBParameterList.Add(New DBParameter("@ReplyToAddress", SqlDbType.VarChar, 50, item.ReplyToAddress))
            DBParameterList.Add(New DBParameter("@ToAddress", SqlDbType.VarChar, 50, item.ToAddress))
            DBParameterList.Add(New DBParameter("@Directory", SqlDbType.VarChar, 259, item.Directory))
            DBParameterList.Add(New DBParameter("@Filename", SqlDbType.VarChar, 259, item.Filename))
            DBParameterList.Add(New DBParameter("@RandomKey", SqlDbType.VarChar, 12, item.RandomKey))
            DBParameterList.Add(New DBParameter("@Status", SqlDbType.VarChar, 50, item.Status))
            DBAccess.Execute(CommandType.StoredProcedure, "CMS_CreateEmployerUpload", DBParameterList.ToArray)
            If DBAccess.LastReturnCode <> 0 Then
                Throw New ApplicationException("The server returned an error with code: " & DBAccess.LastReturnCode)
            Else
                Return True
            End If
        End Function
        ''' <summary>
        ''' Update all the data for a given employer upload.
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UpdateEmployerUpload(ByRef item As EmployerUpload) As Boolean
            Dim DBParameterList As New List(Of DBParameter)
            DBParameterList.Add(New DBParameter("@UploadID", SqlDbType.Int, 0, item.ID))
            DBParameterList.Add(New DBParameter("@SenderUserID", SqlDbType.Int, 0, item.SenderUserID))
            DBParameterList.Add(New DBParameter("@RecipientUserID", SqlDbType.Int, 0, item.RecipientUserID))
            DBParameterList.Add(New DBParameter("@FromAddress", SqlDbType.VarChar, 50, item.FromAddress))
            DBParameterList.Add(New DBParameter("@ReplyToAddress", SqlDbType.VarChar, 50, item.ReplyToAddress))
            DBParameterList.Add(New DBParameter("@ToAddress", SqlDbType.VarChar, 50, item.ToAddress))
            DBParameterList.Add(New DBParameter("@Directory", SqlDbType.VarChar, 259, item.Directory))
            DBParameterList.Add(New DBParameter("@Filename", SqlDbType.VarChar, 259, item.Filename))
            DBParameterList.Add(New DBParameter("@Status", SqlDbType.VarChar, 50, item.Status))
            DBAccess.Execute(CommandType.StoredProcedure, "CMS_UpdateEmployerUpload", DBParameterList.ToArray)
            If DBAccess.LastReturnCode <> 0 Then
                Throw New ApplicationException("The server returned an error with code: " & DBAccess.LastReturnCode)
            Else
                Return True
            End If
        End Function
        ''' <summary>
        ''' Delete the given employer upload from the database.
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeleteEmployerUpload(ByRef item As EmployerUpload) As Boolean
            Dim DBParameterList As New List(Of DBParameter)
            DBParameterList.Add(New DBParameter("@UploadID", SqlDbType.Int, 0, item.ID))
            DBAccess.Execute(CommandType.StoredProcedure, "CMS_UpdateEmployerUpload", DBParameterList.ToArray)
            If DBAccess.LastReturnCode <> 0 Then
                Throw New ApplicationException("The server returned an error with code: " & DBAccess.LastReturnCode)
            Else
                Return True
            End If
        End Function
#End Region

#Region "Employer Uploads"
        Public Function GetUploadsForRecipient(ByVal UserID As Integer) As List(Of EmployerUpload)
            Dim list As New List(Of EmployerUpload)
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "CMS_GetUploadsForRecipient", _
                                                             New DBParameter() _
                                                             {New DBParameter("@RecipientID", SqlDbType.Int, 0, UserID)} _
                                                             ).Rows
                list.Add(New EmployerUpload(NullInt(row("ID")), NullInt(row("SenderUserID")), NullInt(row("RecipientUserID")), _
                                            NullStr(row("FromAddress")), NullStr(row("ReplyToAddress")), NullStr(row("ToAddress")), _
                                            NullStr(row("Directory")), NullStr(row("Filename")), NullStr(row("RandomKey")), NullStr(row("Status")), _
                                            NullDate(row("Updated"))))
            Next
            Return list
        End Function
#End Region


#Region "Opus AKA YourFund"
#Region "Basic Access"
        Public Function GetUserRightsCheck(ByVal userID As Integer, ByVal session As String) As YFLogin
            Dim list As New List(Of YFLogin)
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "YF_UserRightsCheck", _
                                                             New DBParameter() _
                                                             {New DBParameter("@UserID", SqlDbType.Int, 0, userID), _
                                                              New DBParameter("@Session", SqlDbType.VarChar, 50, session)} _
                                                             ).Rows
                'SELECT LG.[UserID], LG.[CurrentSession], LG.[BasicAccess], LG.[OnlineForms], LG.[MemberDetails], LG.[SiteAdmin], 
                '    LG.[AuthoriseUser], LG.[DataMatch], LG.[Estimates], LG.[LG221], LG.[EndOfYear], ET.Case_Code
                'FROM [dbo].[tblYFLogin] LG	INNER JOIN [Sirius-lon-cms-02].[prows2000].dbo.Employer_table ET
                list.Add(New YFLogin(NullInt(row("UserID")), NullStr(row("CurrentSession")), NullInt(row("BasicAccess")), _
                                     NullInt(row("OnlineForms")), NullInt(row("MemberDetails")), NullInt(row("SiteAdmin")), _
                                     NullInt(row("AuthoriseUser")), NullInt(row("DataMatch")), NullInt(row("Estimates")), _
                                     NullInt(row("LG221")), NullInt(row("EndOfYear")), NullStr(row("Case_Code"))))
            Next
            Return list.FirstOrDefault
        End Function

        Public Function CheckLoginSecure(ByVal username As String, ByVal passwordsecure As Byte(), ByVal employer As String) As YFLogin
            Dim list As New List(Of YFLogin)
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "YF_CheckLogin_Secure", _
                                                             New DBParameter() _
                                                             {New DBParameter("@username", SqlDbType.VarChar, 20, username), _
                                                                New DBParameter("@passwordsecure", SqlDbType.Binary, 16, passwordsecure), _
                                                                New DBParameter("@employer", SqlDbType.VarChar, 5, employer)} _
                                                                ).Rows
                'SELECT [UserID], [Username], [Employer], [Firstname], [Lastname], 
                '[EmailAddress], ISNULL([LastLogin], '2005-01-01') 'LastLogin', [PwdExpiry], ISNULL(CurrentSession, '') 'CurrentSession', [Validated], ISNULL([EoyLoginID], 0) 'EoyLoginID'
                'FROM [dbo].[tblYFLogin]
                list.Add(New YFLogin(NullInt(row("UserID")), NullStr(row("Username")), NullStr(row("Employer")), _
                                     NullStr(row("Firstname")), NullStr(row("Lastname")), NullStr(row("EmailAddress")), _
                                     NullDate(row("LastLogin")), NullDate(row("PwdExpiry")), NullStr(row("CurrentSession")), _
                                     NullBool(row("Validated")), NullInt(row("EoyLoginID"))))
            Next
            Return list.FirstOrDefault
        End Function

        Public Function UpdateLogin(ByVal userID As Integer, ByVal session As String) As Integer
            DBAccess.Execute(CommandType.StoredProcedure, "YF_UserLogIn", _
                             New DBParameter() _
                             {New DBParameter("@userID", SqlDbType.Int, 0, userID), _
                              New DBParameter("@session", SqlDbType.VarChar, 50, session)})
            Return DBAccess.LastReturnCode
        End Function

        Public Function AllowAccess(ByVal userID As Integer, ByVal session As String, ByVal mode As Integer) As Integer
            'Parameters
            Dim DBParameterList As New List(Of DBParameter)
            DBParameterList.Add(New DBParameter("@userID", SqlDbType.Int, 0, userID))
            DBParameterList.Add(New DBParameter("@session", SqlDbType.VarChar, 50, session))
            DBParameterList.Add(New DBParameter("@mode", SqlDbType.Int, 0, mode))
            'Output Parameters
            Dim ReturnVal = New DBParameter("@ReturnVal", SqlDbType.Int, 0, Nothing, True) : DBParameterList.Add(ReturnVal)
            'Call the procedure passing the parameters
            DBAccess.Execute(CommandType.StoredProcedure, "YF_UserAccess", DBParameterList.ToArray)
            Return CInt(ReturnVal.ParameterValue)
        End Function
#End Region

#Region "Online Forms"
#End Region

#Region "Member Details"
#End Region

#Region "Site Admin"
#End Region

#Region "Authorise User"
#End Region

#Region "Data Match (DM)"
#End Region

#Region "Estimates"
#End Region

#Region "LG221"
#End Region

#Region "End of Year (EOY)"
        Public Function GetEmployerDetail(ByVal username As String) As EmployerDetail
            Dim list As New List(Of EmployerDetail)
            'SELECT ET.EmployerID, ET.SchemeID, ET.Employer_ID, ET.Employer_Desc, SD.SchemeDescription, SD.SchemeCaseFilter
            'FROM [Sirius-LON-CMS-01].[prows2000].dbo.[Employer_Table] ET
            '	INNER JOIN [Sirius-LON-CMS-01].[prows2000].dbo.[SchemeDetails] SD 
            '	ON ET.SchemeID = SD.SchemeID
            'WHERE ET.AxisCode = @Employer AND Employer_Desc NOT LIKE '%TEST%' 
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "EOY_EmployerDetails", _
                                                             New DBParameter() _
                                                             {New DBParameter("@Employer", SqlDbType.VarChar, 5, username)} _
                                                                ).Rows
                list.Add(New EmployerDetail(NullStr(row("EmployerID")), NullStr(row("SchemeID")), NullStr(row("Employer_ID")), _
                                            NullStr(row("Employer_Desc")), NullStr(row("SchemeDescription")), NullStr(row("SchemeCaseFilter"))))
            Next
            Return list.FirstOrDefault
        End Function

        Public Function GetFilesLoadedCombined(ByVal employer As String) As Dictionary(Of String, Integer)
            Dim dict As New Dictionary(Of String, Integer)
            'SELECT @EOYFileCount = count(*) FROM dbo.EOY_BatchView WHERE Employer = @EmpCode AND Valid >= 1 AND PeriodEnd = @ENDDATE
            'SELECT @ContsFileCount = COUNT(*) FROM dbo.[tblEOYContsUpload] WHERE Employer = @EmpCode AND PeriodStart = @STARTDATE
            'SELECT @EOYFileCount 'EOYFileCount', @ContsFileCount 'ContsFileCount'
            Dim table As DataTable = DBAccess.GetDatatable(CommandType.StoredProcedure, "EOY_FilesLoadedCombined", _
                                                             New DBParameter() _
                                                             {New DBParameter("@EmpCode", SqlDbType.VarChar, 5, employer)})
            If table.Rows.Count > 0 Then
                dict.Add("EOYFileCount", NullInt(table.Rows.Item(0)("EOYFileCount")))
                dict.Add("ContsFileCount", NullInt(table.Rows.Item(0)("ContsFileCount")))
            Else
                dict.Add("EOYFileCount", 0)
                dict.Add("ContsFileCount", 0)
            End If
            Return dict
        End Function

        Public Function GetExclusionsDescriptions() As List(Of String)
            Dim list As New List(Of String)
            'SELECT '(' + CAST(ET.ExclusionID AS VARCHAR(2)) + ') ' + ET.Reason AS Exclusion
            'FROM [Sirius-lon-cms-01].PROWS2000.dbo.EOYExclusionTypes ET
            'ORDER BY ET.ExclusionID
            For Each row As DataRow In DBAccess.GetDatatable(CommandType.StoredProcedure, "EOY_ExclusionsDescriptions").Rows
                list.Add(NullStr(row("Exclusion")))
            Next
            Return list
        End Function
#End Region

#Region "Other"
#End Region
#End Region
    End Class
End Namespace