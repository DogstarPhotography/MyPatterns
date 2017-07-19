' Amendment History:
' 19Jan2010 JG  Created.

Imports System.Data.SqlClient
Imports Sirius.Core
Imports System.IO
Imports System.Collections.Generic

''' <summary>
''' ConnectionManager: Manages the Sirius server environment providing access to the various servers through the 
'''                    use of SQLAccess layers (and connection SQL connection strings for legacy code.
'''
''' Terms used:
''' Environment: A number of servers grouped together and accessible by an application,
'''              e.g. "Development", "Production".
''' Server:      A database server (server name and associated database). A SQL server is
'''              likely to have a number of ConnectionManager servers associated with it.
''' Credentials: A user and (plain or encrypted) password pair.
'''
''' The ConnectionManager must be instantiated with GetInstance() (as there will only ever be one instance of the
''' object), then Intitalise() must be called to set the initial envronment.
'''
''' The simplest way to achive this is with a single line such as:
'''
''' See also Sirius.IO.SQLAccess.
''' <example>
''' Sample code to initialise the ConnectionManager:
''' <code>
''' Private MyConnectionManager As ConnectionManager = _
'''       ConnectionManager.GetInstance.Initialise("Live", ConnectionManager.AccessMode.ReadAccess)
''' </code>
''' </example>
''' <example>
''' Obtaining a SQLAccess instance:
''' <code>
''' Dim MyConnectionManager As ConnectionManager = _
'''       ConnectionManager.GetInstance.Initialise("Live", ConnectionManager.AccessMode.ReadAccess)
''' Dim DBAccess = ConnectionManager.GetInstance.GetSQLAccess("CMS")
''' </code>
''' </example> 
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class ConnectionManager

#Region " Member Variables "
    Private Shared _Instance As ConnectionManager
    Private Shared _InstanceLockObject As New Object

    ' Shared by all users as a singleton instance.
    Private _ServersDS As DataSet
    Private _ServerDT As DataTable
    Private _EnvironmentDT As DataTable

    ' Make connection strings retain its password after connection.
    ' This is required so that SQLConnection.ConnectionString still holds its password after the
    ' object has been used to connect to a database. 
    Private _PersistSecurityInfo As Boolean = True
#End Region

#Region " Construction and Initialisation "

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Allows the single ConnectionManager instance to be obtained.
    ''' </summary>
    ''' <returns>ConnectionManager singleton.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetInstance() As ConnectionManager
        SyncLock _InstanceLockObject
            If _Instance Is Nothing Then
                _Instance = New ConnectionManager
            End If
        End SyncLock

        Return _Instance
    End Function

    Public Enum AccessMode
        ReadAccess = 1
        ReadWriteAccess = 2
        IntegratedSecurity = 3
    End Enum

    ''' <summary>
    ''' Sets the environment and access mode for the Connection Manager. 
    ''' </summary>
    ''' <param name="environment">E.g. "Development", "Live".</param>
    ''' <param name="mode">ReadOnly or ReadWrite.</param>
    ''' <remarks>All databses are opened in the same mode for the duration of the run.</remarks>
    Public Sub Initialise(ByVal environment As String, ByVal mode As AccessMode)
        ReadServersXml()

        _Environment = GetEnvironment(environment)
        If _Environment Is Nothing Then Throw New ApplicationException("Environment '" & environment & "' is not present in Servers.xml")
        _Mode = mode
        _Initialised = True
    End Sub
#End Region

#Region " Properties "
    ' All properties are shared using the singleton instance. If more than one thread is using
    ' the ConnectionManager, then changing environment is ill advised as the effects will be that 
    ' connection objects may be returned for the wrong server.
    Private _AccessCredentials As AccessCredentials = Nothing
    Private _Initialised As Boolean = False
    Private _Environment As ServerEnvironment = Nothing
    Private _Mode As AccessMode = AccessMode.ReadAccess

    ''' <summary>
    ''' User and Password pair.
    ''' </summary>
    ''' <value></value>
    ''' <returns>AccessCredentials object.</returns>
    ''' <remarks>Uses RO/RW settings in Servers.xml if not set.</remarks>
    Public Property AccessCredentials() As AccessCredentials
        Get
            Return _AccessCredentials
        End Get
        Set(ByVal value As AccessCredentials)
            _AccessCredentials = value
        End Set
    End Property

    ''' <summary>
    ''' Returns True if Initialise() has been called.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsInitialised() As Boolean
        Get
            Return _Initialised
        End Get
    End Property

    ''' <summary>
    ''' Returns the current server environment.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Performs an exception if the environment is not known.</remarks>
    Public Property Environment() As String
        Get
            Return ConnectionManager.GetInstance._Environment.Name
        End Get
        Set(ByVal value As String)
            Dim environment As ServerEnvironment = GetEnvironment(value)
            If environment IsNot Nothing Then
                ConnectionManager.GetInstance._Environment = environment
                FetchServerDetails()
            Else
                Throw New ApplicationException("Environment is not present in Servers.xml")
            End If
        End Set
    End Property

    ''' <summary>
    ''' The mode of access i.e. ReadOnly, ReadWrite or IntergratedSecurity.
    ''' </summary>
    ''' <value></value>
    ''' <returns>AccessMode enumeration.</returns>
    ''' <remarks></remarks>
    Public Property Mode() As AccessMode
        Get
            Return _Mode
        End Get
        Set(ByVal value As AccessMode)
            _Mode = value
        End Set
    End Property

    ''' <summary>
    ''' If this flag is set to True then SQL retains the password information in the SQLAccess
    ''' connection string after the connection has been made.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This property does not apply to "Integrated Security" mode. 
    ''' 
    ''' For "ReadOnly" and "ReadWrite" modes, ADO.Net removes the password information contained in the 
    ''' connection string  held in the SQLConnection object after the object has been used to connection 
    ''' to a server for the first time. The result is that the connection string in the SQLConnection 
    ''' object cannot be used to create a second SQLConnection.
    ''' 
    ''' This situation cannot occur if all applications make use of the SQLAccess layer and never 
    ''' attempt to perform ADO.Net operations in higher tier code.
    ''' 
    ''' (PersistSecurityInfo is a prooerty held in the ADO.Net connectino string. For more information
    ''' see http://msdn.microsoft.com/en-us/library/89211k9b(VS.71).aspx.)
    '''</remarks>
    Public Property PersistSecurityInfo() As Boolean
        Get
            Return _PersistSecurityInfo
        End Get
        Set(ByVal value As Boolean)
            _PersistSecurityInfo = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the path name of the Servers.xml configuration file.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks>ConnectrionManager looks to c:\SiriusDev for its configuration file. If the file is
    ''' not found there, (which is the norm for non-development PCs), then the copy on the network
    ''' share is used.
    ''' </remarks>
    Public Shared ReadOnly Property ServersXMLPath() As String
        Get
            Dim serversXMLFile = "Servers.xml"
            Dim configDir = System.Environment.GetEnvironmentVariable("SystemDrive") & "\SiriusDev"
            Dim devXMLFile = configDir & "\" & serversXMLFile

            If Directory.Exists(configDir) AndAlso File.Exists(devXMLFile) Then
                Return devXMLFile
            Else
                Return "\\Sirius\SoftwareDeploy\Configs\" & serversXMLFile
            End If
        End Get
    End Property
#End Region

#Region " Public Methods "
    ''' <summary>
    '''  Returns a SQLAccess object suitable for accessing the database.
    ''' </summary>
    ''' <param name="name">The Name of the SQLAccess object (for debugging purposes).</param>
    ''' <param name="server">The server name e.g. "Common", "Country".</param>
    ''' <returns>SQLAccess object.</returns>
    ''' <remarks>Uses the default country set up by Initialise()</remarks>
    Public Function GetSQLAccess(ByVal name As String, ByVal server As String) As SQLAccess
        CheckInitialised()
        Return GetSQLAccessReal(name, server)
    End Function

    Public Function GetSQLAccess(ByVal server As String) As SQLAccess
        CheckInitialised()
        Return GetSQLAccessReal("", server)
    End Function

    ''' <summary>
    ''' Returns a list of known server environments.
    ''' </summary>
    ''' <returns>List of environment objects.</returns>
    ''' <remarks>An environment typically consists of a Common database, a country database and
    ''' possibly Inventory, Document Image databases. See Servers.xml.</remarks>
    Public Function GetEnvironments() As ServerEnvironments
        CheckInitialised()
        Dim environments As New ServerEnvironments
        For Each erow As DataRow In _EnvironmentDT.Rows
            If Not _EnvironmentDT.Columns.Contains("Path") Then
                environments.Add(New ServerEnvironment(NullStr(erow("Name")), NullStr(erow("Description"))))
            Else
                environments.Add(New ServerEnvironment(NullStr(erow("Name")), NullStr(erow("Description")), NullStr(erow("Path"))))
            End If
        Next
        Return environments
    End Function

    ''' <summary>
    ''' Returns an arraylist of known server environments.
    ''' </summary>
    ''' <returns>List of environment objects.</returns>
    ''' <param name="ExcludeLiveEnvironment"></param>
    ''' <remarks>An environment typically consists of a Common database, a country database and
    ''' possibly Inventory, Document Image databases. See Servers.xml.</remarks>
    Public Function GetEnvironmentNames(ByVal ExcludeLiveEnvironment As Boolean) As List(Of String)
        CheckInitialised()

        Dim environments As New List(Of String)
        For Each erow As DataRow In _EnvironmentDT.Rows
            Dim envname As String = NullStr(erow("Name"))
            If envname.ToLower <> "live" OrElse ExcludeLiveEnvironment = False Then
                environments.Add(NullStr(erow("Name")))
            End If
        Next

        Return environments
    End Function

    ''' <summary>
    ''' Returns an arraylist of known server environments.
    ''' </summary>
    ''' <returns>List of environment objects.</returns>
    ''' <remarks>An environment typically consists of a Common database, a country database and
    ''' possibly Inventory, Document Image databases. See Servers.xml.</remarks>
    Public Function GetEnvironmentNames() As List(Of String)
        Return GetEnvironmentNames(False)
    End Function

    ''' <summary>
    ''' Returns an array of known servers for the environment.
    ''' </summary>
    ''' <returns>Array of server names.</returns>
    ''' <remarks></remarks>
    Public Function GetServersForEnvironment() As String()
        CheckInitialised()
        Dim serverRows As DataRow() = _ServerDT.Select("Environment='" & _Environment.Name & "'")

        Dim servers(serverRows.Length - 1) As String
        For i As Integer = 0 To serverRows.Length - 1
            servers(i) = CStr(serverRows(i)("Name"))
        Next

        Return servers
    End Function

    ''' <summary>
    ''' Tests to see whether a connection can be made to the server/database.
    ''' </summary>
    ''' <param name="server"></param>
    ''' <returns>True is a coenctino can be acheived, else False.</returns>
    ''' <remarks></remarks>
    Public Function TestConnection(ByVal server As String) As Boolean
        CheckInitialised()
        Dim res As Boolean = False
        Dim dbaccess As SQLAccess

        Try
            dbaccess = GetSQLAccessReal("ConnectionManager", server)
            dbaccess.CrashOnError = True
            dbaccess.LogErrors = False

            Dim db As String = CStr(dbaccess.GetScalar(CommandType.Text, "SELECT DB_NAME()"))
            res = True
        Catch ex As Exception
            res = False
        Finally
            dbaccess = Nothing
        End Try

        Return res
    End Function

    ''' <summary>
    ''' Returns the SQL instance and database for the server requested.  
    ''' </summary>
    ''' <param name="server"></param>
    ''' <returns>ServerInfo object.</returns>
    ''' <remarks></remarks>
    Public Function GetServerDetails(ByVal server As String) As ServerInfo
        CheckInitialised()
        Dim cd As ConnectionDetails = GetConnectionStringRecursively(server)
        Return New ServerInfo(cd.Server, cd.DataBase)
    End Function
#End Region

#Region " Private Methods "
    ''' <summary>
    ''' Creates the SQLAccess Instance.
    ''' </summary>
    ''' <param name="name">The name given to the SQLAccess object.</param>
    ''' <param name="server"></param>
    ''' <returns>SQLAccess object of the required server.</returns>
    ''' <remarks>Assumes that the ConnectionManager has been intialised.</remarks>
    Private Function GetSQLAccessReal(ByVal name As String, ByVal server As String) As SQLAccess
        Return New SQLAccess(name, GetConnectionStringRecursively(server).ConnectionString(_PersistSecurityInfo))
    End Function

    ''' <summary>
    ''' Gets the connection string for the required server.
    ''' </summary>
    ''' <param name="server"></param>
    ''' <returns>String.</returns>
    ''' <remarks>Assumes that the ConnectionManager has been intialised.</remarks>
    Private Function GetConnectionStringRecursively(ByVal server As String) As ConnectionDetails
        Dim serverRow As DataRow = GetServer(server)

        If serverRow Is Nothing Then
            Throw New ApplicationException("Server '" & server & "' not found for environment '" & _Environment.Name & "' in " & _
                                           If(ServersXMLPath.StartsWith("\\"), "network ", "local") & " Servers.xml file.")
        End If

        Dim sqlinstance As String = NullStr(serverRow("SQLInstance"))
        Dim userandpassword As String() = Nothing

        If sqlinstance <> "" Then
            Dim database As String = NullStr(serverRow("Database"))
            Dim typeofaccess As String = ""
            Select Case _Mode
                Case AccessMode.ReadAccess
                    typeofaccess = "ReadOnlyAccess"
                Case AccessMode.ReadWriteAccess
                    typeofaccess = "ReadWriteAccess"
                Case AccessMode.IntegratedSecurity
                    typeofaccess = "SSPI"
                    userandpassword = New String() {"Integrated Security", "SSPI"}
            End Select

            If _Mode <> AccessMode.IntegratedSecurity Then
                Dim requiresDecryption As Boolean

                If _AccessCredentials IsNot Nothing Then
                    ' Use credentials provided by caller, not default ones in Servers.xml.
                    userandpassword = _AccessCredentials.ItemArray
                    requiresDecryption = _AccessCredentials.Encrypted
                Else
                    userandpassword = Split(NullStr(serverRow(typeofaccess)), ";")
                    requiresDecryption = True
                End If

                If requiresDecryption Then
                    userandpassword(1) = SiriusEncryption.Decrypt(userandpassword(1))
                End If

                If database = "" OrElse userandpassword.Length <> 2 Then
                    Throw New ApplicationException("Invalid type 1 Server details for " & Environment & "/" & server)
                End If
            End If

            Return New ConnectionDetails(sqlinstance, database, userandpassword(0), userandpassword(1))
        Else
            Throw New ApplicationException("SQL Instance is blank!!!")
        End If
    End Function

    Private Sub CheckInitialised()
        If _Initialised = False Then
            Throw New ApplicationException("Initialise() has not been performed.")
        End If
    End Sub

    Private Function EnvironmentExists(ByVal environment As String) As Boolean
        If _ServerDT.Select("Environment='" & environment & "'").Length = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function GetEnvironment(ByVal environment As String) As ServerEnvironment
        Dim rows As DataRow() = _EnvironmentDT.Select("Name='" & environment & "'")
        If rows.Length = 0 Then
            Return Nothing
        Else
            If Not _EnvironmentDT.Columns.Contains("Path") Then
                Return New ServerEnvironment(CStr(rows(0)("Name")), NullStr(rows(0)("Description")))
            Else
                Return New ServerEnvironment(CStr(rows(0)("Name")), NullStr(rows(0)("Description")), NullStr(rows(0)("Path")))
            End If
        End If
    End Function

    Private Function GetServer(ByVal server As String) As DataRow
        Dim rows As DataRow() = _ServerDT.Select( _
           "Environment='" & _Environment.Name & "' and Name='" & server & "'")
        If rows.Length = 0 Then
            Return Nothing
        Else
            Return rows(0)
        End If
    End Function

    Private Sub FetchServerDetails()
        If _ServersDS IsNot Nothing Then _ServersDS.Dispose()
        _ServersDS = New DataSet
        _ServersDS.ReadXml(ConnectionManager.ServersXMLPath)
        _ServerDT = _ServersDS.Tables("Server")
        _EnvironmentDT = _ServersDS.Tables("Environment")
    End Sub

    Private Shared Sub ReadServersXml()
        Dim path As String = ConnectionManager.ServersXMLPath()

        Try
            With _Instance
                ._ServersDS = New DataSet
                ._ServersDS.ReadXml(path)
                ._ServerDT = ._ServersDS.Tables("Server")
                ._EnvironmentDT = ._ServersDS.Tables("Environment")
            End With

        Catch ex As Exception
            Throw New ApplicationException("Unable to read Servers.xml")
        End Try
    End Sub
#End Region

    Private Class ConnectionDetails
        Private _Server As String
        Private _Database As String
        Private _UserName As String
        Private _Password As String

        Public ReadOnly Property UserName() As String
            Get
                Return _UserName
            End Get
        End Property

        Public ReadOnly Property Password() As String
            Get
                Return _Password
            End Get
        End Property

        Public ReadOnly Property Server() As String
            Get
                Return _Server
            End Get
        End Property

        Public ReadOnly Property DataBase() As String
            Get
                Return _Database
            End Get
        End Property

        Public Sub New(ByVal server As String, ByVal database As String, ByVal username As String, ByVal password As String)
            _Server = server
            _Database = database
            _UserName = username
            _Password = password
        End Sub

        Public Function ConnectionString(ByVal persistPasswordInfo As Boolean) As String
            Dim connection As New System.Text.StringBuilder

            If Password = "SSPI" Then
                connection.Append("Data Source=")
                connection.Append(_Server)
                connection.Append(";")
                connection.Append(_UserName)
                connection.Append("=")
                connection.Append(_Password)
                connection.Append(";Initial Catalog=")
                connection.Append(_Database)
            Else
                If persistPasswordInfo Then connection.Append("Persist Security Info=True;")
                connection.Append("Data Source=")
                connection.Append(_Server)
                connection.Append(";User ID=")
                connection.Append(_UserName)
                connection.Append(";Password=")
                connection.Append(_Password)
                connection.Append(";Initial Catalog=")
                connection.Append(_Database)
            End If

            Return connection.ToString
        End Function
    End Class
End Class

Public Class AccessCredentials
    Private _UserName As String
    Private _Password As String
    Private _Encrypted As Boolean

    Public ReadOnly Property UserName() As String
        Get
            Return _UserName
        End Get
    End Property

    Public ReadOnly Property Password() As String
        Get
            Return _Password
        End Get
    End Property

    Public ReadOnly Property Encrypted() As Boolean
        Get
            Return _Encrypted
        End Get
    End Property

    Public Sub New(ByVal username As String, ByVal password As String, ByVal encrypted As Boolean)
        _UserName = username
        _Password = password
        _Encrypted = encrypted
    End Sub

    Public Function ItemArray() As String()
        Return New String() {_UserName, _Password}
    End Function
End Class

Public Class ServerEnvironment
    Private _Name As String
    Private _Description As String
    Private _Path As String

    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return _Description
        End Get
    End Property

    Public ReadOnly Property Path() As String
        Get
            Return _Path
        End Get
    End Property

    Public Sub New(ByVal name As String, ByVal description As String)
        Me.New(name, description, "")
    End Sub

    Public Sub New(ByVal name As String, ByVal description As String, _
    ByVal path As String)
        _Name = name
        _Description = description
        _Path = path
    End Sub

    Public Overrides Function ToString() As String
        Return _Name.PadRight(12) & _Description
    End Function
End Class

Public Class ServerEnvironments
    Inherits System.Collections.Generic.List(Of ServerEnvironment)

    Public Function FindEnvironment(ByVal Name As String) As ServerEnvironment
        For Each env As ServerEnvironment In Me
            If env.Name = Name Then Return env
        Next
        Return Nothing
    End Function
End Class

Public Class ServerInfo
    Private _SQLInstance As String
    Private _Database As String

    Public ReadOnly Property SQLInstance() As String
        Get
            Return _SQLInstance
        End Get
    End Property

    Public ReadOnly Property Database() As String
        Get
            Return _Database
        End Get
    End Property

    Public Sub New(ByVal sqlinstance As String, ByVal database As String)
        _SQLInstance = sqlinstance
        _Database = database
    End Sub
End Class
