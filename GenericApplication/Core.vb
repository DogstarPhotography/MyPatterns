Imports System.ComponentModel
Imports System.Security.Principal

''' <summary>
''' Shared methods that are available to all Sirius applications.
''' </summary>
''' <remarks>This library is used by Sirius.IO and Sirius.DataAccess and should be used by added to all 
''' projects using the Sirius N-Tier assemblies.</remarks>
Public Class Core
#Region " Null Field Handling "
    ''' <summary>
    ''' Returns the boolean value or False for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullBool(ByVal value As Object) As Boolean
        If IsDBNull(value) OrElse value Is Nothing Then
            Return False
        End If
        Select Case value.ToString.ToLower
            Case "1", "true", "t"
                Return True
            Case "0", "false", "f"
                Return False
        End Select
        Dim d As Boolean = False : Boolean.TryParse(value.ToString, d)
        Return d
    End Function

    ''' <summary>
    ''' Returns the string passed or "" for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullStr(ByVal value As Object) As String
        If IsDBNull(value) OrElse IsNullString(value) Then Return "" Else Return Convert.ToString(value)
    End Function

    ''' <summary>
    ''' Returns the Char passed or ""c for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullChar(ByVal value As Object) As Char
        If IsDBNull(value) OrElse IsNullString(value) Then Return " "c Else Return Convert.ToChar(value)
    End Function

    ''' <summary>
    ''' Returns the integer passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullInt(ByVal value As Object) As Integer
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToInt32(value)
    End Function

    ''' <summary>
    ''' Returns the integer value passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullInt16(ByVal value As Object) As Int16
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToInt16(value)
    End Function

    ''' <summary>
    ''' Returns the integer value passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullInt64(ByVal value As Object) As Int64
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToInt64(value)
    End Function

    ''' <summary>
    ''' Returns the decimal value passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullDec(ByVal value As Object) As Decimal
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToDecimal(value)
    End Function

    ''' <summary>
    ''' Returns the double passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullDbl(ByVal value As Object) As Double
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToDouble(value)
    End Function

    ''' <summary>
    ''' Returns the byte passed or 0 for a DBNull value.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function NullByte(ByVal value As Object) As Byte
        If IsDBNull(value) OrElse IsNullString(value) Then Return 0 Else Return Convert.ToByte(value)
    End Function

    ''' <summary>
    ''' Returns True if the object is a string, else False.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsNullString(ByVal value As Object) As Boolean
        If Not TypeOf value Is String Then Return False
        Return String.IsNullOrEmpty(CStr(value))
    End Function

    ''' <summary>
    ''' Returns the date passed in string format or Nothing if invalid.
    ''' </summary>
    ''' <param name="value">Valid date in string format for the current culture.</param>
    ''' <returns>Datetime?</returns>
    ''' <remarks>Test for Nothing with "If dat Is Nothing" rather than
    '''          "If dat = Nothing" in code handling DateTime? types.</remarks>
    Public Shared Function NullDate(ByVal value As Object) As DateTime?
        If value Is Nothing Then Return Nothing
        If value.ToString.Trim = "" Then Return Nothing
        Dim d As New DateTime
        If Date.TryParse(value.ToString, d) Then Return d
        Throw New ApplicationException("DateTime? value is not Null and holds an invalid date")
    End Function
#End Region

#Region " Numeric Tests "
    ''' <summary>
    ''' Returns True if the string contains only digits (0-9) else returns False.
    ''' </summary>
    ''' <param name="number"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsDigits(ByVal number As String) As Boolean
        Try
            IsDigits(number, 1, 10)
        Catch ex As Exception
            Dim x As String = ""
        End Try
        Return IsDigits(number, 1, 10)
    End Function

    ''' <summary>
    ''' Returns True if the string contains only digits (0-9) else returns False.
    ''' </summary>
    ''' <param name="number">The string to be tested</param>
    ''' <param name="minlength">Minimum number of characters.</param>
    ''' <param name="maxlength">Maximum number of characters.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsDigits(ByVal number As String, ByVal minlength As Integer, ByVal maxlength As Integer) As Boolean
        If number.Length > maxlength Then Return False
        If number.Length < minlength Then Return False

        For Each ch As Char In number
            Select Case ch
                Case Is < "0"c, Is > "9"c : Return False
            End Select
        Next

        Return True
    End Function
#End Region

#Region " NI Number Validation "
    Private Shared regExNINo As New System.Text.RegularExpressions.Regex("^[A-CEGHJ-PR-TW-Z]{1}[A-CEGHJ-NPR-TW-Z]{1}[0-9]{6}[A-DFM]{1}$")

    ''' <summary>
    ''' Validates whether a text string is a valid NI number.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>True if a valid NINumber was passed, else False.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsValidNINumber(ByVal value As String) As Boolean
        value = value.ToUpper
        If IsTemporaryNINumber(value) Then Return True
        If regExNINo.IsMatch(value) Then Return True
    End Function

    ''' <summary>
    ''' Validates whether a text string is a valid temporary NI Number (ie. T[EN]999999[FM]).
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns>True if the NI Number is a temporary one.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsTemporaryNINumber(ByVal value As String) As Boolean
        If value.Length <> 9 Then Return False
        value = value.ToUpper
        If (value.StartsWith("TN") OrElse value.StartsWith("TE")) AndAlso _
           (value.EndsWith("F") OrElse value.EndsWith("M") OrElse value.EndsWith("A")) AndAlso _
            IsDigits(value.Substring(2, 6)) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region " Config.xml handling "
    Public Const ConfigFileName As String = "Config.xml"

    ''' <summary>
    ''' Returns the value for a key in the Config.xml file.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConfigParameter(ByVal name As String) As String
        Return ConfigParameter(ConfigFileName, name)
    End Function

    ''' <summary>
    ''' Returns the value for a key in the passed config file.
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConfigParameter(ByVal fileName As String, ByVal name As String) As String
        Dim ds As New DataSet
        name = name.ToLower

        Try
            ds.ReadXml(fileName)
            For Each row As DataRow In ds.Tables("Parameter").Rows
                If CStr(row("Name")).ToLower = name Then
                    Return CStr(row("Value"))
                End If
            Next

            Throw New Exception("ConfigParameter is missing: & " & name)
        Catch ex As Exception
            Throw New Exception("ConfigParameter read error: & " & ex.Message)
        Finally
            ds.Dispose()
        End Try
    End Function

    ''' <summary>
    '''  Returns all the keys/values held in the passed config file.
    ''' </summary>
    ''' <param name="name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ConfigParameters(ByVal name As String) As Dictionary(Of String, String)
        Dim ds As New DataSet
        Dim results As New Dictionary(Of String, String)

        Try
            ds.ReadXml(ConfigFileName)
            For Each row As DataRow In ds.Tables("Parameters").Rows
                If CStr(row("Name")).ToLower = name Then
                    results.Add(CStr(row("Name")), CStr(row("Value")))
                End If
            Next
        Catch ex As Exception
            ds.Dispose()
            Throw New Exception("ConfigParameter read error: & " & ex.Message)
        End Try

        Return results
    End Function
#End Region

#Region " Miscellaneous "
    ''' <summary>
    ''' Returns the assembly's program version.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ProgramVersion() As String
        Dim pinfo = Split(System.Reflection.Assembly.GetCallingAssembly.ToString, ",")
        Return pinfo(0) & " Version " & Split(pinfo(1), "=")(1)
    End Function

    ''' <summary>
    ''' This function ensures that any text to displayed as HTML will be correctly 
    ''' formatted. (HTML punctuation is escaped and multiple spaces observed.) 
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns>HTML formatted string</returns>
    ''' <remarks></remarks>
    Public Shared Function FormatTextForHTML(ByVal str As String, Optional ByVal XMLFormat As Boolean = False) As String
        Dim sb As New System.Text.StringBuilder(str.Length + 20)
        Dim lastch As Char = Nothing

        For Each ch As Char In str
            Select Case ch
                Case " "c
                    If XMLFormat Then
                        sb.Append(ch)
                    Else
                        If lastch = " " Then
                            sb.Append("&nbsp;")
                        Else
                            sb.Append(" ")
                        End If
                    End If
                Case "<"c : sb.Append("&lt;")
                Case ">"c : sb.Append("&gt;")
                Case "&"c : sb.Append("&amp;")
                Case """"c : sb.Append("&quot;")
                Case Else : sb.Append(ch)
            End Select
            lastch = ch
        Next

        Return sb.ToString
    End Function

    ''' <summary>
    ''' Returns the date in standard Sirius date format.
    ''' </summary>
    ''' <param name="value">DateTime</param>
    ''' <returns>dd/MM/yyyy</returns>
    ''' <remarks></remarks>
    Public Shared Function FormatDate(ByVal value As DateTime) As String
        Return value.ToShortDateString
    End Function

    ''' <summary>
    '''  Returns the date in standard Sirius date.
    ''' </summary>
    ''' <param name="value">Nullable DateTime</param>
    ''' <returns>dd/MM/yyyy or &lt;Not set&gt;</returns>
    ''' <remarks></remarks>
    Public Shared Function FormatDate(ByVal value As DateTime?) As String
        Return FormatDate(value, "<Not set>")
    End Function

    ''' <summary>
    '''  Returns the date in standard Sirius date. 
    ''' </summary>
    ''' <param name="value">Nullable DateTime</param>
    ''' <param name="notSetString">A string to be displayed if the Date is Null</param>
    ''' <returns>dd/MM/yyyy or user-passed notSetString</returns>
    ''' <remarks></remarks>
    Public Shared Function FormatDate(ByVal value As DateTime?, ByVal notSetString As String) As String
        If value Is Nothing Then Return notSetString
        Return CDate(value).ToShortDateString
    End Function

    ''' <summary>
    ''' Gets the OS version as a string. Because Microsoft do not explicitly state the
    ''' OS it's all a bit fiddly. By using this routine there is only one place to go
    ''' to keep things up to date.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetOSVersion() As String
        Select Case Environment.OSVersion.Platform
            Case PlatformID.Win32S
                Return "Win 3.1"
            Case PlatformID.Win32Windows
                Select Case Environment.OSVersion.Version.Minor
                    Case 0
                        Return "Win95"
                    Case 10
                        Return "Win98"
                    Case 90
                        Return "WinME"
                    Case Else
                        Return "Unknown"
                End Select
            Case PlatformID.Win32NT
                Select Case Environment.OSVersion.Version.Major
                    Case 3
                        Return "NT 3.51"
                    Case 4
                        Return "NT 4.0"
                    Case 5
                        Select Case Environment.OSVersion.Version.Minor
                            Case 0
                                Return "Win2000"
                            Case 1
                                Return "WinXP"
                            Case 2
                                Return "Win2003"
                        End Select
                    Case 6
                        Return "Vista/Win2008Server"
                    Case Else
                        Return "Unknown"
                End Select
            Case PlatformID.WinCE
                Return "Win CE"
        End Select

        Return "Unknown"
    End Function

    ''' <summary>
    ''' Returns the major, minor, revision and build values for the running assembly.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExecutableVersion() As Version
        Dim assembly = System.Reflection.Assembly.GetExecutingAssembly
        Return New Version(Split(Split(assembly.FullName, ",")(1), "=")(1))
    End Function

    ''' <summary>
    ''' Returns the major, minor, revision and build values for the filename/
    ''' </summary>
    ''' <param name="executableName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExecutableVersion(ByVal executableName As String) As Version
        Dim fullName = System.Reflection.AssemblyName.GetAssemblyName(executableName).FullName
        Return New Version(Split(Split(fullName, ",")(1), "=")(1))
    End Function

    ''' <summary>
    ''' Returns the major, minor, revision and build values for the passed assembly.
    ''' </summary>
    ''' <param name="ass"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExecutableVersion(ByVal ass As System.Reflection.Assembly) As Version
        Return New Version(Split(Split(ass.FullName, ",")(1), "=")(1))
    End Function

    ''' <summary>
    ''' Checks that the current user is in an SiriusNET Active Directory group.
    ''' </summary>
    ''' <param name="group">eg. "Domain Users"</param>
    ''' <returns>True if the logged in user is in the specified Active Directory group.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsInADGroup(ByVal group As String) As Boolean
        Return IsInADGroup(group, "SiriusNET")
    End Function

    ''' <summary>
    ''' Checks that the current user is in an Active Directory group for teh required domain.
    ''' </summary>
    ''' <param name="group">eg. "Domain Users"</param>
    ''' <param name="domain">eg. "SiriusNET" (default)</param>
    ''' <returns>True if the logged in user is in the specified Active Directory group.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsInADGroup(ByVal group As String, ByVal domain As String) As Boolean
        Return (New WindowsPrincipal(WindowsIdentity.GetCurrent())).IsInRole(domain & "\" & group)
    End Function

    ''' <summary>
    ''' Checks that the current user is in a local computer group.
    ''' </summary>
    ''' <param name="group">eg. "Users"</param>
    ''' <returns>True if the logged in user is in the specified group.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsInLocalComputerGroup(ByVal group As String) As Boolean
        Return (New WindowsPrincipal(WindowsIdentity.GetCurrent())).IsInRole(group)
    End Function

    ''' <summary>
    ''' GetCommandLineArgumentsClickOnce to allow passing of parameters to other Click Once applications
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCommandLineArgumentsClickOnce() As String()
        If Environment.GetCommandLineArgs.Length = 1 Then
            If AppDomain.CurrentDomain.SetupInformation.ActivationArguments IsNot Nothing _
                    AndAlso AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData IsNot Nothing Then
                Dim aData As String() = AppDomain.CurrentDomain.SetupInformation.ActivationArguments.ActivationData
                Dim splitArgs As String() = Split(aData(0), " ")
                Dim args(splitArgs.Length) As String

                args(0) = "ClickOnceApp"
                For i As Integer = 0 To splitArgs.Length - 1
                    args(i + 1) = splitArgs(i)
                Next

                Return args
            Else
                Return Environment.GetCommandLineArgs
            End If
        Else
            Return Environment.GetCommandLineArgs
        End If
    End Function

    ''' <summary>
    ''' Call any ClickOnce application with or without command line arguments
    ''' </summary>
    ''' <param name="PublisherName"></param>
    ''' <param name="ProductName"></param>
    ''' <param name="Arguments"></param>
    ''' <remarks></remarks>
    Public Shared Sub StartClickOnceApplication(ByVal PublisherName As String, ByVal ProductName As String, Optional ByVal Arguments As String = "")
        Dim FileLocation As String = Environment.GetFolderPath(Environment.SpecialFolder.Programs) & _
            "\" & PublisherName & "\" & ProductName & ".appref-ms"

        If System.IO.File.Exists(FileLocation) = False Then
            MsgBox("The " & ProductName & " application is not present on your system." & vbCrLf & _
                   "Please contact the ICT to update your machine.", MsgBoxStyle.Exclamation, "Missing Application")
            Exit Sub
        End If

        If Arguments = "" Then
            Process.Start(FileLocation)
        Else
            Process.Start(FileLocation, """" & Arguments & """")
        End If
    End Sub

    ''' <summary>
    ''' Determines whether the Windows Forms application is running in Development Studio.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RunningInVisualStudio() As Boolean
        Return Process.GetCurrentProcess.ProcessName.EndsWith(".vshost")
    End Function
#End Region

#Region " Registry Handling "
    ''' <summary>
    ''' Reads a registry string from the Sirius software sub-key requested.
    ''' </summary>
    ''' <param name="app">The application name (eg. "CMS".)</param>
    ''' <param name="key">The subkey name.</param>
    ''' <returns>String value held in the registry key or the null string.</returns>
    ''' <remarks></remarks>
    Public Shared Function ReadRegistryKey(ByVal app As String, ByVal key As String) As String
        Return CStr(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Sirius\" & app, key, ""))
    End Function

    ''' <summary>
    ''' Checks for the existence of an Sirius software registry key.
    ''' </summary>
    ''' <param name="app"></param>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RegistryKeyExists(ByVal app As String, ByVal key As String) As Boolean
        Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Sirius\" & app, key, Nothing) IsNot Nothing
    End Function

    ''' <summary>
    ''' Writes the regitry key for the application, creating sub-keys as required.
    ''' </summary>
    ''' <param name="app">The application name (eg. "CMS".)</param>
    ''' <param name="key">The subkey name.</param>
    ''' <param name="value">String value of the key.</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteRegistryKey(ByVal app As String, ByVal key As String, ByVal value As String)
        If RegistryKeyExists(app, key) = False Then
            If RegistryKeyExists("HKEY_CURRENT_USER\Software\", "Sirius") = False Then
                My.Computer.Registry.CurrentUser.CreateSubKey("HKEY_CURRENT_USER\Software\Sirius")
            End If

            If RegistryKeyExists("HKEY_CURRENT_USER\Software\Sirius", key) = False Then
                My.Computer.Registry.CurrentUser.CreateSubKey("HKEY_CURRENT_USER\Software\Sirius\" & key)
            End If
        End If

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Sirius\" & app, key, value)
    End Sub
#End Region

#Region " Simple Objects for Report Binding "
    ''' <summary>
    ''' "Dictionary" object (String, Decimal) that can be used by the table control in Reporting Services. 
    ''' </summary>
    ''' <remarks>Use System.Generics.Dictionary(Of String, Decimal) instead if possible.</remarks>
    Public Class StringDecimalPair
        Private _Description As String
        Private _Value As Decimal

        ''' <summary>
        ''' Description
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Description() As String
            Get
                Return _Description
            End Get
        End Property

        ''' <summary>
        ''' Decimal value
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value() As Decimal
            Get
                Return _Value
            End Get
        End Property

        Public Sub New(ByVal description As String, ByVal value As Decimal)
            _Description = description
            _Value = value
        End Sub
    End Class

    ''' <summary>
    ''' "Dictionary" object (String, Integer) that can be used by the table control in Reporting Services. 
    ''' </summary>
    ''' <remarks>Use System.Generics.Dictionary(Of String, Integer) instead if possible.</remarks>
    Public Class StringIntegerPair
        Private _Description As String
        Private _Value As Integer

        ''' <summary>
        ''' Description
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Description() As String
            Get
                Return _Description
            End Get
        End Property

        ''' <summary>
        ''' Integer value
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value() As Integer
            Get
                Return _Value
            End Get
        End Property

        Public Sub New(ByVal description As String, ByVal value As Integer)
            _Description = description
            _Value = value
        End Sub
    End Class

    ''' <summary>
    ''' "Dictionary" object (String, String) that can be used by the table control in Reporting Services. 
    ''' </summary>
    ''' <remarks>Use System.Generics.Dictionary(Of String, String) instead if possible.</remarks>
    Public Class StringStringPair
        Private _Description As String
        Private _Value As String

        ''' <summary>
        ''' Description
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Description() As String
            Get
                Return _Description
            End Get
        End Property

        ''' <summary>
        ''' String value
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Value() As String
            Get
                Return _Value
            End Get
        End Property

        Public Sub New(ByVal description As String, ByVal value As String)
            _Description = description
            _Value = value
        End Sub
    End Class
#End Region

#Region " Generic BindingListView class for use with data grids "
    ''' <summary>
    ''' BindingListView class allows filtering and sorting of data within a DataGridView when the
    ''' grid is bound to a generic list. For instance ... 
    ''' 
    '''  dim blv as New BindingListView(Of Staff)(staffReader.AllStaff)
    '''  dgvStaff.DataSource = blv
    ''' 
    ''' Filtering can be achieved as follows:
    '''   
    '''  blv.Filter = "Display Name='John Gardner'" or blv.Filter = "Leave=False"
    ''' 
    ''' Filters are cleared using:
    '''  
    '''  blv.RemoveFilter() or blv.Filter = ""
    ''' </summary>
    ''' <typeparam name="T">Any List of objects</typeparam>
    ''' <remarks></remarks>
    Public Class BindingListView(Of T)
        Inherits BindingList(Of T)
        Implements IBindingListView
        Implements IRaiseItemChangedEvents

        Private _Sorted As Boolean = False
        Private _Filtered As Boolean = False
        Private _FilterString As String = Nothing
        Private _SortDirection As ListSortDirection = ListSortDirection.Ascending
        Private _SortProperty As PropertyDescriptor = Nothing
        Private _SortDescriptions As New ListSortDescriptionCollection
        Private _OriginalCollection As New List(Of T)

        ''' <summary>
        ''' Creates a blank Binding List View for a list of objects suitable for use in a datagrid
        ''' that requires column sorting and filtering.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New()
        End Sub

        ''' <summary>
        ''' Creates a Binding List View for a list of objects suitable for use in a datagrid
        ''' that requires column sorting.
        ''' </summary>
        ''' <param name="list">Existing list of T.</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal list As List(Of T))
            MyBase.New(list)
        End Sub

        Protected Overloads Overrides ReadOnly Property SupportsSearchingCore() As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overloads Overrides Function FindCore(ByVal [property] As PropertyDescriptor, ByVal key As Object) As Integer
            For i As Integer = 0 To Count - 1
                Dim item As T = Me(i)
                If [property].GetValue(item).Equals(key) Then
                    Return i
                End If
            Next
            Return -1
        End Function

        Protected Overloads Overrides ReadOnly Property SupportsSortingCore() As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overloads Overrides ReadOnly Property IsSortedCore() As Boolean
            Get
                Return _Sorted
            End Get
        End Property

        Protected Overloads Overrides ReadOnly Property SortDirectionCore() As ListSortDirection
            Get
                Return _SortDirection
            End Get
        End Property

        Protected Overloads Overrides ReadOnly Property SortPropertyCore() As PropertyDescriptor
            Get
                Return _SortProperty
            End Get
        End Property

        Protected Overloads Overrides Sub ApplySortCore(ByVal [property] As PropertyDescriptor, ByVal direction As ListSortDirection)
            _SortDirection = direction
            _SortProperty = [property]
            Dim comparer As New SortComparer(Of T)([property], direction)
            ApplySortInternal(comparer)
        End Sub

        Private Sub ApplySortInternal(ByVal comparer As SortComparer(Of T))
            If _OriginalCollection.Count = 0 Then _OriginalCollection.AddRange(Me)

            Dim listRef As List(Of T) = TryCast(Me.Items, List(Of T))
            If listRef IsNot Nothing Then
                listRef.Sort(comparer)
                _Sorted = True
                OnListChanged(New ListChangedEventArgs(ListChangedType.Reset, -1))
            End If
        End Sub

        Protected Overloads Overrides Sub RemoveSortCore()
            If Not _Sorted Then
                Clear()
                For Each item As T In _OriginalCollection
                    Add(item)
                Next
                _OriginalCollection.Clear()
                _SortProperty = Nothing
                _SortDescriptions = Nothing
                _Sorted = False
            End If
        End Sub

#Region " IBindingListView Members "
        ''' <summary>
        ''' Sort the List(Of T) according to the sort descriptors specified. (Use GetSortDescriptor() to set up descriptors).
        ''' </summary>
        ''' <param name="sortDescriptors "></param>
        ''' <remarks></remarks>
        Public Sub ApplySort(ByVal sortDescriptors As ListSortDescriptionCollection) Implements IBindingListView.ApplySort
            _SortProperty = Nothing
            _SortDescriptions = sortDescriptors
            Dim comparer As SortComparer(Of T) = New SortComparer(Of T)(sortDescriptors)
            ApplySortInternal(comparer)
        End Sub

        ''' <summary>
        ''' Filter on a property's value, eg. Property='Value'
        ''' </summary>
        ''' <value></value>
        ''' <returns>Filtered List(Of T)</returns>
        ''' <remarks></remarks>
        Public Property Filter() As String Implements IBindingListView.Filter
            Get
                Return _FilterString
            End Get
            Set(ByVal value As String)
                If value = "" Then
                    RemoveFilter()
                Else
                    _FilterString = value
                    _Filtered = True
                    UpdateFilter()
                End If
            End Set
        End Property

        ''' <summary>
        ''' Removes any filter and restores the original List(Of T)
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RemoveFilter() Implements IBindingListView.RemoveFilter
            If Not _Filtered Then
                Return
            End If
            _FilterString = Nothing
            _Filtered = False
            _Sorted = False
            _SortDescriptions = Nothing
            _SortProperty = Nothing
            Clear()
            For Each item As T In _OriginalCollection
                Add(item)
            Next
            _OriginalCollection.Clear()
        End Sub

        ''' <summary>
        ''' Returns the current list of sort descriptors.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SortDescriptions() As ListSortDescriptionCollection Implements IBindingListView.SortDescriptions
            Get
                Return _SortDescriptions
            End Get
        End Property

        ''' <summary>
        ''' Returns True
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SupportsAdvancedSorting() As Boolean Implements IBindingListView.SupportsAdvancedSorting
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Returns True
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property SupportsFiltering() As Boolean Implements IBindingListView.SupportsFiltering
            Get
                Return True
            End Get
        End Property
#End Region

        Private Enum [Operator]
            Equals
            GreaterThan
            LessThan
            GreaterThanOrEqualTo
            LessThanOrEqualTo
            NotEquals
            NotKnown
        End Enum

        Protected Overridable Sub UpdateFilter()
            Dim strings As String() = Nothing
            Dim operatorType As [Operator] = BindingListView(Of T).Operator.NotKnown

            Dim operators As String() = {"=", ">", "<", ">=", "<=", "<>", "!="}
            Dim operatorTypes As [Operator]() = {BindingListView(Of T).Operator.Equals, _
                                                 BindingListView(Of T).Operator.GreaterThan, _
                                                 BindingListView(Of T).Operator.LessThan, _
                                                 BindingListView(Of T).Operator.GreaterThanOrEqualTo, _
                                                 BindingListView(Of T).Operator.LessThanOrEqualTo, _
                                                 BindingListView(Of T).Operator.NotEquals, _
                                                 BindingListView(Of T).Operator.NotEquals}
            For i = 0 To operators.Length - 1
                If _FilterString.Contains(operators(i)) Then
                    strings = Split(_FilterString, operators(i))
                    operatorType = operatorTypes(i)
                End If
            Next

            If operatorType = BindingListView(Of T).Operator.NotKnown OrElse strings.Length <> 2 Then
                Throw New Exception("Filter string must be of the syntax: PropertyName<operator>'Value'")
            End If

            Dim propertyName = strings(0).Trim
            Dim propertyValue = strings(1).Trim.Replace("""", "").Replace("'", "").ToLower
            Dim propertyDescriptor As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))(propertyName)

            If _OriginalCollection.Count = 0 Then _OriginalCollection.AddRange(Me)

            Dim currentCollection As List(Of T) = New List(Of T)(Me)
            Clear()

            Select Case Type.GetTypeCode(propertyDescriptor.PropertyType)
                Case TypeCode.Int64, TypeCode.Int32, TypeCode.Int16, TypeCode.Decimal, _
                     TypeCode.Double, TypeCode.Single, TypeCode.UInt64, TypeCode.UInt32, _
                     TypeCode.Int16, TypeCode.Byte, TypeCode.SByte

                    If IsNumeric(propertyValue) = False Then
                        Throw New Exception("Property value is not valid for numeric type: " & propertyValue)
                    End If

                    Dim numericValue As Double = CDbl(propertyValue)

                    For Each item As T In currentCollection
                        Dim value As Double = CDbl(propertyDescriptor.GetValue(item))
                        Select Case operatorType
                            Case BindingListView(Of T).Operator.Equals : If value = numericValue Then Add(item)
                            Case BindingListView(Of T).Operator.GreaterThan : If value > numericValue Then Add(item)
                            Case BindingListView(Of T).Operator.GreaterThanOrEqualTo : If value >= numericValue Then Add(item)
                            Case BindingListView(Of T).Operator.LessThan : If value < numericValue Then Add(item)
                            Case BindingListView(Of T).Operator.LessThanOrEqualTo : If value <= numericValue Then Add(item)
                            Case BindingListView(Of T).Operator.NotEquals : If value <> numericValue Then Add(item)
                        End Select
                    Next
                Case Else
                    ' For everything else try a string comparison ...
                    For Each item As T In currentCollection
                        Dim value As String = CStr(propertyDescriptor.GetValue(item)).ToLower
                        Select Case operatorType
                            Case BindingListView(Of T).Operator.Equals : If value = propertyValue Then Add(item)
                            Case BindingListView(Of T).Operator.GreaterThan : If value > propertyValue Then Add(item)
                            Case BindingListView(Of T).Operator.GreaterThanOrEqualTo : If value >= propertyValue Then Add(item)
                            Case BindingListView(Of T).Operator.LessThan : If value < propertyValue Then Add(item)
                            Case BindingListView(Of T).Operator.LessThanOrEqualTo : If value <= propertyValue Then Add(item)
                            Case BindingListView(Of T).Operator.NotEquals : If value <> propertyValue Then Add(item)
                        End Select
                    Next
            End Select
        End Sub

#Region " IBindingList overrides "
        Shadows ReadOnly Property AllowNew() As Boolean 'Implements IBindingList.AllowNew
            Get
                Return CheckReadOnly()
            End Get
        End Property

        Shadows ReadOnly Property AllowRemove() As Boolean 'Implements IBindingList.AllowRemove
            Get
                Return CheckReadOnly()
            End Get
        End Property

        Private Function CheckReadOnly() As Boolean
            If _Sorted OrElse _Filtered Then
                Return False
            Else
                Return True
            End If
        End Function
#End Region

        Protected Overloads Overrides Sub InsertItem(ByVal index As Integer, ByVal item As T)
            For Each propertyDescriptor As PropertyDescriptor In TypeDescriptor.GetProperties(item)
                If propertyDescriptor.SupportsChangeEvents Then
                    propertyDescriptor.AddValueChanged(item, AddressOf OnItemChanged)
                End If
            Next

            MyBase.InsertItem(index, item)
        End Sub

        Protected Overloads Overrides Sub RemoveItem(ByVal index As Integer)
            Dim item As T = Items(index)
            Dim propertyDescriptors As PropertyDescriptorCollection = TypeDescriptor.GetProperties(item)
            For Each propertyDescriptor As PropertyDescriptor In propertyDescriptors
                If propertyDescriptor.SupportsChangeEvents Then
                    propertyDescriptor.RemoveValueChanged(item, AddressOf OnItemChanged)
                End If
            Next
            MyBase.RemoveItem(index)
        End Sub

        Sub OnItemChanged(ByVal sender As Object, ByVal args As EventArgs)
            Dim index As Integer = Items.IndexOf(DirectCast(sender, T))
            OnListChanged(New ListChangedEventArgs(ListChangedType.ItemChanged, index))
        End Sub

#Region " IRaiseItemChangedEvents Members "
        ReadOnly Property RaisesItemChangedEvents() As Boolean 'Implements IRaiseItemChangedEvents.RaisesItemChangedEvents
            Get
                Return True
            End Get
        End Property
#End Region

        ''' <summary>
        ''' Returns a sort descriptor for one of the properties of T. 
        ''' </summary>
        ''' <param name="propertyName"></param>
        ''' <param name="direction"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSortDescriptor(ByVal propertyName As String, _
                                          ByVal direction As System.ComponentModel.ListSortDirection) As ListSortDescription
            Dim sortProperty = TypeDescriptor.GetProperties(GetType(T)).Find(propertyName, True)
            If sortProperty Is Nothing Then
                Throw New Exception(GetType(T).ToString & " does not contain the property " & propertyName)
            End If
            Return New ListSortDescription(sortProperty, direction)
        End Function
    End Class

    Private Class SortComparer(Of T)
        Implements IComparer(Of T)

        Private _SortCollection As ListSortDescriptionCollection = Nothing
        Private _PropertyDescriptor As PropertyDescriptor = Nothing
        Private _Direction As ListSortDirection = ListSortDirection.Ascending

        Public Sub New(ByVal propDesc As PropertyDescriptor, ByVal direction As ListSortDirection)
            _PropertyDescriptor = propDesc
            _Direction = direction
        End Sub

        Public Sub New(ByVal sortCollection As ListSortDescriptionCollection)
            _SortCollection = sortCollection
        End Sub

#Region " IComparer(Of T) Members "
        Function Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
            If _PropertyDescriptor IsNot Nothing Then
                ' Simple sort 
                Dim xValue As Object = _PropertyDescriptor.GetValue(x)
                Dim yValue As Object = _PropertyDescriptor.GetValue(y)
                Return CompareValues(xValue, yValue, _Direction)
            Else
                If _SortCollection IsNot Nothing AndAlso _SortCollection.Count > 0 Then
                    Return RecursiveCompareInternal(x, y, 0)
                Else
                    Return 0
                End If
            End If
        End Function
#End Region

        Private Function CompareValues(ByVal xValue As Object, ByVal yValue As Object, ByVal direction As ListSortDirection) As Integer
            Dim returnValue As Integer = 0

            If xValue Is Nothing AndAlso yValue Is Nothing Then Return 0

            If TypeOf xValue Is IComparable Then
                ' Can ask the x value
                If TypeOf xValue Is String AndAlso IsDate(xValue) Then
                    Dim xdate As Object = NullorDate(xValue)
                    Dim ydate As Object = NullorDate(yValue)
                    If ydate Is Nothing Then ydate = DateTime.MinValue
                    returnValue = (DirectCast(xdate, IComparable)).CompareTo(ydate)
                Else
                    If yValue Is Nothing Then
                        returnValue = 1
                    Else
                        returnValue = (DirectCast(xValue, IComparable)).CompareTo(yValue)
                    End If
                End If
            Else
                If TypeOf yValue Is IComparable Then
                    ' Can ask the y value
                    If TypeOf (yValue) Is String AndAlso IsDate(yValue) Then
                        Dim xdate As Object = NullorDate(xValue)
                        Dim ydate As Object = NullorDate(yValue)
                        If xdate Is Nothing Then xdate = DateTime.MinValue
                        returnValue = (DirectCast(ydate, IComparable)).CompareTo(xdate)
                    Else
                        returnValue = (DirectCast(yValue, IComparable)).CompareTo(xValue)
                    End If
                Else
                    If Not xValue.Equals(yValue) Then
                        ' Neither comparable, so compare String representations
                        returnValue = xValue.ToString().CompareTo(yValue.ToString())
                    End If
                End If
            End If

            If direction = ListSortDirection.Ascending Then
                Return returnValue
            Else
                Return returnValue * -1
            End If
        End Function

        Private Function RecursiveCompareInternal(ByVal x As T, ByVal y As T, ByVal index As Integer) As Integer
            If index >= _SortCollection.Count Then Return 0

            Dim listSortDescription As ListSortDescription = _SortCollection(index)
            Dim xValue As Object = listSortDescription.PropertyDescriptor.GetValue(x)
            Dim yValue As Object = listSortDescription.PropertyDescriptor.GetValue(y)
            Dim returnValue As Integer = CompareValues(xValue, yValue, listSortDescription.SortDirection)

            If returnValue = 0 Then
                Return RecursiveCompareInternal(x, y, System.Threading.Interlocked.Increment(index))
            Else
                Return returnValue
            End If
        End Function

        Private Function NullorDate(ByVal value As Object) As DateTime?
            If value Is Nothing Then Return Nothing
            If value.ToString.Trim = "" Then Return Nothing
            Dim d As New DateTime
            If Date.TryParse(value.ToString, d) Then Return d
            Return Nothing
        End Function
    End Class
#End Region

#Region "UserCredentials"
    ''' <summary>
    ''' Provides a means of checking against active directory to prevent access to an application
    ''' </summary>
    ''' <remarks>
    ''' <code>
    ''' credentials = New UserCredentials(Application.ProductName, IsLocal)
    ''' If credentials.NoAccess Then
    '''    MessageBox.Show("You do not have permission to access the " &amp; Application.ProductName &amp; " application. Please consult ICT if this permission is required.", _
    '''                    "Not allowed!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '''    Throw New ApplicationException("No permissions to use " &amp; Application.ProductName &amp; " for " &amp; credentials.ADAccountName)
    ''' End If
    ''' Me.Text &amp;= "    " &amp; VersionInfo() &amp; "   (" &amp; credentials.ToString &amp; ")" &amp; If(IsLocal, "   **Local**", "")
    ''' </code>
    ''' </remarks>
    Public Class UserCredentials
        Private _IsAdministrator As Boolean
        Private _IsInternal As Boolean
        Private _IsExternal As Boolean
        Private _ADAccountName As String

        Public ReadOnly Property IsAdministrator() As Boolean
            Get
                Return _IsAdministrator
            End Get
        End Property

        Public ReadOnly Property IsInternal() As Boolean
            Get
                Return _IsInternal
            End Get
        End Property

        Public ReadOnly Property IsExternal() As Boolean
            Get
                Return _IsExternal
            End Get
        End Property

        Public ReadOnly Property ADAccountName() As String
            Get
                Return _ADAccountName
            End Get
        End Property

        Public ReadOnly Property NoAccess() As Boolean
            Get
                If _IsAdministrator = False AndAlso _IsInternal = False AndAlso _IsExternal = False Then
                    Return True
                Else
                    Return False
                End If
            End Get
        End Property

        Private Sub New()
            'Prevent default constructor
        End Sub

        Public Sub New(ByVal accountPrefix As String, ByVal isLocal As Boolean)
            If isLocal Then
                _IsAdministrator = False
                ' If we are not online, we can't obtain EHD group permissions so grant viewer access.
                _IsInternal = True
                _IsExternal = False
                _ADAccountName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("SiriusNET\", "")
            Else
                _IsAdministrator = IsInADGroup(accountPrefix & "Admin")
                _IsInternal = IsInADGroup(accountPrefix & "Internal")
                _IsExternal = IsInADGroup(accountPrefix & "External")
                _ADAccountName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Replace("SiriusNET\", "")
            End If
        End Sub

        Public Overrides Function ToString() As String
            Return (CStr(IIf(_IsAdministrator, "A", "")) & _
                    CStr(IIf(_IsInternal, "I", "")) & _
                    CStr(IIf(_IsExternal, "E", "")))
        End Function
    End Class

#End Region
End Class