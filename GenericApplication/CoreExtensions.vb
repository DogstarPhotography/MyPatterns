Imports System.ComponentModel
Imports System.Runtime.CompilerServices

''' <summary>
''' Data type extensions available to all Sirius applications.
''' </summary>
''' <remarks>Add a reference and an <c>Imports Sirius.CoreExtensions</c> line to use in a project.</remarks>
Public Module CoreExtensions
    ''' <summary>
    ''' This procedure gets the description attribute of an enum constant, if any. Otherwise it gets 
    ''' the string name of the enum member.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks>Usage:  myenum.Member.Description()
    ''' Add the Description attribute to each member of the enumeration.</remarks>
    <Extension()> _
    Public Function Description(ByVal value As [Enum]) As String
        Dim fi As Reflection.FieldInfo = value.GetType().GetField(value.ToString())
        Dim aattr() As DescriptionAttribute = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        If aattr.Length > 0 Then
            Return aattr(0).Description
        Else
            Return value.ToString()
        End If
    End Function

    ''' <summary>
    ''' Returns the Sirius standard dd/mm/yyyy format for a datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormat(ByVal value As DateTime) As String
        If value = Nothing Then Return ("<Not set")
        Return CType(value, DateTime).ToString("dd/MM/yyyy")
    End Function

    ''' <summary>
    ''' Returns the Sirius standard dd/mm/yyyy hh:mm format for a datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormatLong(ByVal value As DateTime) As String
        Return value.ToString("dd/MM/yyyy HH:mm")
    End Function

    ''' <summary>
    ''' Returns the Sirius standard dd/mm/yyyy hh:mm:ss format for a datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormatFull(ByVal value As DateTime) As String
        Return value.ToString("dd/MM/yyyy HH:mm:ss")
    End Function

    ''' <summary>
    ''' Return the Sirius standard dd/mm/yyyy format for a nullable datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormat(ByVal value As DateTime?, Optional ByVal nullValue As String = "<Not set>") As String
        If value Is Nothing Then Return nullValue
        Return CType(value, DateTime).SiriusFormat
    End Function

    ''' <summary>
    ''' Returns the Sirius standard dd/mm/yyyy hh:mm format for a nullable datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormatLong(ByVal value As DateTime?, Optional ByVal nullValue As String = "<Not set>") As String
        If value Is Nothing Then Return nullValue
        Return CType(value, DateTime).SiriusFormatLong
    End Function

    ''' <summary>
    ''' Returns the Sirius standard dd/mm/yyyy format hh:mm:ss for a nullable datetime.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusFormatFull(ByVal value As DateTime?, Optional ByVal nullValue As String = "<Not set>") As String
        If value Is Nothing Then Return nullValue
        Return CType(value, DateTime).SiriusFormatFull
    End Function
    ''' <summary>
    ''' Returns an Sirius standard yyyMMdd date stamp for use in filenames, etc.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusDateStamp(ByVal value As DateTime) As String
        Return Format(value, "yyyyMMdd")
    End Function
    ''' <summary>
    ''' Returns an Sirius standard hhmmss time stamp for use in filenames, etc.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusTimeStamp(ByVal value As DateTime) As String
        Return Format(value, "hhmmss")
    End Function
    ''' <summary>
    ''' Returns an Sirius standard yyyMMdd_hhmmss date and time stamp for use in filenames, etc.
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function SiriusDateTimeStamp(ByVal value As DateTime) As String
        Return Format(value, "yyyyMMdd_hhmmss")
    End Function
End Module
