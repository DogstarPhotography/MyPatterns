Imports System.IO
Imports System.Xml.Serialization

''' <summary>
''' Simple file based event log.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Log
    'Spoiler Information :
    '   EventLogEntryType.Error         = 1
    '   EventLogEntryType.Warning       = 2
    '   EventLogEntryType.Information   = 4

    Public Shared ReadOnly Instance As Log = New Log

    Private MyLogfile As XMLLog
    Private MyLogFilename As String = ""
    Private MySource As String = ""
    Private SyncObject As New Object
    Private MyLogLevel As EventLogEntryType = EventLogEntryType.Error

    Private Sub New()
        'Disabled default constructor
    End Sub

    Public Sub NewLogFile(ByVal LogFilename As String, ByVal Source As String)
        Try
            If LogFilename.Length = 0 Then
                Throw New Exception("LogFilename must be a valid file name.")
            End If
            MyLogFilename = LogFilename
            MySource = Source
            MyLogLevel = EventLogEntryType.Information
            SyncObject = New Object
            'Delete previous logfile
            If File.Exists(MyLogFilename) = True Then
                File.Delete(MyLogFilename)
            End If
            'Write a startup entry
            'WriteEntry("Log file created", EventLogEntryType.Information)
            MyLogfile = New XMLLog
            SyncLock SyncObject
                'Create first entry
                MyLogfile.LogItems.Add(New LogItem(Now, EventLogEntryType.Information, "Log file created"))
                'Save
                Serialize(MyLogFilename, MyLogfile)
            End SyncLock
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

    Public Sub WriteInformation(ByVal Message As String, Optional ByVal EntryType As EventLogEntryType = EventLogEntryType.Information)
        'Don't log stuff we are ignoring
        If LogEntry(EntryType) = False Then Exit Sub
        WriteEntry(Message, EntryType)
    End Sub

    Public Sub WriteWarning(ByVal Message As String, Optional ByVal EntryType As EventLogEntryType = EventLogEntryType.Warning)
        'Don't log stuff we are ignoring
        If LogEntry(EntryType) = False Then Exit Sub
        Message = " Warning: " & Message
        WriteEntry(Message, EntryType)
    End Sub

    Public Sub WriteError(ByVal Message As String, ByVal ex As Exception, Optional ByVal EntryType As EventLogEntryType = EventLogEntryType.Error)
        'Don't log stuff we are ignoring
        If LogEntry(EntryType) = False Then Exit Sub
        Message = Message & " Exception: " & ex.Message
        WriteEntry(Message, EntryType)
    End Sub

    Public Property LogLevel() As EventLogEntryType
        Get
            Return MyLogLevel
        End Get
        Set(ByVal value As EventLogEntryType)
            MyLogLevel = value
        End Set
    End Property

    Private Sub WriteEntry(ByVal Message As String, ByVal EntryType As EventLogEntryType)
        Try
            If MyLogFilename.Length = 0 Then
                Throw New Exception("You must call NewLogFile before attempting to write a log entry.")
            End If
            'Prepare message
            Dim Entry As String
            Select Case EntryType
                Case EventLogEntryType.Error
                    Entry = Format(Now, "hh:mm:ss") & " - " & MySource & " - ERROR - " & Message
                Case Else
                    Entry = Format(Now, "hh:mm:ss") & " - " & MySource & " - " & Message
            End Select
            'Needs to be atomic
            SyncLock SyncObject

                'Read
                'Dim Logfile As XMLLog = Deserialize(MyLogFilename)
                'Append
                MyLogfile.LogItems.Add(New LogItem(Now, EntryType, Message))
                'Save
                Serialize(MyLogFilename, MyLogfile)

            End SyncLock
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

    Private Function LogEntry(ByVal EntryType As EventLogEntryType) As Boolean
        'These are the exclusions
        If MyLogLevel = EventLogEntryType.Warning And EntryType = EventLogEntryType.Information Then Return False
        If MyLogLevel = EventLogEntryType.Error And EntryType = EventLogEntryType.Information Then Return False
        If MyLogLevel = EventLogEntryType.Error And EntryType = EventLogEntryType.Warning Then Return False
        'Log everything else
        Return True
    End Function

    Private Sub Serialize(ByVal Filename As String, ByVal Target As XMLLog)
        Try
            Dim newWriter As New StreamWriter(Filename)
            Dim newSerializer As New XmlSerializer(GetType(XMLLog))
            newSerializer.Serialize(newWriter, Target)
            newWriter.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function Deserialize(ByVal Filename As String) As XMLLog
        Try
            Dim newXMLLog As XMLLog
            'Read file
            Dim newSerializer As New XmlSerializer(GetType(XMLLog))
            Dim newReader As New StreamReader(Filename)
            newXMLLog = CType(newSerializer.Deserialize(newReader), XMLLog)
            'All done
            Return newXMLLog
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Log Subclasses"
    Private Class XMLLog

        <XmlArray("LogItems")> _
        <XmlArrayItem("LogItem", GetType(LogItem))> _
        Public LogItems As System.Collections.Generic.List(Of LogItem) = New System.Collections.Generic.List(Of LogItem)

    End Class

    Private Class LogItem

        <XmlAttributeAttribute()> Public DateTimeStamp As DateTime
        <XmlAttributeAttribute()> Public EntryType As EventLogEntryType
        <XmlAttributeAttribute()> Public Message As String

		Public Sub New(ByVal DateTimeStamp As DateTime, ByVal EntryType As EventLogEntryType, ByVal Message As String)
			Me.DateTimeStamp = DateTimeStamp
			Me.EntryType = EntryType
			Me.Message = Message
		End Sub

    End Class
#End Region
End Class
