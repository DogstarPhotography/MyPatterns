Imports System.Xml.Serialization

Public Class Library_Text
    ''' <summary>
    ''' Return a string that is no longer than the max length given
    ''' </summary>
    ''' <param name="TheString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function MaxStringLength(ByVal TheString As String, ByVal MaxLength As Integer) As String
        Dim Result As String = ""
        Try
            Result = TheString.Substring(0, MaxLength)
        Catch aoore As ArgumentOutOfRangeException
            Result = TheString
        Catch ex As Exception
            Throw ex
        End Try
        Return Result
    End Function

    Public Shared Function UrlEncode(ByVal Source As String, Optional ByVal UsePlusForSpace As Boolean = True) As String
        Try
            Dim Result As String = ""
            For Each Character As String In Source
                Select Case Asc(Character)
                    Case 48 To 57, 65 To 90, 97 To 122
                        Result = Result & Character
                    Case 32
                        If UsePlusForSpace = True Then
                            Result = Result & "+"
                        Else
                            Result = Result & "%20"
                        End If
                    Case Else
                        Result = Result & "%" & Hex(Asc(Character)).PadLeft(2, "0"c)
                End Select
            Next Character

            Return Result

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function UrlDecode(ByVal Source As String, Optional ByVal UsePlusForSpace As Boolean = True) As String
        Try
            Dim Result As String = ""
            Dim Index As Integer = 0
            Do While Index < Source.Length
                Dim Character As String = Source.Substring(Index, 1)
                Select Case Character
                    Case "+"
                        If UsePlusForSpace = True Then
                            Result = Result & "+"
                        Else
                            Result = Result & " "
                        End If
                        Index += 1
                    Case "%"
                        Dim num As Integer = CInt(Val("&H" & Source.Substring(Index + 1, 2)))
                        Result = Result & Chr(num)
                        Index += 3
                    Case Else
                        Result = Result & Character
                        Index += 1
                End Select
            Loop
            Return Result
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Serialization"
    ''' <summary>
    ''' Generic method to serialize an object and save it into a file
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Filename"></param>
    ''' <param name="Target"></param>
    ''' <remarks></remarks>
    Public Shared Sub Serialize(Of T As New)(ByVal Filename As String, ByVal Target As T)
        Try
            Dim newWriter As New StreamWriter(Filename)
            Dim newSerializer As New XmlSerializer(GetType(T))
            newSerializer.Serialize(newWriter, Target)
            newWriter.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ''' <summary>
    ''' Generic method to deserialize a file and create an object of the given type from it
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Filename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Deserialize(Of T As New)(ByVal Filename As String) As T
        Try
            Dim newT As T
            'Read file
            Dim newSerializer As New XmlSerializer(GetType(T))
            Dim newReader As New StreamReader(Filename)
            newT = CType(newSerializer.Deserialize(newReader), T)
            'All done
            Return newT
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Generic method to serialize an object and return it as text
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Target"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SerializeText(Of T As New)(ByVal Target As T) As String
        Try
            Dim newWriter As New StringWriter()
            Dim newSerializer As New XmlSerializer(GetType(T))
            newSerializer.Serialize(newWriter, Target)
            newWriter.Close()
            Return newWriter.ToString
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' Generic method to deserialize the given text and create an object of the given type from it
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeserializeText(Of T As New)(ByVal Text As String) As T
        Try
            Dim newT As T
            'Read text
            Dim newSerializer As New XmlSerializer(GetType(T))
            Dim newReader As New StringReader(Text)
            newT = CType(newSerializer.Deserialize(newReader), T)
            'All done
            Return newT
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

    Public Shared Function IsCSV(ByVal CSV As String) As Boolean
        Try
            Dim arrCSV As Char() = CSV.ToCharArray
            For Each curChar As Char In arrCSV
                If Char.IsDigit(curChar) = False AndAlso curChar <> "," Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#Region "Coalesce functions"

    Public Shared Function Coalesce(ByVal ParamArray arguments As Object()) As Object
        Dim argument As Object
        For Each argument In arguments
            If Not argument Is Nothing Then
                Return argument
            End If
        Next
        ' No argument was found that wasn't nothing
        Return Nothing
    End Function

    Public Shared Function Coalesce(ByVal ParamArray parameters As String()) As String
        For Each parameter As String In parameters
            If Not parameter Is Nothing Then
                Return parameter
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
