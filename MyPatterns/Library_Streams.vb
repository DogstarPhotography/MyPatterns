Imports System.Text

Module Library_Streams
    ''' <summary>
    ''' Convert a Stream to a Byte Array.
    ''' </summary>
    ''' <param name="source">Stream</param>
    ''' <returns>Byte Array</returns>
    ''' <remarks></remarks>
    Public Function ToByteArray(source As Stream) As Byte()
        Using mem As New MemoryStream()
            source.Position = 0
            source.CopyTo(mem)
            Return mem.ToArray()
        End Using
    End Function
    ''' <summary>
    ''' Convert a Stream to a String using the given encoding.
    ''' </summary>
    ''' <param name="source">Stream</param>
    ''' <param name="encoding">Encoding</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function StreamToString(source As Stream, encoding As Encoding) As String
        'Note that this does not take account of encoding so may cause problems
        Using stream_reader As New StreamReader(source, encoding)
            Dim text As String = stream_reader.ReadToEnd()
            stream_reader.Close()
            Return text
        End Using
    End Function
    ''' <summary>
    ''' Convert a String to a Stream using the given Encoding.
    ''' </summary>
    ''' <param name="text">String</param>
    ''' <param name="encoding">Encoding</param>
    ''' <returns>Stream</returns>
    ''' <remarks></remarks>
    Public Function StringToStream(text As String, encoding As Encoding) As Stream
        Dim bytes As Byte() = encoding.GetBytes(text)
        Dim mem As New MemoryStream(bytes)
        Return mem
    End Function
    ''' <summary>
    ''' Convert a MemoryStream to a String in ASCII encoding.
    ''' </summary>
    ''' <param name="source">MemoryStream</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function StreamToASCIIString(source As MemoryStream) As String
        Return Encoding.ASCII.GetString(source.ToArray)
    End Function
    ''' <summary>
    ''' Convert a String to a MemoryStream (ASCII encoding).
    ''' </summary>
    ''' <param name="text">String</param>
    ''' <returns>MemoryStream</returns>
    ''' <remarks></remarks>
    Public Function ASCIIStringToStream(text As String) As MemoryStream
        Dim bytes As Byte() = Encoding.ASCII.GetBytes(text)
        Dim mem As New MemoryStream(bytes)
        Return mem
    End Function
    ''' <summary>
    ''' Convert a Stream to a String in Base64 encoding.
    ''' </summary>
    ''' <param name="source">Stream</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function StreamToBase64String(source As Stream) As String
        Dim text As String
        Using mem As New MemoryStream()
            Dim buffer As Byte() = New Byte(1024) {}
            Dim bytesRead As Integer
            While bytesRead = source.Read(buffer, 0, buffer.Length)
                mem.Write(buffer, 0, bytesRead)
            End While
            text = Convert.ToBase64String(mem.GetBuffer(), 0, CInt(mem.Length))
        End Using
        Return text
    End Function
    ''' <summary>
    '''  Convert a String to a Stream using Base64 Encoding.
    ''' </summary>
    ''' <param name="text">String</param>
    ''' <returns>Stream</returns>
    ''' <remarks></remarks>
    Public Function Base64StringToStream(text As String) As Stream
        Dim bytes As Byte() = Convert.FromBase64String(text)
        Dim mem As New MemoryStream(bytes)
        Return mem
    End Function
    ''' <summary>
    ''' Return the file's contents in a string.
    ''' </summary>
    ''' <param name="file_name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileContents(ByVal file_name As String) As String
        'See Library_Files for a better implementation of GetFileContents
        Using stream_reader As New StreamReader(file_name)
            Dim text As String = stream_reader.ReadToEnd()
            stream_reader.Close()
            Return text
        End Using
    End Function
    ''' <summary>
    '''  Write a file from a string.
    ''' </summary>
    ''' <param name="file_name"></param>
    ''' <param name="file_contents"></param>
    ''' <remarks></remarks>
    Public Sub SetFileContents(ByVal file_name As String, ByVal file_contents As String)
        'See Library_Files for a better implementation of SetFileContents
        Using stream_writer As New StreamWriter(file_name)
            stream_writer.Write(file_contents)
            stream_writer.Close()
        End Using
    End Sub

End Module
