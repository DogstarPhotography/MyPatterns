Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Linq

Module Library_Files
    ''' <summary>
    ''' Remove any read only flags from a directory and the files within it.
    ''' </summary>
    ''' <param name="folder">String</param>
    ''' <remarks></remarks>
    Public Sub RemoveReadOnly(folder As String)
        Dim di As New DirectoryInfo(folder)
        di.Attributes = di.Attributes And Not FileAttributes.[ReadOnly]
        For Each entry As FileInfo In di.GetFiles
            entry.IsReadOnly = False
        Next
    End Sub
    ''' <summary>
    ''' Get a list of files.
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="filter"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFiles(path As String, filter As String) As List(Of String)
        Dim info As New IO.DirectoryInfo(path)
        Dim files As IO.FileInfo() = info.GetFiles(filter)
        Return (From i In files Select i.Name).ToList
    End Function
    ''' <summary>
    ''' Get a temporary folder.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>See http://stackoverflow.com/questions/16656/creating-temporary-folders</remarks>
    Public Function GetTempFolder() As String
        Dim folder As String = Path.Combine(Path.GetTempPath, Path.GetRandomFileName)
        Do While Directory.Exists(folder) Or File.Exists(folder)
            folder = Path.Combine(Path.GetTempPath, Path.GetRandomFileName)
        Loop
        Return folder
    End Function
    ''' <summary>
    ''' Read the actual bytes of a file.
    ''' </summary>
    ''' <param name="filepath"></param>
    ''' <returns></returns>
    ''' <remarks>For certain text files this will include the Byte Order Marker (BOM) at the beginning of the file so may produce unexpected results.
    ''' See http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html</remarks>
    Public Function GetBinaryFileBytes(filepath As String) As Byte()
        Using fs As New FileStream(filepath, FileMode.Open, FileAccess.Read), br As New BinaryReader(fs)
            Dim data(CInt(fs.Length)) As Byte
            data = br.ReadBytes(CInt(fs.Length))
            Return data
        End Using
    End Function
    ''' <summary>
    ''' Read the bytes of a text file without the Byte Order Marker(BOM).
    ''' </summary>
    ''' <param name="filepath"></param>
    ''' <returns></returns>
    ''' <remarks>See http://andrewmatthewthompson.blogspot.co.uk/2011/02/byte-order-mark-found-using-net.html</remarks>
    Public Function GetTextFileBytes(filepath As String) As Byte()
        Using fs As New FileStream(filepath, FileMode.Open, FileAccess.Read), sr As New StreamReader(fs, True)
            Dim data As String = sr.ReadToEnd
            Return Encoding.UTF8.GetBytes(data)
        End Using
    End Function
    ''' <summary>
    ''' Return the file's contents in a string.
    ''' </summary>
    ''' <param name="file_name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileContents(ByVal file_name As String) As String
        Return File.ReadAllText(file_name) 'This is MUCH easier than a block of 'Using StreamWriter' code
        'Using stream_reader As New StreamReader(file_name)
        '    Dim text As String = stream_reader.ReadToEnd()
        '    stream_reader.Close()
        '    Return text
        'End Using
    End Function
    ''' <summary>
    '''  Write a file from a string.
    ''' </summary>
    ''' <param name="file_name"></param>
    ''' <param name="file_contents"></param>
    ''' <remarks></remarks>
    Public Sub SetFileContents(ByVal file_name As String, ByVal file_contents As String)
        File.WriteAllText(file_name, file_contents) 'This is MUCH easier than a block of 'Using StreamWriter' code
        'Using stream_writer As New StreamWriter(file_name)
        '    stream_writer.Write(file_contents)
        '    stream_writer.Close()
        'End Using
    End Sub

    Public Function GetfileBytes(filename As String) As Byte()
        Return File.ReadAllBytes(filename) 'This is MUCH easier than a block of 'Using FileStream' code
    End Function




End Module
