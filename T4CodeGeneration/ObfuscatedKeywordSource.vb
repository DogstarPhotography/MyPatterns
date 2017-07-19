
''' <summary>
''' TO DO - Replace Asc, Chr and reference to Microsoft.VisualBasic with System.Text.Encoding.ASCII
''' TO DO - Consider using BitConverter ToChar and GetBytes instead
''' </summary>
Public Class ObfuscatedKeywordSource

    'Private Shared Internal As String = "Th!sIsN0tS3cure!"
    Private Shared Blocks As Integer = 16
    Private Shared BlockSize As Integer = 256
    Private Shared MaxLen As Integer = Blocks * BlockSize
    Private Shared enc As New System.Text.ASCIIEncoding

    Public Shared Function Keyword() As String
        Return AESHMACEncryption.Decrypt(Encrypted, Internal)
    End Function

    Private Shared Function Encrypted() As String
        Dim data As String = ""
        For index = 0 To ExtractEncrypted.GetUpperBound(0)
            Dim item As Byte() = BitConverter.GetBytes(Block(ExtractEncrypted(index)))
            data = data & enc.GetString(item)
        Next
        Return data
    End Function
    Private Shared Function Internal() As String
        Dim data As String = ""
        For index = 0 To ExtractInternal.GetUpperBound(0)
            Dim item As Integer = Block(ExtractInternal(index))
            data = data & Chr(item)
        Next
        Return data
    End Function

    Private Shared Function Block() As Integer()
        Dim data(0 To MaxLen - 1) As Integer
        For index = 0 To BlockSize - 1
            data(index) = Blocks1(index)
        Next
        For index = BlockSize To (BlockSize * 2) - 1
            data(index) = Blocks2(index)
        Next
        For index = (BlockSize * 2) To (BlockSize * 3) - 1
            data(index) = Blocks3(index)
        Next
        'etc
        Return data
    End Function

    Private Shared ExtractEncrypted As Integer() = {0, 1, 2}
    Private Shared ExtractInternal As Integer() = {0, 1, 2}
    Private Shared Blocks1 As Integer() = {0, 0, 0}
    Private Shared Blocks2 As Integer() = {0, 0, 0}
    Private Shared Blocks3 As Integer() = {255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255}
    'etc

    Private Function RandomBlock() As String
        Dim BlockSize As Integer = 256
        Dim Inner As Integer
        Dim rnd As New Random
        Dim result As String = ""
        For Inner = 0 To BlockSize - 1
            result = result & rnd.Next(1, 255) & ", "
        Next
        result = result.Substring(0, result.Length - 2)
        Return result
    End Function

    Private Function ToHashedMappedCSV(ByRef Block() As Integer, Word As String) As String
        'Dim Word As String = "Th1sIsN0t5ecure!"
        'Hash the word
        Dim hshWord As String = Hash.SHA256(Word)
        'Convert the word to an array of integers representing ASCII characters
        Dim ascHashedWord(0 To hshWord.Length - 1) As Integer
        For index = 0 To hshWord.Length - 1
            ascHashedWord(index) = Asc(hshWord.Substring(index, 1))
        Next
        'Build an array that maps the item to a position in the data block
        Dim hpos As Integer
        Dim rnd As New Random
        Dim mapHashedWord(0 To hshWord.Length - 1) As Integer
        For index = 0 To ascHashedWord.GetUpperBound(1) - 1
            hpos = rnd.Next(1, BlockSize)
            hpos = FindNextNumberPos(Block, hpos, ascHashedWord(index))
            If hpos = -1 Then Throw New Exception("Random number generation failed")
            mapHashedWord(index) = hpos
        Next
        'Output the map array as a CSV
        Dim csvOut As String = ""
        For index = 0 To mapHashedWord.GetUpperBound(1) - 1
            csvOut = csvOut & mapHashedWord(index) & ", "
        Next
        csvOut = csvOut.Substring(0, csvOut.Length - 2)
        Return csvOut
    End Function

    Private Function BuildEncrypted() As Integer()
        Dim result As Integer()

    End Function

    Private Function FindNextNumberPos(ByRef source() As Integer, start As Integer, target As Integer) As Integer
        Dim pos As Integer = -1
        For index = start To source.GetUpperBound(1) - 1
            If source(index) = target Then
                pos = index
                Exit For
            End If
        Next
        If pos = -1 Then
            For index = 0 To source.GetUpperBound(1) - 1
                If source(index) = target Then
                    pos = index
                    Exit For
                End If
            Next
        End If
        Return pos
    End Function

End Class
