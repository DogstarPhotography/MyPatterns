Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Public Class SiriusEncryption
    ' The encrytion/decryption pair used TripleDES and supports any string. Output is a string of
    ' printable characters suitable for incorporating into a configuration file. 

    ''' <summary>
    ''' Takes as string and returns an encrypted (printable) string, suitable for placing in a config file.
    ''' </summary>
    ''' <param name="plainText"></param>
    ''' <returns>Hexadecimal string</returns>
    ''' <remarks>Use the Decrypt methid to get the original string back again.</remarks>
    Public Shared Function Encrypt(ByVal plainText As String) As String
        Dim encryptedbytes As Byte() = TripleDES.Encrypt(plainText)
        Dim s As New StringBuilder(encryptedbytes.Length * 2)

        ' Guarantee that the string returned is printable - Hex format will do ...
        For Each b As Byte In encryptedbytes
            Dim hexval As String = Hex(b)
            If hexval.Length = 1 Then s.Append("0")
            s.Append(Hex(b))
        Next

        Return s.ToString
    End Function

    ''' <summary>
    ''' Takes as string created by Encrypt() and returns an unencrypted string.
    ''' </summary>
    ''' <param name="input"></param>
    ''' <returns>Plaintext string.</returns>
    ''' <remarks></remarks>
    Public Shared Function Decrypt(ByVal input As String) As String
        Dim b(CInt(input.Length / 2) - 1) As Byte

        ' Convert the Hex string to its original byte format ...
        For i As Integer = 0 To input.Length - 1 Step 2
            b(CInt(i / 2)) = Convert.ToByte(input(i) & input(i + 1), 16)
        Next

        Return TripleDES.Decrypt(b)
    End Function

    ''' <summary>
    ''' The algoithm makes use of the symmetric TripleDES functions supported in the .Net Farmework.
    ''' </summary>
    ''' <remarks></remarks>
    Private Class TripleDES
        Private Const IVMaxIndex As Integer = 23
        Private Shared rgbKey As Byte()
        Private Shared rgbIV As Byte()

        Shared Sub New()
            Dim k As Byte() = GetKey()
            Dim i As Byte() = {65, 110, 68, 26, 69, 178, 200, 219}
            rgbKey = k
            rgbIV = i
        End Sub

        Public Shared Function Encrypt(ByVal plainText As String) As Byte()
            Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
            Dim inputInBytes() As Byte = utf8encoder.GetBytes(plainText)
            Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
            Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateEncryptor(rgbKey, rgbIV)
            Dim encryptedStream As MemoryStream = New MemoryStream()
            Dim cryptStream As CryptoStream = New CryptoStream(encryptedStream, cryptoTransform, CryptoStreamMode.Write)

            cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
            cryptStream.FlushFinalBlock()
            encryptedStream.Position = 0

            Dim result(CInt(encryptedStream.Length - 1)) As Byte
            encryptedStream.Read(result, 0, CInt(encryptedStream.Length))
            cryptStream.Close()
            Return result
        End Function

        Public Shared Function Decrypt(ByVal inputInBytes() As Byte) As String
            Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
            Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
            Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateDecryptor(rgbKey, rgbIV)

            Dim decryptedStream As MemoryStream = New MemoryStream()
            Dim cryptStream As CryptoStream = New CryptoStream(decryptedStream, cryptoTransform, CryptoStreamMode.Write)
            cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
            cryptStream.FlushFinalBlock()
            decryptedStream.Position = 0

            Dim result(CInt(decryptedStream.Length - 1)) As Byte
            decryptedStream.Read(result, 0, CInt(decryptedStream.Length))
            cryptStream.Close()
            Dim myutf As UTF8Encoding = New UTF8Encoding()
            Return myutf.GetString(result)
        End Function

        ''' <summary>
        ''' Gets a relatively random key sequence using a pseduorandom number plus some  
        ''' data in the Servers.xml file. 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>Servers.xml must have a comment as the second line of the file.</remarks>
        Private Shared Function GetKey() As Byte()
            Dim path As String = ConnectionManager.ServersXMLPath()

            Using sr As New StreamReader(path)
                sr.ReadLine()
                Dim cl As String = sr.ReadLine()

                If cl.StartsWith("<!-- ") = False OrElse cl.Length < 40 Then
                    Throw New ApplicationException("Invalid Servers.xml file.")
                Else
                    Dim rnd As New Random(3894671)
                    Dim bytes(IVMaxIndex) As Byte

                    For i As Integer = 0 To IVMaxIndex
                        Dim sum As Integer = 0
                        sum += Asc(cl.Substring(i, 1)) + rnd.Next(255)
                        bytes(i) = CByte(sum Mod 256)
                    Next

                    Return bytes
                End If
            End Using
        End Function
    End Class
End Class
