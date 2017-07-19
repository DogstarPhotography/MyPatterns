Imports System.Text
Imports System.IO
Imports System.Security.Cryptography
Imports System.Configuration

Namespace Asymmetric
    ''' <summary>
    ''' Common code for encryption tasks
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RSAEncryption

#Region "Asymmetric Encryption"

        Public Shared Function Encrypt(cleartext As String, publickeyxml As String) As String

            Dim clear As Byte() = ASCIIEncoding.UTF8.GetBytes(cleartext)
            Dim encrypted As Byte() = Encrypt(clear, publickeyxml)
            Dim encryptedtext As String = Convert.ToBase64String(encrypted)
            Return encryptedtext

        End Function

        Public Shared Function Decrypt(encryptedtext As String, privatekeyxml As String) As String

            Dim encrypted As Byte() = Convert.FromBase64String(encryptedtext)
            Dim decrypted As Byte() = Decrypt(encrypted, privatekeyxml)
            Dim plaintext As String = ASCIIEncoding.UTF8.GetString(decrypted)
            Return plaintext

        End Function

#End Region

#Region "RSA Encryption"

#Region "Create and Export RSA keys"
        ''' <summary>
        ''' Generates a new RSA key pair and exports it to xml without persisting it to a container
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' See https://msdn.microsoft.com/en-us/library/system.security.cryptography.rsa.toxmlstring(v=vs.110).aspx for important security items related to storing public/private keys in XML
        ''' </remarks>
        Public Shared Function NewRSAKeyXML() As String
            Static increment As Integer
            ' creates the CspParameters object and sets the key container name used to store the RSA key pair 
            Dim cp As New CspParameters()
            increment += 1
            cp.KeyContainerName = "TEMP_RSA_KEY_CONTAINER_" & increment.ToString
            cp.Flags = CspProviderFlags.UseDefaultKeyContainer
            ' instantiates the rsa instance accessing the key container MyKeyContainerName 
            Dim rsa As New RSACryptoServiceProvider(cp)
            'writes out the current key pair used in the rsa instance
            'Console.WriteLine("Key is : " & rsa.ToXmlString(True))
            Dim xml As String = rsa.ToXmlString(True)
            ' Delete the key entry in the container.
            rsa.PersistKeyInCsp = False
            ' Call Clear to release resources and delete the key from the container.
            rsa.Clear()
            Return xml
        End Function

        Public Shared Sub PersistRSAKeyXML(xml As String, ContainerName As String)
            ' Create the CspParameters object and set the key container 
            '  name used to store the RSA key pair.
            Dim cp As New CspParameters()
            cp.KeyContainerName = ContainerName
            cp.Flags = CspProviderFlags.UseDefaultKeyContainer

            ' Create a new instance of RSACryptoServiceProvider that accesses the key container.
            Dim rsa As New RSACryptoServiceProvider(cp)

            'Load from xml
            rsa.FromXmlString(xml)
            rsa.PersistKeyInCsp = True
            'rsa.ExportParameters(True)


            ' Display the key information to the console.
            'Console.WriteLine("Key retrieved from container : {0}", rsa.ToXmlString(True))
        End Sub

        Public Shared Function FetchRSAKeyPublicXML(ContainerName As String) As String
            Return FetchRSAKeyXML(ContainerName, False)
        End Function

        Public Shared Function FetchRSAKeyXML(ContainerName As String, includeprivate As Boolean) As String
            ' Create the CspParameters object and set the key container 
            '  name used to store the RSA key pair.
            Dim cp As New CspParameters()
            cp.KeyContainerName = ContainerName
            cp.Flags = CspProviderFlags.UseDefaultKeyContainer
            ' Create a new instance of RSACryptoServiceProvider that accesses the key container.
            Dim rsa As New RSACryptoServiceProvider(cp)
            'Fetch from container
            Return rsa.ToXmlString(includeprivate)
        End Function

        Public Shared Sub DeleteRSAKeyXML(ByVal ContainerName As String)
            ' Create a new instance of CspParameters.  Pass 
            ' 13 to specify a DSA container or 1 to specify 
            ' an RSA container.  The default is 1. 
            Dim cspParams As New CspParameters

            ' Specify the container name using the passed variable.
            cspParams.KeyContainerName = ContainerName
            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

            'Create a new instance of RSACryptoServiceProvider.  
            'Pass the CspParameters class to use the  
            'key in the container. 
            Dim rsa As New RSACryptoServiceProvider(cspParams)

            'Delete the key entry in the container.
            rsa.PersistKeyInCsp = False

            'Call Clear to release resources and delete the key from the container.
            rsa.Clear()

        End Sub

#End Region

#Region "RSA with XML Keys"

        Public Shared Function Encrypt(ByVal DataToEncrypt() As Byte, ByVal publickeyxml As String) As Byte()
            'Create a new instance of RSACryptoServiceProvider. 
            Dim rsa As New RSACryptoServiceProvider()
            rsa.FromXmlString(publickeyxml)
            rsa.PersistKeyInCsp = False
            'Encrypt the passed byte array and specify OAEP padding.   
            'OAEP padding is only available on Microsoft Windows XP or 
            'later.   
            Return rsa.Encrypt(DataToEncrypt, False)
        End Function

        Public Shared Function Decrypt(ByVal DataToDecrypt() As Byte, ByVal privatekeyxml As String) As Byte()
            'Create a new instance of RSACryptoServiceProvider. 
            Dim rsa As New RSACryptoServiceProvider()
            rsa.FromXmlString(privatekeyxml)
            rsa.PersistKeyInCsp = False
            'Decrypt the passed byte array and specify OAEP padding.   
            'OAEP padding is only available on Microsoft Windows XP or 
            'later.   
            Return rsa.Decrypt(DataToDecrypt, False)
        End Function

#End Region

#Region "https://msdn.microsoft.com/en-us/library/system.security.cryptography.rsacryptoserviceprovider(v=vs.110).aspx"

        'public Shared sub Main()
        '    Try
        '        'Create a UnicodeEncoder to convert between byte array and string. 
        '        Dim ByteConverter As New UnicodeEncoding()

        '        'Create byte arrays to hold original, encrypted, and decrypted data. 
        '        Dim dataToEncrypt As Byte() = ByteConverter.GetBytes("Data to Encrypt")
        '        Dim encryptedData() As Byte
        '        Dim decryptedData() As Byte

        '        'Create a new instance of RSACryptoServiceProvider to generate 
        '        'public and private key data. 
        '        Using RSA As New RSACryptoServiceProvider

        '            'Pass the data to ENCRYPT, the public key information  
        '            '(using RSACryptoServiceProvider.ExportParameters(false), 
        '            'and a boolean flag specifying no OAEP padding.
        '            encryptedData = RSAEncrypt(dataToEncrypt, RSA.ExportParameters(False), False)

        '            'Pass the data to DECRYPT, the private key information  
        '            '(using RSACryptoServiceProvider.ExportParameters(true), 
        '            'and a boolean flag specifying no OAEP padding.
        '            decryptedData = RSADecrypt(encryptedData, RSA.ExportParameters(True), False)

        '            'Display the decrypted plaintext to the console. 
        '            Console.WriteLine("Decrypted plaintext: {0}", ByteConverter.GetString(decryptedData))
        '        End Using
        '    Catch e As ArgumentNullException
        '        'Catch this exception in case the encryption did 
        '        'not succeed.
        '        Console.WriteLine("Encryption failed.")
        '    End Try
        'End Sub

        Public Shared Function Encrypt(ByVal DataToEncrypt() As Byte, ByVal RSAKeyInfo As RSAParameters, ByVal DoOAEPPadding As Boolean) As Byte()
            Try
                Dim encryptedData() As Byte
                'Create a new instance of RSACryptoServiceProvider. 
                Using RSA As New RSACryptoServiceProvider

                    'Import the RSA Key information. This only needs 
                    'to include the public key information.
                    RSA.ImportParameters(RSAKeyInfo)

                    'Encrypt the passed byte array and specify OAEP padding.   
                    'OAEP padding is only available on Microsoft Windows XP or 
                    'later.  
                    encryptedData = RSA.Encrypt(DataToEncrypt, DoOAEPPadding)
                End Using
                Return encryptedData
                'Catch and display a CryptographicException   
                'to the console. 
            Catch e As CryptographicException
                Console.WriteLine(e.Message)

                Return Nothing
            End Try
        End Function


        Public Shared Function Decrypt(ByVal DataToDecrypt() As Byte, ByVal RSAKeyInfo As RSAParameters, ByVal DoOAEPPadding As Boolean) As Byte()
            Try
                Dim decryptedData() As Byte
                'Create a new instance of RSACryptoServiceProvider. 
                Using RSA As New RSACryptoServiceProvider
                    'Import the RSA Key information. This needs 
                    'to include the private key information.
                    RSA.ImportParameters(RSAKeyInfo)

                    'Decrypt the passed byte array and specify OAEP padding.   
                    'OAEP padding is only available on Microsoft Windows XP or 
                    'later.  
                    decryptedData = RSA.Decrypt(DataToDecrypt, DoOAEPPadding)
                    'Catch and display a CryptographicException   
                    'to the console. 
                End Using
                Return decryptedData
            Catch e As CryptographicException
                Console.WriteLine(e.ToString())

                Return Nothing
            End Try
        End Function

#End Region

#Region "RSA with Keys in Containers"

        'See https://msdn.microsoft.com/en-us/library/ca5htw4f(v=vs.110).aspx

        'Sub Main()
        '    Try
        '        Dim KeyContainerName As String = "MyKeyContainer"

        '        'Create a new key and persist it in  
        '        'the key container.
        '        PersistRSAKeyContainer(KeyContainerName)

        '        'Create a UnicodeEncoder to convert between byte array and string. 
        '        Dim ByteConverter As New UnicodeEncoding

        '        'Create byte arrays to hold original, encrypted, and decrypted data. 
        '        Dim dataToEncrypt As Byte() = ByteConverter.GetBytes("Data to Encrypt")
        '        Dim encryptedData() As Byte
        '        Dim decryptedData() As Byte

        '        'Pass the data to ENCRYPT, the name of the key container,  
        '        'and a boolean flag specifying no OAEP padding.
        '        encryptedData = RSAEncrypt(dataToEncrypt, KeyContainerName, False)

        '        'Pass the data to DECRYPT, the name of the key container,  
        '        'and a boolean flag specifying no OAEP padding.
        '        decryptedData = RSADecrypt(encryptedData, KeyContainerName, False)

        '        'Display the decrypted plaintext to the console. 
        '        Console.WriteLine("Decrypted plaintext: {0}", ByteConverter.GetString(decryptedData))

        '        DeleteRSAKeyContainer(KeyContainerName)

        '    Catch e As ArgumentNullException
        '        'Catch this exception in case the encryption did 
        '        'not succeed.
        '        Console.WriteLine("Encryption failed.")
        '    End Try
        'End Sub

        Public Shared Sub PersistRSAKeyContainer(ByVal ContainerName As String)
            ' Create a new instance of CspParameters.  Pass 
            ' 13 to specify a DSA container or 1 to specify 
            ' an RSA container.  The default is 1. 
            Dim cspParams As New CspParameters

            ' Specify the container name using the passed variable.
            cspParams.KeyContainerName = ContainerName
            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

            'Create a new instance of RSACryptoServiceProvider to generate 
            'a new key pair.  Pass the CspParameters class to persist the  
            'key in the container. 
            Dim rsa As New RSACryptoServiceProvider(cspParams)
            rsa.PersistKeyInCsp = True

        End Sub

        Public Shared Sub DeleteRSAKeyContainer(ByVal ContainerName As String)
            ' Create a new instance of CspParameters.  Pass 
            ' 13 to specify a DSA container or 1 to specify 
            ' an RSA container.  The default is 1. 
            Dim cspParams As New CspParameters

            ' Specify the container name using the passed variable.
            cspParams.KeyContainerName = ContainerName
            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

            'Create a new instance of RSACryptoServiceProvider.  
            'Pass the CspParameters class to use the  
            'key in the container. 
            Dim rsa As New RSACryptoServiceProvider(cspParams)

            'Delete the key entry in the container.
            rsa.PersistKeyInCsp = False

            'Call Clear to release resources and delete the key from the container.
            rsa.Clear()

        End Sub

        Public Shared Function Encrypt(ByVal DataToEncrypt() As Byte, ByVal ContainerName As String, ByVal DoOAEPPadding As Boolean) As Byte()
            Try
                ' Create a new instance of CspParameters.  Pass 
                ' 13 to specify a DSA container or 1 to specify 
                ' an RSA container.  The default is 1. 
                Dim cspParams As New CspParameters

                ' Specify the container name using the passed variable.
                cspParams.KeyContainerName = ContainerName
                cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

                'Create a new instance of RSACryptoServiceProvider. 
                'Pass the CspParameters class to use the key  
                'from the key in the container. 
                Dim rsa As New RSACryptoServiceProvider(cspParams)

                'Encrypt the passed byte array and specify OAEP padding.   
                'OAEP padding is only available on Microsoft Windows XP or 
                'later.   
                Return rsa.Encrypt(DataToEncrypt, DoOAEPPadding)
                'Catch and display a CryptographicException   
                'to the console. 
            Catch e As CryptographicException
                Console.WriteLine(e.Message)

                Return Nothing
            End Try
        End Function

        Public Shared Function Decrypt(ByVal DataToDecrypt() As Byte, ByVal ContainerName As String, ByVal DoOAEPPadding As Boolean) As Byte()
            Try
                ' Create a new instance of CspParameters.  Pass 
                ' 13 to specify a DSA container or 1 to specify 
                ' an RSA container.  The default is 1. 
                Dim cspParams As New CspParameters

                ' Specify the container name using the passed variable.
                cspParams.KeyContainerName = ContainerName
                cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

                'Create a new instance of RSACryptoServiceProvider. 
                'Pass the CspParameters class to use the key  
                'from the key in the container. 
                Dim rsa As New RSACryptoServiceProvider(cspParams)

                'Decrypt the passed byte array and specify OAEP padding.   
                'OAEP padding is only available on Microsoft Windows XP or 
                'later.   
                Return rsa.Decrypt(DataToDecrypt, DoOAEPPadding)
                'Catch and display a CryptographicException   
                'to the console. 
            Catch e As CryptographicException
                Console.WriteLine(e.ToString())

                Return Nothing
            End Try
        End Function

#End Region

#Region "https://msdn.microsoft.com/en-us/library/tswxhw92(v=vs.85).aspx"


        '    'public Shared sub Main()
        '    '    Try
        '    '        ' Create a key and save it in a container.
        '    '        GenerateRSAKey("MyKeyContainer")

        '    '        ' Retrieve the key from the container.
        '    '        GetKeyFromContainer("MyKeyContainer")

        '    '        ' Delete the key from the container.
        '    '        DeleteKeyFromContainer("MyKeyContainer")

        '    '        ' Create a key and save it in a container.
        '    '        GenerateRSAKey("MyKeyContainer")

        '    '        ' Delete the key from the container.
        '    '        DeleteKeyFromContainer("MyKeyContainer")
        '    '    Catch e As CryptographicException
        '    '        Console.WriteLine(e.Message)
        '    '    End Try
        '    'End Sub

        '    Public Shared Sub GenerateRSAKey(ByVal ContainerName As String)
        '        ' Create the CspParameters object and set the key container 
        '        ' name used to store the RSA key pair.
        '        Dim cp As New CspParameters()
        '        cp.KeyContainerName = ContainerName
        '            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

        '        ' Create a new instance of RSACryptoServiceProvider that accesses
        '        ' the key container MyKeyContainerName.
        '        Dim rsa As New RSACryptoServiceProvider(cp)

        '        ' Display the key information to the console.
        '        Console.WriteLine("Key added to container:  {0}", rsa.ToXmlString(True))
        '    End Sub

        '    Public Shared Sub GetKeyFromContainer(ByVal ContainerName As String)
        '        ' Create the CspParameters object and set the key container 
        '        '  name used to store the RSA key pair.
        '        Dim cp As New CspParameters()
        '        cp.KeyContainerName = ContainerName
        '            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

        '        ' Create a new instance of RSACryptoServiceProvider that accesses
        '        ' the key container MyKeyContainerName.
        '        Dim rsa As New RSACryptoServiceProvider(cp)

        '        ' Display the key information to the console.
        '        Console.WriteLine("Key retrieved from container : {0}", rsa.ToXmlString(True))
        '    End Sub

        '    Public Shared Sub DeleteKeyFromContainer(ByVal ContainerName As String)
        '        ' Create the CspParameters object and set the key container 
        '        '  name used to store the RSA key pair.
        '        Dim cp As New CspParameters()
        '        cp.KeyContainerName = ContainerName
        '            cspParams.Flags = CspProviderFlags.UseDefaultKeyContainer

        '        ' Create a new instance of RSACryptoServiceProvider that accesses
        '        ' the key container.
        '        Dim rsa As New RSACryptoServiceProvider(cp)

        '        ' Delete the key entry in the container.
        '        rsa.PersistKeyInCsp = False

        '        ' Call Clear to release resources and delete the key from the container.
        '        rsa.Clear()

        '        Console.WriteLine("Key deleted.")
        '    End Sub

#End Region

#End Region

    End Class

End Namespace