''' <summary>
''' Settings used for symmetric encryption.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class EncryptionSettings

    'Preconfigured Encryption Parameters
    Public Shared ReadOnly BlockBitSize As Integer = 128
    Public Shared ReadOnly KeyBitSize As Integer = 256

    'Preconfigured Password Key Derivation Parameters
    Public Shared ReadOnly SaltBitSize As Integer = 64
    Public Shared ReadOnly Iterations As Integer = 10000
    Public Shared ReadOnly MinimumPasswordLength As Integer = 12

    'Calculated values
    Public Shared ReadOnly BlockBytes As Integer = BlockBitSize \ 8
    Public Shared ReadOnly KeyBytes As Integer = KeyBitSize \ 8
    Public Shared ReadOnly SaltBytes As Integer = SaltBitSize \ 8

End Class
