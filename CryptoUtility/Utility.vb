Imports Crypto.Symmetric
Imports Crypto
Public Class Utility
    Private Sub Utility_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Add methods to cboMethod
        cboMethod.Items.Add("AESEncryption")
        cboMethod.Items.Add("AESHMACEncryption")
        cboMethod.Items.Add("Hash PBKDF2")
        cboMethod.Items.Add("Hash SHA256")
        cboMethod.Items.Add("Hash SHA512")
        cboMethod.SelectedIndex = 0

    End Sub
    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        'Sanity check
        If txtPassword.Text.Length < EncryptionSettings.MinimumPasswordLength Then
            MsgBox("Password must be at least " & EncryptionSettings.MinimumPasswordLength & " characters long.")
            Exit Sub
        End If
        'Work
        Select Case cboMethod.Text
            Case "AESEncryption"
                If txtSalt.Text.Length < EncryptionSettings.MinimumPasswordLength Then
                    MsgBox("Salt must be at least " & EncryptionSettings.MinimumPasswordLength & " characters long for AES encryption.")
                    Exit Sub
                End If
                txtOutput.Text = AESEncryption.Encrypt(txtInput.Text, txtPassword.Text, txtSalt.Text)
            Case "AESHMACEncryption"
                txtOutput.Text = AESHMACEncryption.Encrypt(txtInput.Text, txtPassword.Text)
            Case "Hash PBKDF2"
                txtOutput.Text = Hash.PBKDF2(txtInput.Text, txtPassword.Text, CInt(Val(txtLength.Text)))
            Case "Hash SHA256"
                txtOutput.Text = Hash.SHA256(txtInput.Text)
            Case "Hash SHA512"
                txtOutput.Text = Hash.SHA512(txtInput.Text)
            Case Else
                'Do nothing
        End Select
        'Write to clipboard
        Clipboard.Clear()
        Clipboard.SetText(txtOutput.Text)
    End Sub

    Private Sub btnPasswordFill_Click(sender As Object, e As EventArgs) Handles btnPasswordFill.Click
        txtPassword.Text = Randomiser.Generate()
    End Sub

    Private Sub btnSaltFill_Click(sender As Object, e As EventArgs) Handles btnSaltFill.Click
        txtSalt.Text = Randomiser.Generate()
    End Sub

End Class
