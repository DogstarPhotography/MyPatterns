Public Class frmFetchWithCode

    Private conSample As SqlConnection
    Private cmdFetchSimple As SqlCommand
    Private rdrSimpleData As SqlDataReader

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Connect to DB
        conSample = New SqlConnection(My.Settings.SQLConnectionString)
        conSample.Open()

        'Create a command
        cmdFetchSimple = conSample.CreateCommand()
        cmdFetchSimple.CommandText = "SELECT * FROM Simple"

        'Display
        Label1.Text = "SQL Connection: " & conSample.State.ToString() & ", """ & conSample.ConnectionString & """"
        Label1.Text = Label1.Text & vbCrLf & vbCrLf
        Label1.Text = Label1.Text & "SQL Command: " & cmdFetchSimple.CommandText
        Label1.Text = Label1.Text & vbCrLf & vbCrLf

        'Fetch simple data
        rdrSimpleData = cmdFetchSimple.ExecuteReader()

        'Did we get data?
        If rdrSimpleData.HasRows = False Then
            'Display message
            Label1.Text = Label1.Text & "No data returned"
        Else
            'Read each row
            Dim row As Integer = 1
            Do While rdrSimpleData.Read = True
                'Read each column
                For col As Integer = 1 To rdrSimpleData.FieldCount
                    'Display data
                    If rdrSimpleData.IsDBNull(col - 1) Then
                        Label1.Text = Label1.Text & "R" & row & "C" & col & ": NULL"
                    Else
                        Label1.Text = Label1.Text & "R" & row & "C" & col & ": " & rdrSimpleData.GetValue(col - 1).ToString
                    End If
                    Label1.Text = Label1.Text & ", "
                Next
                row += 1
                Label1.Text = Label1.Text & vbCrLf
            Loop
        End If

        'Tidy up
        rdrSimpleData.Close()
        conSample.Close()

    End Sub
End Class