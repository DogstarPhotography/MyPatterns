
Public Class Form1
    Private sqlReader As SQLInfoReader

#Region " Form Events "
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ConnectionManager.GetInstance.Initialise("Live", ConnectionManager.AccessMode.ReadAccess)
        BindEnvironmentCombo("Backup")
    End Sub

    Private Sub cbxEnvironment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxEnvironment.SelectedIndexChanged
        BindServerCombo()
    End Sub

    Private Sub cbxServer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxServer.SelectedIndexChanged
        GetTablesandViews()
    End Sub

    Private Sub lbxTables_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbxTables.SelectedIndexChanged
        txtCode.Text = GetVBCode(SelectedTable())
    End Sub

    Private Sub chkIncludeAttributes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeAttributes.CheckedChanged
        If txtCode.Text <> "" Then txtCode.Text = GetVBCode(SelectedTable())
    End Sub

    Private Sub chkReadOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReadOnly.CheckedChanged
        txtCode.Text = GetVBCode(SelectedTable())
    End Sub

    Private Sub chkIncludeNew_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeNew.CheckedChanged
        txtCode.Text = GetVBCode(SelectedTable())
    End Sub

    Private Sub txtCode_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.DoubleClick
        Clipboard.SetText(txtCode.Text)
        MessageBox.Show("VB Code has been copied to the clipboard", "Clipboard", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnShowReaderTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowReaderTemplate.Click
        If btnShowReaderTemplate.Text = "Reader Template" Then
            Dim vbCode As New System.Text.StringBuilder
            Dim tableName As String = SelectedTable()
            Dim sr As New System.IO.StreamReader("ReaderTemplate.txt")

            Do
                If sr.Peek <= 0 Then Exit Do
                Dim line = sr.ReadLine
                line = line.Replace("$", tableName)
                line = line.Replace("#", CStr(cbxServer.SelectedItem))
                vbCode.AppendLine(line)
            Loop

            txtCode.Text = vbCode.ToString
            btnShowReaderTemplate.Text = "Class Template"
        Else
            lbxTables_SelectedIndexChanged(Nothing, Nothing)
            btnShowReaderTemplate.Text = "Reader Template"
        End If
    End Sub

    Private Sub txtCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCode.TextChanged
        If txtCode.Text <> "" Then
            pnlOptions.Enabled = True
        Else
            pnlOptions.Enabled = False
        End If
    End Sub

    Private Sub btnSchema_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSchema.Click
        Dim schemaForm As New Schema
        schemaForm.Columns = sqlReader.GetTableColumns(SelectedTable())
        schemaForm.TableName = SelectedTable()
        schemaForm.ShowDialog()
        schemaForm.Close()
    End Sub
#End Region

#Region " Control Binding "
    Private Sub BindEnvironmentCombo(ByVal selectedEnvironment As String)
        Dim i As Integer = 0
        For Each env In ConnectionManager.GetInstance.GetEnvironmentNames()
            cbxEnvironment.Items.Add(env)
            If env = selectedEnvironment Then
                cbxEnvironment.SelectedIndex = i
            End If
            i += 1
        Next

        If cbxEnvironment.SelectedIndex = -1 Then cbxEnvironment.SelectedIndex = 0

        BindServerCombo()
    End Sub

    Private Sub BindServerCombo()
        ConnectionManager.GetInstance.Environment = CStr(cbxEnvironment.SelectedItem)
        Dim server = ConnectionManager.GetInstance.GetServersForEnvironment()
        cbxServer.DataSource = server
        cbxServer.SelectedIndex = 0
    End Sub

    Private Sub GetTablesandViews()
        Dim server = CStr(cbxServer.SelectedItem)
        sqlReader = New SQLInfoReader(server)
        lbxTables.Items.Clear()
        txtCode.Text = ""

        If ConnectionManager.GetInstance.TestConnection(server) = False Then
            MessageBox.Show("Unable to connect to database.")
        Else
            For Each table In sqlReader.GetTablesandViews
                lbxTables.Items.Add(If(table.Type = TableOrView.DatabaseType.Table, "T", "V") & " " & table.Name)
            Next
        End If
    End Sub

    Private Function SelectedTable() As String
        Return Split(CStr(lbxTables.SelectedItem), " ")(1)
    End Function
#End Region

#Region " Source code generation "
    Private Function GetVBCode(ByVal tableName As String) As String
        Dim columns = sqlReader.GetTableColumns(tableName)
        Dim vbCode As New System.Text.StringBuilder

        vbCode.AppendLine("Imports System.Data.Linq")
        vbCode.AppendLine("Imports System.Data.Linq.Mapping")
        vbCode.AppendLine("")

        If chkIncludeAttributes.Checked Then
            vbCode.AppendLine("<Table()> Public Class " & tableName)
        Else
            vbCode.AppendLine("Public Class " & tableName)
        End If

        For Each column In columns
            vbCode.AppendLine("    Private _" & column.ColumnName & " As " & column.VBColumnType)
        Next

        vbCode.AppendLine("")

        For Each column In columns
            Dim readOnlyString = ""
            If chkReadOnly.Checked Then readOnlyString = "ReadOnly "

            If chkIncludeAttributes.Checked Then
                vbCode.Append("    <Column()> Public ")
            Else
                vbCode.Append("    Public ")
            End If
            vbCode.AppendLine(readOnlyString & "Property " & column.ColumnNameFormatted & " As " & column.VBColumnType)
            vbCode.AppendLine("        Get")
            vbCode.AppendLine("            Return _" & column.ColumnName)
            vbCode.AppendLine("        End Get")

            If chkReadOnly.Checked = False Then
                vbCode.AppendLine("        Set(ByVal value As " & column.VBColumnType & ")")
                vbCode.AppendLine("            _" & column.ColumnName & " = value")
                vbCode.AppendLine("        End Set")
            End If

            vbCode.AppendLine("    End Property")
            vbCode.AppendLine("")
        Next

        If chkIncludeNew.Checked Then
            vbCode.AppendLine()
            vbCode.Append("    Public Sub New(")

            For i As Integer = 0 To columns.Count - 1
                If i <> 0 Then vbCode.Append(New String(" "c, 19))
                vbCode.Append("ByVal p" & columns(i).ColumnName & " As " & columns(i).VBColumnType)
                If i < columns.Count - 1 Then
                    vbCode.AppendLine(", _")
                End If
            Next
            vbCode.AppendLine(")")

            For Each column In columns
                vbCode.AppendLine("    _" & column.ColumnName & " = p" & column.ColumnName)
            Next

            vbCode.AppendLine("    End Sub")
        End If

        vbCode.AppendLine("End Class")

        Return vbCode.ToString
    End Function
#End Region
End Class
