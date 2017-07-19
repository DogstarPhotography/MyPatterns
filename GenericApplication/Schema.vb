Imports System.Windows.Forms

Public Class Schema
    Private _Columns As List(Of TableColumn)
    Private _TableName As String

    Public Property Columns() As List(Of TableColumn)
        Get
            Return _Columns
        End Get
        Set(ByVal value As List(Of TableColumn))
            _Columns = value
        End Set
    End Property

    Public Property TableName() As String
        Get
            Return _TableName
        End Get
        Set(ByVal value As String)
            _TableName = value
            Me.Text = _TableName & " schema"
        End Set
    End Property

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub Schema_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dgvSchema.DataSource = _Columns.ToList
    End Sub
End Class
