Imports System.Text

Public Class HeirarchyExampleTester

    Public Function TestMethod() As String
        Dim sb As New StringBuilder
        Dim catalog As New Catalog("..\..\exampleXml.xml")
        For Each artist As Artist In catalog.Artists
            sb.AppendLine(artist.Name)
            For Each song As Song In artist.Song
                sb.AppendLine("   " & song.Text)
            Next
        Next
        Return sb.ToString
    End Function

End Class
