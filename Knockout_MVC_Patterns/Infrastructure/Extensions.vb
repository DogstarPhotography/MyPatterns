Imports Newtonsoft.Json
Imports System.IO
Imports System.Runtime.CompilerServices


Public Module Extensions

    <Extension()> _
    Public Function ToJson(obj As Object) As String
        Dim js As JsonSerializer = JsonSerializer.Create(New JsonSerializerSettings)
        Dim jw As New StringWriter
        js.Serialize(jw, obj)
        Return jw.ToString
    End Function

End Module
