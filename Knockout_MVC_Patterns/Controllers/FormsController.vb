Option Strict Off
Option Infer On

Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports System.Xml
Imports System.Web.Mvc

Public Class FormsController
    Inherits System.Web.Mvc.Controller

    Function Index() As ActionResult
        'Get the html content
        Dim formtext As String = FetchFileText("/Content/Forms/Form_01.html")
        Dim form As New FormModel
        form.Content = GetFormContent(formtext)
        'Get the menu
        Dim menutext As String = FetchFileText("/Content/Forms/Form_01_Menu.html")
        form.Menu = menutext
        'Prepare state by reading from xml file
        Dim dataxml As String = FetchFileText("/Content/Forms/Form_01_Data.xml")
        Dim statetext As String = BuildStateJS(dataxml)
        form.State = WrapScriptBlock(statetext)
        'Add script
        Dim scripttext As String = FetchFileText("/Content/Forms/Form_01_ViewModel.js")
        form.Script = WrapScriptBlock(scripttext)
        Return View(form)
    End Function

    ''' <summary>
    ''' We are using a wrapper model as a parameter rather than a string 
    ''' because the model binder will attempt to bind to Json data passed as a string, then fail, and return null.
    ''' </summary>
    <HttpPost>
    Function Save(model As JsonWrapperModel) As ActionResult
        Try
            If model IsNot Nothing AndAlso model.Json.Length > 0 Then
                'Debug.Print(model.Json.ToString)
                Dim dictjson = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(model.Json)
                Dim data As New FormData
                For Each kvp As KeyValuePair(Of String, Object) In dictjson
                    'Debug.Print(kvp.Key & ": " & kvp.Value.ToString)
                    Dim item As New DataItem With {.Name = kvp.Key, .Value = kvp.Value.ToString}
                    data.Data.Add(item)
                Next
                'Store data as xml
                'Debug.Print(data.Serialize)
                WriteFileText("/Content/Forms/Form_01_Data_Updated.xml", data.Serialize)
                Return New HttpStatusCodeResult(200)
            End If
        Catch
            'Ignore errors as we will return a 500 code if we get here
        End Try
        Return New HttpStatusCodeResult(500)
    End Function

#Region "Helpers"

    Protected Function FetchFileText(ByVal filepath As String) As String
        Dim serverfilepath As String = Server.MapPath(filepath)
        Return IO.File.ReadAllText(serverfilepath)
    End Function

    Protected Sub WriteFileText(filepath As String, text As String)
        Dim serverfilepath As String = Server.MapPath(filepath)
        IO.File.WriteAllText(serverfilepath, text)
    End Sub
    ''' <summary>
    ''' Create htmldoc and return everything within the formcontent element using http://htmlagilitypack.codeplex.com/
    ''' </summary>
    Protected Function GetFormContent(formtext As String) As String
        Dim hdoc As New HtmlDocument
        hdoc.LoadHtml(formtext)
        Dim body As HtmlNode = hdoc.DocumentNode.Descendants().Where(Function(n) n.Id = "formcontent").FirstOrDefault
        Return body.InnerHtml
    End Function

    Private Function WrapScriptBlock(scripttext As String) As String
        If scripttext.ToLower.Contains("</script>") Then Return scripttext
        Return "<script> " & vbCrLf & scripttext & vbCrLf & " </script>"
    End Function
    ''' <summary>
    ''' Build a Knockout state with which to initialise our Knockout ViewModel from the serialized FormData which is stored as xml for demonstration purposes
    ''' </summary>
    Private Function BuildStateJS(xml As String) As String
        Dim data As New FormData
        data.Deserialize(xml)
        Dim sb As New StringBuilder
        sb.AppendLine("var state = {")
        For Each item As DataItem In data.Data
            If IsNumeric(item.Value) Then
                sb.AppendLine(item.Name & ": " & item.Value & ",")
            ElseIf item.Value.ToLower = "false" OrElse item.Value.ToLower = "true" Then
                sb.AppendLine(item.Name & ": " & item.Value.ToLower & ",")
            Else
                sb.AppendLine(item.Name & ": '" & item.Value & "',")
            End If
        Next
        'Remove the last comma (not the last character as we have been appending lines)
        Dim index As Integer = sb.ToString.LastIndexOf(",")
        sb.Remove(index, 1)
        sb.AppendLine("};")
        Return sb.ToString
    End Function

#End Region

End Class
