Imports System.Runtime.CompilerServices
Imports System.IO

''' <summary>
''' Demonstrates functions that return html from a model and that can be used within a View
''' </summary>
''' <remarks>Any function here must be wrapped like this to prevent the view from encoding the markup as text: @Html.Raw([function]) </remarks>
Public Class HtmlGeneration
    ''' <summary>
    ''' Creates a table to display the item by using an HtmlTextWriter to create the Html.
    ''' </summary>
    ''' <param name="data">DynamicScriptGenerationData List</param>
    ''' <returns>String</returns>
    ''' <remarks>Obviously, for all but the most minor of purposes, this requires a lot more code than using a partial view 
    ''' so we must consider whether or not this technique is optimal for a particular problem.</remarks>
    Public Shared Function DisplayDynamicScriptGenerationDataTable(data As List(Of DynamicScriptGenerationData)) As String
        Dim stringwriter As New StringWriter
        Using writer As New HtmlTextWriter(stringwriter)
            'Attributes have to be added before calling RenderBeginTag
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "table table-striped")
            writer.RenderBeginTag(HtmlTextWriterTag.Table)
            For Each item As DynamicScriptGenerationData In data
                writer.RenderBeginTag(HtmlTextWriterTag.Tr)
                writer.RenderBeginTag(HtmlTextWriterTag.Td)
                writer.Write(item.ID)
                writer.RenderEndTag() '/td
                writer.RenderBeginTag(HtmlTextWriterTag.Td)
                writer.Write(item.Name)
                writer.RenderEndTag() '/td
                writer.RenderBeginTag(HtmlTextWriterTag.Td)
                writer.Write(item.Description)
                writer.RenderEndTag() '/td
                'Add a form
                writer.RenderBeginTag(HtmlTextWriterTag.Td)
                writer.AddAttribute("action", "/")
                writer.AddAttribute("method", "post")
                writer.RenderBeginTag(HtmlTextWriterTag.Form)
                writer.AddAttribute(HtmlTextWriterAttribute.Type, "Submit")
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "btn btn-primary")
                writer.AddAttribute(HtmlTextWriterAttribute.Id, "btnSubmit_" & item.ID)
                writer.RenderBeginTag(HtmlTextWriterTag.Button)
                writer.Write("Submit")
                writer.RenderEndTag() '/button
                writer.RenderEndTag() '/form
                writer.RenderEndTag() '/td
                writer.RenderEndTag() '/tr
            Next
            writer.RenderEndTag() '/table
        End Using
        Return stringwriter.ToString
    End Function

End Class
