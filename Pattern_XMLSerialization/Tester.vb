Imports System.Xml.Serialization
Imports System.IO

Public Class Tester

    ''' <summary>
    ''' Create some test data, serialize it to a file, then read the file and display the data
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnSerialize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSerialize.Click

        'Get sample data
        Dim iniFile As SectionFile = CreateSectionFile()

        'Serialize to file
        Dim newWriter As TextWriter = New StreamWriter("TEST.XML")
        Dim newSerializer As XmlSerializer = New XmlSerializer(GetType(SectionFile))
        newSerializer.Serialize(newWriter, iniFile)
        newWriter.Close()

        'Read and display
        Dim newReader As TextReader = New StreamReader("TEST.XML")
        txtXML.Text = newReader.ReadToEnd()

    End Sub

    Private Function CreateSectionFile() As SectionFile

        Dim newSectionFile As SectionFile = New SectionFile
        Dim ctrSection As Integer
        Dim ctrData As Integer

        'Create some data
        newSectionFile.Name = "TEST"
        newSectionFile.Description = "An INI-like file"
        For ctrSection = 1 To 3
            Dim newSection As Section = New Section
            newSection.Name = "SECTION_" & ctrSection.ToString
            newSection.Description = "This is section " & ctrSection.ToString
            For ctrData = 1 To 3
                Dim newNameValue As NameValue = New NameValue
                newNameValue.Name = "NAME_" & ctrData.ToString
                newNameValue.Value = "VALUE_" & ctrData.ToString
                newSection.NameValuePairs.Add(newNameValue)
            Next
            newSectionFile.Sections.Add(newSection)
        Next

        Return newSectionFile

    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        If dlgSaveFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim newWriter As TextWriter = New StreamWriter(dlgSaveFile.FileName)
            newWriter.Write(txtXML.Text)
            newWriter.Close()
        End If

    End Sub
End Class
