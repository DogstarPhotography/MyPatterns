Imports System.Xml

Module Library_XML

    ' Return a formatted XML string for this document.
    Public Function XmlDocToString(ByVal xml_doc As XmlDocument) As String
        Dim string_writer As New IO.StringWriter()
        Dim xml_text_writer As New XmlTextWriter(string_writer)

        xml_text_writer.Formatting = Formatting.Indented
        xml_doc.WriteTo(xml_text_writer)
        xml_text_writer.Flush()

        Return string_writer.ToString()
    End Function

    ' Add an XML declaration to the document.
    Public Function AddXmlDeclaration( _
      ByVal XmlDoc As XmlDocument, _
      Optional ByVal version As String = "1.0", _
      Optional ByVal encoding As String = Nothing, _
      Optional ByVal standalone As String = Nothing) _
      As XmlDeclaration

        ' Make the declaration.
        Dim xml_declaration As XmlDeclaration
        xml_declaration = XmlDoc.CreateXmlDeclaration( _
            version, encoding, standalone)

        ' Prepend the child so we know it's in the right place.
        XmlDoc.PrependChild(xml_declaration)
        Return xml_declaration
    End Function

    ' Create an XmlProcessingInstruction.
    Public Function AddXmlProcessingInstruction( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal target As String, _
      ByVal data As String) _
      As XmlProcessingInstruction

        ' Make the processing instruction.
        Dim xml_processing_instruction As XmlProcessingInstruction
        xml_processing_instruction = _
            XmlDoc.CreateProcessingInstruction(target, data)
        parent_node.AppendChild(xml_processing_instruction)
        Return xml_processing_instruction
    End Function

    ' Associate the XML file with an XSL file.
    Public Function AddXmlXSLFile( _
      ByVal XmlDoc As XmlDocument, _
      ByVal xsl_file As String) _
      As XmlProcessingInstruction

        Return AddXmlProcessingInstruction( _
            XmlDoc, XmlDoc, _
            "xml-stylesheet", _
            "type=""text/xsl"" href=""" & xsl_file & """")
    End Function

    ' Create a DOCTYPE section.
    Public Function AddXmlDocumentType( _
      ByVal XmlDoc As XmlDocument, _
      ByVal document_name As String, _
      ByVal internal_subset As String, _
      Optional ByVal public_id As String = Nothing, _
      Optional ByVal system_id As String = Nothing) _
      As XmlDocumentType

        ' Make the DOCTYPE section.
        Dim XmlDoc_type As XmlDocumentType
        XmlDoc_type = XmlDoc.CreateDocumentType( _
             document_name, public_id, system_id, _
             internal_subset)
        XmlDoc.AppendChild(XmlDoc_type)
        Return XmlDoc_type
    End Function

    ' Add a new XmlElement to the node.
    Public Function AddXmlElement( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal element_name As String) _
      As XmlElement

        ' Make the new element.
        Dim xml_element As XmlElement
        xml_element = XmlDoc.CreateElement(element_name)
        parent_node.AppendChild(xml_element)
        Return xml_element
    End Function

    ' Add an XmlAttribute to the node.
    Public Function AddXmlAttribute( _
      ByVal XmlDoc As XmlDocument, _
      ByVal xml_node As XmlNode, _
      ByVal attribute_name As String, _
      ByVal attribute_value As String) _
      As XmlAttribute

        ' Make the new attribute.
        Dim xml_attribute As XmlAttribute
        xml_attribute = XmlDoc.CreateAttribute(attribute_name)
        xml_attribute.Value = attribute_value
        xml_node.Attributes.Append(xml_attribute)
        Return xml_attribute
    End Function

    ' Add an XmlWhitespace object containing a vbCrLf and the
    ' indicated number of spaces to the parent node.
    Public Function AddXmlNewLine( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal num_spaces As Integer) _
      As XmlWhitespace

        ' Make the XmlWhitespace object.
        Dim xml_whitespace As XmlWhitespace
        xml_whitespace = XmlDoc.CreateWhitespace( _
            vbCrLf & Space(num_spaces))

        ' Add it to the parent node.
        parent_node.AppendChild(xml_whitespace)
        Return xml_whitespace
    End Function

    ' Add a comment to the node.
    Public Function AddXmlComment( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal comment_value As String) _
      As XmlComment

        ' Make the new comment.
        Dim xml_comment As XmlComment
        xml_comment = XmlDoc.CreateComment(" " & comment_value & " ")
        parent_node.AppendChild(xml_comment)
        Return xml_comment
    End Function

    ' Add some CDATA to the node.
    Public Function AddXmlCdataSection( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal cdata_text As String) _
      As XmlCDataSection

        ' Make the new CDATA section.
        Dim xml_cdata_section As XmlCDataSection
        xml_cdata_section = XmlDoc.CreateCDataSection(cdata_text)
        parent_node.AppendChild(xml_cdata_section)
        Return xml_cdata_section
    End Function

    ' Add a new XmlText node to the node.
    Public Function AddXmlText( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal element_text As String) _
      As XmlText

        ' Make the new XmlText object.
        Dim xml_text As XmlText
        xml_text = XmlDoc.CreateTextNode(element_text)
        parent_node.AppendChild(xml_text)
        Return xml_text
    End Function

    ' Add an entitiy reference to the node.
    Public Function AddXmlEntityReference( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal entity_name As String) _
      As XmlEntityReference

        ' Make the new entity reference.
        Dim xml_entity_reference As XmlEntityReference
        xml_entity_reference = XmlDoc.CreateEntityReference(entity_name)
        parent_node.AppendChild(xml_entity_reference)
        Return xml_entity_reference
    End Function

    ' Add an XmlWhitespace object.
    Public Function AddXmlWhitespace( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal whitespace As String) _
      As XmlWhitespace

        ' Make the XmlWhitespace object.
        Dim xml_whitespace As XmlWhitespace
        xml_whitespace = XmlDoc.CreateWhitespace(whitespace)

        ' Add it to the parent node.
        parent_node.AppendChild(xml_whitespace)
        Return xml_whitespace
    End Function

    ' Add an XmlSignificantWhitespace object.
    Public Function AddXmlSignificantWhitespace( _
      ByVal XmlDoc As XmlDocument, _
      ByVal parent_node As XmlNode, _
      ByVal whitespace As String) _
      As XmlSignificantWhitespace

        ' Make the XmlSignificantWhitespace object.
        Dim xml_whitespace As XmlSignificantWhitespace
        xml_whitespace = XmlDoc.CreateSignificantWhitespace(whitespace)

        ' Add it to the parent node.
        parent_node.AppendChild(xml_whitespace)
        Return xml_whitespace
    End Function

End Module
