'Imports LateBindiing
'Imports Pattern_XMLSerialization

Public Class Tester

	'    Private WithEvents newController As New Pattern_Others.Controller

	'    Private Sub btnClientController_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClientController.Click

	'        newController.Test()

	'    End Sub

	'    Private Sub newController_Message(ByVal Content As String) Handles newController.Message

	'        InvoketxtMessages(txtMessages.Text & vbCrLf & Format(Now, "HH:mm:ss") & " " & Content)

	'    End Sub

	'#Region "Reflection"
	'    Private WithEvents MyBinder As Binder

	'    Private Sub btnReflection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReflection.Click

	'        MyBinder = New Binder
	'        MyBinder.Test("")

	'    End Sub

	'    Private Sub MyBinder_Message(ByVal Content As String) Handles MyBinder.Message

	'        InvoketxtReflection(txtReflection.Text & vbCrLf & Format(Now, "HH:mm:ss") & " " & Content)

	'    End Sub
	'#End Region

	'#Region "Thread Safe Control Updates"
	'    ' This method demonstrates a pattern for making thread-safe
	'    ' calls on a Windows Forms control. 
	'    '
	'    ' If the calling thread is different from the thread that
	'    ' created the TextBox control, this method creates a
	'    ' SetTextCallback and calls itself asynchronously using the
	'    ' Invoke method.
	'    ' or if the calling thread is the same as the thread that created
	'    ' the TextBox control, the Text property is set directly. 
	'    '
	'    ' This delegate enables asynchronous calls for setting
	'    ' the text property on a TextBox control.
	'    Delegate Sub InvokeTextboxCallback(ByVal NewValue As String)

	'    Private Sub InvoketxtMessages(ByVal NewValue As String)
	'        ' InvokeRequired required compares the thread ID of the
	'        ' calling thread to the thread ID of the creating thread.
	'        ' If these threads are different, it returns true.
	'        If Me.txtMessages.InvokeRequired Then
	'            Dim MyCallback As New InvokeTextboxCallback(AddressOf InvoketxtMessages)
	'            Invoke(MyCallback, New Object() {NewValue})
	'        Else
	'            txtMessages.Text = NewValue
	'        End If
	'    End Sub

	'    Private Sub InvoketxtReflection(ByVal NewValue As String)
	'        ' InvokeRequired required compares the thread ID of the
	'        ' calling thread to the thread ID of the creating thread.
	'        ' If these threads are different, it returns true.
	'        If Me.txtReflection.InvokeRequired Then
	'            Dim MyCallback As New InvokeTextboxCallback(AddressOf InvoketxtReflection)
	'            Invoke(MyCallback, New Object() {NewValue})
	'        Else
	'            txtReflection.Text = NewValue
	'        End If
	'    End Sub
	'#End Region

	'#Region "XMLSerialization"
	'    Private Sub btnSerialize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSerialize.Click

	'        Select Case cboXMLSelector.Text
	'            Case "SerializableComplexClass"
	'                Dim newXML As New SerializableComplexClass
	'                'XXX
	'                newXML.XXXsArray.Add(New XXX)
	'                newXML.XXXsArray.Add(New XXX)
	'                newXML.XXXsArray.Add(New XXX)
	'                Dim newXXXs As New XXXs
	'                newXXXs.XXX.Add(New XXX)
	'                newXXXs.XXX.Add(New XXX)
	'                newXXXs.XXX.Add(New XXX)
	'                newXML.XXXsElement = newXXXs
	'                'YYY
	'                newXML.YYYsArray.Add(New YYY)
	'                newXML.YYYsArray.Add(New YYY)
	'                newXML.YYYsArray.Add(New YYY)
	'                Dim newYYYs As New YYYs
	'                newYYYs.YYY.Add(New YYY)
	'                newYYYs.YYY.Add(New YYY)
	'                newYYYs.YYY.Add(New YYY)
	'                newXML.YYYsElement = newYYYs
	'                'ZZZ
	'                newXML.ZZZsArray.Add(New ZZZ)
	'                newXML.ZZZsArray.Add(New ZZZ)
	'                newXML.ZZZsArray.Add(New ZZZ)
	'                Dim newZZZs As New ZZZs
	'                newZZZs.ZZZ.Add(New ZZZ)
	'                newZZZs.ZZZ.Add(New ZZZ)
	'                newZZZs.ZZZ.Add(New ZZZ)
	'                newXML.ZZZsElement = newZZZs
	'                'Serialize
	'				txtSerialization.Text = Serialization.SerializeToText(Of SerializableComplexClass)(newXML)
	'            Case "XXXs"
	'                Dim newXML As New XXXs
	'                newXML.XXX.Add(New XXX)
	'                newXML.XXX.Add(New XXX)
	'                newXML.XXX.Add(New XXX)
	'                'Serialize
	'				txtSerialization.Text = Serialization.SerializeToText(Of XXXs)(newXML)
	'            Case "YYYs"
	'                Dim newXML As New YYYs
	'                newXML.YYY.Add(New YYY)
	'                newXML.YYY.Add(New YYY)
	'                newXML.YYY.Add(New YYY)
	'                'Serialize
	'				txtSerialization.Text = Serialization.SerializeToText(Of YYYs)(newXML)
	'            Case "ZZZs"
	'                Dim newXML As New ZZZs
	'                newXML.ZZZ.Add(New ZZZ)
	'                newXML.ZZZ.Add(New ZZZ)
	'                newXML.ZZZ.Add(New ZZZ)
	'                'Serialize
	'				txtSerialization.Text = Serialization.SerializeToText(Of ZZZs)(newXML)
	'            Case "XXX"
	'				txtSerialization.Text = Serialization.SerializeToText(Of XXX)(New XXX)
	'            Case "YYY"
	'				txtSerialization.Text = Serialization.SerializeToText(Of YYY)(New YYY)
	'            Case "ZZZ"
	'				txtSerialization.Text = Serialization.SerializeToText(Of ZZZ)(New ZZZ)
	'            Case Else
	'                'Ignore
	'        End Select

	'    End Sub

	'    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

	'    End Sub
	'#End Region

End Class
