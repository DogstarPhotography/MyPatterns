''' <summary>
''' The client consumes the templatemethod
''' </summary>
''' <remarks></remarks>
Public Class Client
    ''' <summary>
    ''' In this case we will use both implementations from our client to show the differences in implementation
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub UseTemplateMethod()

        Dim UseOne As New ImplementationOne
        Dim UseTwo As New ImplementationTwo

        UseOne.TemplateMethod()
        UseTwo.TemplateMethod()
        UseTwo.FollowUp()

    End Sub
End Class
