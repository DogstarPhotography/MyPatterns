''' <summary>
''' This class implements AbstractTemplateMethod by providing only the MustOverride portions of the template
''' </summary>
''' <remarks></remarks>
Public Class ImplementationOne
    Inherits AbstractTemplateMethod
    ''' <summary>
    ''' Only OperationTwo must be overriden when inheriting AbstractTemplateMethod
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OperationTwo()
        'CODE_HERE
    End Sub
End Class
