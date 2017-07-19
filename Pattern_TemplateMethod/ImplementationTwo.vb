''' <summary>
''' This class implements AbstractTemplateMethod by providing only the MustOverride portions of the template
''' </summary>
''' <remarks>This class also overrides HookThree</remarks>
Public Class ImplementationTwo
    Inherits AbstractTemplateMethod
    ''' <summary>
    ''' Only OperationTwo must be overriden when inheriting AbstractTemplateMethod
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub OperationTwo()
        'CODE_HERE
    End Sub
    ''' <summary>
    ''' HookThree MAY be overriden when inheriting AbstractTemplateMethod
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub HookThree()
        'We do not need to call HookThree in this case
        'MyBase.HookThree()
        'CODE_HERE
    End Sub
    ''' <summary>
    ''' FollowUp calls a method in AbstractTemplateMethod 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub FollowUp()
        MyBase.OperationFour()
    End Sub
End Class

