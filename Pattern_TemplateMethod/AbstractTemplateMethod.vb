''' <summary>
''' This class implements a template method that calls various procedures in order
''' </summary>
''' <remarks></remarks>
Public MustInherit Class AbstractTemplateMethod
    ''' <summary>
    ''' Executes the operations in a specific order and cannot be overriden by inheritors
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub TemplateMethod()
        OperationOne() 'Cannot be used or overriden by inheritors
        OperationTwo() 'Must be overriden by inheritors
        HookThree() 'May be overriden by inheritors
        OperationFour() 'Can be used but not overriden by inheritors
    End Sub
    ''' <summary>
    ''' OperationOne cannot be used or overriden by inheritors so is always executed by this class
    ''' </summary>
    ''' <remarks>This procedure is private so it cannot be used or overriden</remarks>
    Private Sub OperationOne()
        'CODE_HERE
    End Sub
    ''' <summary>
    ''' OperationTwo must be implemented by inheritors
    ''' </summary>
    ''' <remarks>This procedure is protected so only inheritors can use it</remarks>
    Protected MustOverride Sub OperationTwo()
    ''' <summary>
    ''' HookThree may be overriden by inheritors but has no implementation in this class
    ''' </summary>
    ''' <remarks>HookThree returns immediately and is protected so only inheritors can use it</remarks>
    Protected Overridable Sub HookThree()
        Return
    End Sub
    ''' <summary>
    ''' OperationFour cannot be overriden by inheritors so is always executed by this class
    ''' </summary>
    ''' <remarks>OperationFour can be used by inheritors however</remarks>
    Protected Sub OperationFour()
        'CODE_HERE
    End Sub
End Class
