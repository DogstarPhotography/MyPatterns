''' <summary>
''' Inheritor of strategy
''' </summary>
''' <remarks></remarks>
Public Class StrategyAlpha
    Inherits Strategy
    ''' <summary>
    ''' Use specific implementations of the algorithms
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.SetAlgorithmA(New AlgorithmAImplementationOne)
        MyBase.SetAlgorithmB(New AlgorithmBImplementationOne)
    End Sub
End Class
