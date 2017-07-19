''' <summary>
''' Inheritor of strategy
''' </summary>
''' <remarks></remarks>
Public Class StrategyBeta
    Inherits Strategy
    ''' <summary>
    ''' Use specific implementations of the algorithms
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.SetAlgorithmA(New AlgorithmAImplementationTwo)
        MyBase.SetAlgorithmB(New AlgorithmBImplementationOne)
    End Sub
End Class
