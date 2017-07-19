''' <summary>
''' The strategy class is composed of two holders for algorithms that must be set before they can be used
''' </summary>
''' <remarks>Inheritors must assign specific implementations to Algorithm A and Algorithm B</remarks>
Public MustInherit Class Strategy
    'We need a holder for each algorithm
    Private MyAlgorithmA As IAlgorithmA
    Private MyAlgorithmB As IAlgorithmB
    ''' <summary>
    ''' Set which algorithmA to use
    ''' </summary>
    ''' <param name="NewAlgorithm">IAlgorithmA</param>
    ''' <remarks></remarks>
    Protected Sub SetAlgorithmA(ByVal NewAlgorithm As IAlgorithmA)
        MyAlgorithmA = NewAlgorithm
    End Sub
    ''' <summary>
    ''' Set which AlgorithmB to use
    ''' </summary>
    ''' <param name="NewAlgorithm">IAlgorithmB</param>
    ''' <remarks></remarks>
    Protected Sub SetAlgorithmB(ByVal NewAlgorithm As IAlgorithmB)
        MyAlgorithmB = NewAlgorithm
    End Sub
    ''' <summary>
    ''' Execute the algorithms
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RunMyAlgorithms()
        Try
            MyAlgorithmA.UseAlgorithm()
            MyAlgorithmB.UseAlgorithm()
        Catch ex As Exception
            Throw New Exception("You must set algorithms A and B before running")
        End Try
    End Sub
End Class
