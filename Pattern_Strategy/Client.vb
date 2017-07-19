Public Class Client

    Public Sub TestStrategyPattern()

        Dim MyClient As Strategy

        MyClient = New StrategyAlpha
        MyClient.RunMyAlgorithms()

        MyClient = New StrategyBeta
        MyClient.RunMyAlgorithms()

    End Sub
End Class
