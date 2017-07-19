Module Implementation

    Public Sub Main()

        'Create a model
        Dim NewModel As New Model

        'Pass the model to the controller
        Dim NewController As New Controller(NewModel)

        'Start
        NewController.Begin()

    End Sub

End Module
