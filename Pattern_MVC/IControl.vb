Public Interface IControl
    Inherits IObserver

    Sub Begin()

    Sub Finish()

    Sub Execute(ByVal TheInstruction As String)

End Interface
