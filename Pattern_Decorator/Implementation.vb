Public Class Implementation

    Public Sub Test1()

        'Create a new item instance using the Component type so we can be polymorphic
        Dim TheItem As Component
        TheItem = New Item

        'Add a decoration
        TheItem = New Decoration_Named(TheItem)
        TheItem = New Decoration_Other(TheItem)
        'We can add more decorations here

        'Use the MyFunction method to get a value
        Dim FunctionValue As String = TheItem.MyFunction()

        'FunctionValue will now contain both values from the component AND the decoration(s)
        'Note that the value depends on the precise implementation details inside the decoration classes

    End Sub

    Public Sub Test2()

        'Create a new item instance using the Component type so we can be polymorphic
        Dim TheInstance As IComponent
        TheInstance = New Instance

        'Add a decoration - need to cast here as we have Option Strict On
        TheInstance = New Decoration_Instance(CType(TheInstance, Decorator))
        'We can add more decorations here

        'Use the MyFunction method to get a value
        Dim FunctionValue As String = TheInstance.MyFunction()

        'FunctionValue will now contain both values from the component AND the decoration(s)
        'Note that the value depends on the precise implementation details inside the decoration classes

    End Sub

End Class
