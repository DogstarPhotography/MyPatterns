''' <summary>
''' Classes that implement IProduct can be created by classes that implement ICreator
''' </summary>
''' <remarks>We could replace this interface with a MustInherit class if we wanted to provide additional functionality to inheritors</remarks>
Public Interface IProduct
    ''' <summary>
    ''' Example method
    ''' </summary>
    ''' <remarks></remarks>
    Sub MySub()
End Interface
