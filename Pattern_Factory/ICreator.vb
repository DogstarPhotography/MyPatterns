''' <summary>
''' Classes that implement ICreator create instances of classes that implement IProduct
''' </summary>
''' <remarks>We could replace this interface with a MustInherit class if we wanted to provide additional functionality to inheritors</remarks>
Public Interface ICreator
    ''' <summary>
    ''' Create an IProduct type object
    ''' </summary>
    ''' <returns>IProduct</returns>
    ''' <remarks>Inheritors must return an instance of a class that implements IProduct</remarks>
    Function CreateInstance() As IProduct
End Interface
