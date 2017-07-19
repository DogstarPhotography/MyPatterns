''' <summary>
''' Concrete instance of a class that implements ICreator
''' </summary>
''' <remarks>Implements ICreator</remarks>
Public Class CreatorInstance
    Implements ICreator
    ''' <summary>
    ''' Returns a concrete instance of ProductInstance class
    ''' </summary>
    ''' <returns>ProductInstance (Implements IProduct)</returns>
    ''' <remarks></remarks>
    Public Function CreateInstance() As IProduct Implements ICreator.CreateInstance
        Return New ProductInstance
    End Function
End Class
