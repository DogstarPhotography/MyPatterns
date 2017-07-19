''' <summary>
''' Class that inherits from factory
''' </summary>
''' <remarks></remarks>
Public Class FactoryTwo
    Inherits FactoryBase
    ''' <summary>
    ''' Class-specific implementation of CreateProduct
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides Function CreateProduct() As ProductBase
        'Return a specific product
        Return New ProductOne
    End Function
End Class
