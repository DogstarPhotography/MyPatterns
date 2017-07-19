''' <summary>
''' Abstract factory class
''' </summary>
''' <remarks></remarks>
Public MustInherit Class FactoryBase
    ''' <summary>
    ''' Return a concrete instance of a product
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>The product class is mustinherit as is this class so inheritors of this class can only return inheritors of product</remarks>
    Public MustOverride Function CreateProduct() As ProductBase

End Class
