Public Class Implementation

    Public Sub Test()

        'Create two variables for referencing Factory and Product
        Dim MyFactory As FactoryBase
        Dim TheProduct As ProductBase

        'Create concrete instances from factory one
        MyFactory = New FactoryOne
        TheProduct = MyFactory.CreateProduct

        'Run a method
        TheProduct.MySub()

        'Create concrete instances from factory two
        MyFactory = New FactoryTwo
        TheProduct = MyFactory.CreateProduct

        'Run a method
        TheProduct.MySub()


    End Sub

End Class
