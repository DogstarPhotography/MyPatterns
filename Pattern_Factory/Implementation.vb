Public Class Implementation

    Public Sub Test()

        'Create an instance of a creator
        Dim MyCreator As ICreator
        'We can implement different creators by choosing which class we use here
        'All creators however must implement ICreator
        MyCreator = New CreatorInstance

        'Create an instance of a product
        Dim MyProduct As IProduct
        'The creator used above returns an instance of a product
        MyProduct = MyCreator.CreateInstance

    End Sub

End Class
