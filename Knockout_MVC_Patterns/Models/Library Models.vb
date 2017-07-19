Public Class LibraryModel

    Private Property books As List(Of BookModel) = New List(Of BookModel)
    Private nextId As Integer = 1
    Property Name As String

    Public Sub New()
        Name = "My Home Library"
        AddBook(New BookModel With {.title = "Oliver Twist", .Author = "Charles Dickens", .Year = 1837})
        AddBook(New BookModel With {.Title = "Winnie-the-Pooh", .Author = "A. A. Milne", .Year = 1926})
        AddBook(New BookModel With {.Title = "The Hobbit", .Author = "J. R. R. Tolkien", .Year = 1937})
        AddBook(New BookModel With {.Title = "The Bicentennial Man", .Author = "Isaac Asimov", .Year = 1976})
        AddBook(New BookModel With {.Title = "The Green Mile", .Author = "Stephen King", .Year = 1996})
    End Sub

    Public Function GetBooks() As IEnumerable(Of BookModel)
        Return books
    End Function

    Function GetBook(id As Integer) As Object
        Return books.Find(Function(b) b.Id = id)
    End Function

    Public Sub AddBook(book As BookModel)
        book.Id = nextId
        nextId += 1
        books.Add(book)
    End Sub

    Public Function UpdateBook(book As BookModel) As Boolean
        Dim index As Integer = books.FindIndex(Function(b) b.Id = book.Id)
        If index = -1 Then Return False
        books.RemoveAt(index)
        books.Insert(index, book)
        Return True
    End Function

    Public Sub RemoveBook(id As Integer)
        books.RemoveAll(Function(b) b.Id = id)
    End Sub

End Class

Public Class BookModel

    Property Id As Integer
    Property Title As String
    Property Author As String
    Property Year As Integer

End Class
