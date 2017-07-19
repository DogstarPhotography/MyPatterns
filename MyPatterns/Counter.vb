''' <summary>
''' Allows the consumer to count in an object oriented manner rather than manipulating a value
''' </summary>
''' <remarks>Default Start is 0 and default Step is 1</remarks>
Public Class Counter

    Private MyCounter As Integer '-2,147,483,648 through 2,147,483,647 (signed) 
    Private MyStep As UShort '0 to 65,535 (unsigned)
    Private MyStart As Integer '-2,147,483,648 through 2,147,483,647 (signed) 

    Public Sub New()
        MyStart = 0
        MyCounter = MyStart
        MyStep = 1
    End Sub

    Public Sub New(ByVal Start As Integer, ByVal [Step] As UShort)
        MyStart = Start
        MyCounter = MyStart
        MyStep = [Step]
    End Sub

    Public Property Index() As Integer
        Get
            Return MyCounter
        End Get
        Set(ByVal value As Integer)
            MyCounter = value
        End Set
    End Property

    Public Property [Step]() As UShort
        Get
            Return MyStep
        End Get
        Set(ByVal value As UShort)
            MyStep = value
        End Set
    End Property

    Public Sub Increment()
        MyCounter = MyCounter + MyStep
    End Sub

    Public Sub Decrement()
        MyCounter = MyCounter - MyStep
    End Sub

    Public Sub Reset()
        MyCounter = MyStart
    End Sub
End Class
''' <summary>
''' For counting up to really big numbers you can use a BigCounter
''' </summary>
''' <remarks>Default Start is 0 and default Step is 1</remarks>
Public Class BigCounter

    Private MyCounter As Long '-9,223,372,036,854,775,808 through 9,223,372,036,854,775,807 (signed)
    Private MyStep As UInteger '0 to 4,294,967,295 (unsigned)
    Private MyStart As Long '-9,223,372,036,854,775,808 through 9,223,372,036,854,775,807 (signed)

    Public Sub New()
        MyStart = 0
        MyCounter = MyStart
        MyStep = 1
    End Sub

    Public Sub New(ByVal Start As Long, ByVal [Step] As UInteger)
        MyStart = Start
        MyCounter = MyStart
        MyStep = [Step]
    End Sub

    Public Property Index() As Long
        Get
            Return MyCounter
        End Get
        Set(ByVal value As Long)
            MyCounter = value
        End Set
    End Property

    Public Property [Step]() As UInteger
        Get
            Return MyStep
        End Get
        Set(ByVal value As UInteger)
            MyStep = value
        End Set
    End Property

    Public Sub Increment()
        MyCounter = MyCounter + MyStep
    End Sub

    Public Sub Decrement()
        MyCounter = MyCounter - MyStep
    End Sub

    Public Sub Reset()
        MyCounter = MyStart
    End Sub
End Class