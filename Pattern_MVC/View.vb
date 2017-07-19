''' <summary>
''' The view implements an observer interface so it can observe the model
''' </summary>
''' <remarks></remarks>
Public Class View
    Implements IObserver

#Region "IObserver Implementation"
    Private MyModel As IModel
    ''' <summary>
    ''' Default constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    ''' <summary>
    ''' Register this observer with the given Model
    ''' </summary>
    ''' <param name="TheModel">IModel</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TheModel As IModel)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        MyModel = TheModel
        MyModel.RegisterObserver(Me)
    End Sub
    ''' <summary>
    ''' Do something with the update data
    ''' </summary>
    ''' <param name="TheData">String</param>
    ''' <remarks></remarks>
    Public Sub UpdateObserver(ByVal TheData As String) Implements IObserver.UpdateObserver

    End Sub
    ''' <summary>
    ''' Finished so remove reference to Model
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub Finalize()
        MyModel.RemoveObserver(Me)
        MyBase.Finalize()
    End Sub
#End Region

#Region "Public MVC Methods"
    Private MyController As IControl

    Public Sub CreateView()

    End Sub

    Public Sub UpdateView(ByVal TheData As String)

    End Sub

    Public Sub ViewEvent(ByVal TheInstruction As String)

        MyController.Execute(TheInstruction)

    End Sub
#End Region
End Class
