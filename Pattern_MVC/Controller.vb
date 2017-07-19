''' <summary>
''' The controller implements an observer interface so it can observe the model
''' </summary>
''' <remarks></remarks>
Public Class Controller
    Implements IControl

#Region "IObserver Implementation"
    Private MyModel As IModel
    '''' <summary>
    '''' Register this observer with the given Model
    '''' </summary>
    '''' <param name="TheModel">IModel</param>
    '''' <remarks></remarks>
    'Public Sub New(ByVal TheModel As IObservable)
    '    MyModel = TheModel
    '    MyModel.RegisterObserver(Me)
    'End Sub

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

#Region "IControl Implementation"
    Private MyView As View

    Public Sub New(ByVal TheModel As Model)
        'Grab the model
        MyModel = TheModel
        MyModel.RegisterObserver(Me)
        'Create the view
        MyView = New View(MyModel)
        MyView.CreateView()
    End Sub

    Public Sub Begin() Implements IControl.Begin

        MyView.Show()

    End Sub

    Public Sub Finish() Implements IControl.Finish

        MyView.Close()

    End Sub

    Public Sub Execute(ByVal TheInstruction As String) Implements IControl.Execute

        'Convert TheInstruction to TheData
        Dim TheData As String
        TheData = TheInstruction.ToUpper

        'Update the model
        MyModel.UpdateState(TheData)

    End Sub

#End Region

End Class
