'Instructions:
'   Find and replace Method_V1 with the name of the method
'   Complete CODE_HERE section
'   Change/replace Result class as required
'   Rename Pattern_AsyncMethod
''' <summary>
''' Delegate for Method_V1Completed Event
''' </summary>
''' <param name="sender"></param>
''' <param name="e"></param>
''' <remarks></remarks>
Public Delegate Sub Method_V1CompletedEventHandler(ByVal sender As Object, ByVal e As Method_V1CompletedEventArgs)

Public Class Pattern_AsyncMethod
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>This function wll block the consumer until it has finished</remarks>
    Public Function Method_V1() As Result
        'Do the work here
        'CODE_HERE
        Return New Result
    End Function

#Region "Async Method"
    'Event to be raised when completed
    Public Event OnMethod_V1Completed As Method_V1CompletedEventHandler
    'Declare a delegate to 'wrap' the worker method
    Private Delegate Sub WorkerEventHandler(ByVal AsyncOp As AsyncOperation, ByVal CompletionProcedureDelegate As SendOrPostCallback)
    'Instance of delegate
    Private WorkerDelegate As WorkerEventHandler
    'Threadpool delegate
    Private OnCompletedDelegate As SendOrPostCallback
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        'Create the delegate
        OnCompletedDelegate = New SendOrPostCallback(AddressOf OnMethod_V1AsyncCompleted)
    End Sub
    ''' <summary>
    ''' The *async version of Method_V1 function will not block but the consumer of this class must handle an event
    ''' </summary>
    ''' <param name="UserState"></param>
    ''' <remarks></remarks>
    Public Overloads Sub Method_V1Async(ByVal UserState As Object)
        ' Create an AsyncOperation
        Dim AsyncOp As AsyncOperation = AsyncOperationManager.CreateOperation(UserState)
        ' Start the asynchronous operation.
        WorkerDelegate = New WorkerEventHandler(AddressOf DoWork)
        WorkerDelegate.BeginInvoke(AsyncOp, OnCompletedDelegate, Nothing, Nothing)
    End Sub
    ''' <summary>
    ''' This method performs the actual asynchronous work, it is executed on the worker thread.
    ''' </summary>
    ''' <param name="AsyncOp"></param>
    ''' <param name="MyCompletionProcedureDelegate"></param>
    ''' <remarks></remarks>
    Private Sub DoWork(ByVal AsyncOp As AsyncOperation, ByVal MyCompletionProcedureDelegate As SendOrPostCallback)
        'Declare variables
        Dim TheResult As Result
        Dim AnError As Exception = Nothing
        'Do the work here by calling the Method_V1 blocking method
        TheResult = Method_V1()
        'Create a state object and set it's values
        Dim MyState As New Method_V1State(TheResult, AnError, AsyncOp)
        'Create eventargs based on state object
        Dim MyEventArgs As New Method_V1CompletedEventArgs(TheResult, AnError, MyState)
        'Post a completion message
        AsyncOp.PostOperationCompleted(OnCompletedDelegate, MyEventArgs)
        'Note that after the call to PostOperationCompleted, AsyncOp
        ' is no longer usable, and any attempt to use it will cause.
        ' an exception to be thrown.
    End Sub
    ''' <summary>
    ''' This method is invoked via the AsyncOperation object,
    ''' so it is guaranteed to be executed on the correct thread.
    ''' </summary>
    ''' <param name="UserState"></param>
    ''' <remarks></remarks>
    Protected Sub OnMethod_V1AsyncCompleted(ByVal UserState As Object) 'MS convention is that this sub must be protected and start with On
        'Turn the userstate into an eventargs object
        Dim MyEventArgs As Method_V1CompletedEventArgs = CType(UserState, Method_V1CompletedEventArgs)
        'Return the data to the consumer via an event
        RaiseEvent OnMethod_V1Completed(Me, MyEventArgs)
        'Async method is now finished
    End Sub
#End Region
End Class

#Region "Supporting Classes for Pattern_AsyncMethod"
''' <summary>
''' Simple class that contains data
''' </summary>
''' <remarks></remarks>
Public Class Result
    'Property
    Public Value As Object
End Class
''' <summary>
''' Class that inherits AsyncCompletedEventArgs and adds a Results property and a constructor
''' </summary>
''' <remarks></remarks>
Public Class Method_V1CompletedEventArgs
    Inherits System.ComponentModel.AsyncCompletedEventArgs

    Private MyResults As Result
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="NewResults">Result</param>
    ''' <param name="NewException">Exception</param>
    ''' <param name="NewState">Object</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal NewResults As Result, ByVal NewException As Exception, ByVal NewState As Object)
        MyBase.new(NewException, False, NewState)
        MyResults = NewResults
    End Sub
    ''' <summary>
    ''' Results property
    ''' </summary>
    ''' <value>Result</value>
    ''' <returns>Result</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Results() As Result
        Get
            Return MyResults
        End Get
    End Property
End Class
''' <summary>
''' Internal class for transferring the data from one thread to another
''' </summary>
''' <remarks></remarks>
Friend Class Method_V1State
    'Properties
    Public Results As Result = Nothing
    Public AsyncOp As AsyncOperation = Nothing
    Public AnException As Exception = Nothing
    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="NewResults">Result</param>
    ''' <param name="NewException">Exception</param>
    ''' <param name="NewAsyncOp">AsyncOperation</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal NewResults As Result, ByVal NewException As Exception, ByVal NewAsyncOp As AsyncOperation)
        Me.Results = NewResults
        Me.AnException = NewException
        Me.AsyncOp = NewAsyncOp
    End Sub
End Class
#End Region