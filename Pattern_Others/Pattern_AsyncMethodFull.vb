
'Delegate for Method_V3Completed Event
Public Delegate Sub Method_V3CompletedEventHandler(ByVal sender As Object, ByVal e As Method_V3CompletedEventArgs)

'Delegate for ProgressChanged event
Public Delegate Sub Method_V3ProgressChangedEventHandler(ByVal e As Method_V3ProgressChangedEventArgs)

Public Class Pattern_AsyncMethodFull

    'Event to be raised when completed
    Public Event OnMethod_V3Completed As Method_V3CompletedEventHandler
    'Eventy to be raise as progress changes
    Public Event OnMethod_V3ProgressChanged As Method_V3ProgressChangedEventHandler

    'Declare a delegate to 'wrap' the worker method
    Private Delegate Sub WorkerEventHandler(ByVal AsyncOp As AsyncOperation, ByVal CompletionProcedureDelegate As SendOrPostCallback)

    'Instance of delegate
    Private WorkerDelegate As WorkerEventHandler
    'Threadpool delegates
    Private OnCompletedDelegate As SendOrPostCallback
    Private OnProgressReportDelegate As SendOrPostCallback
    'Dictionary to store tasks
    Private TaskTokens As New HybridDictionary()

    Public Sub New()
        'Create the delegate
        OnCompletedDelegate = New SendOrPostCallback(AddressOf OnMethod_V3AsyncCompleted)
        OnProgressReportDelegate = New SendOrPostCallback(AddressOf OnMethod_V3ReportProgress)
    End Sub

    'This function wll block
    Public Function Method_V3() As String
        'Do the work here
        'CODE_HERE
        Return ""
    End Function

    'This function will not block but the caller of this class must handle an event
    Public Overloads Sub Method_V3Async(ByVal UserState As Object, ByVal TaskID As Object)

        ' Create an AsyncOperation for taskId.
        Dim AsyncOp As AsyncOperation = AsyncOperationManager.CreateOperation(UserState)

        'Multiple threads will access the task dictionary,
        ' so it must be locked to serialize access.
        SyncLock TaskTokens.SyncRoot
            If TaskTokens.Contains(TaskID) Then
                Throw New ArgumentException("Task ID parameter must be unique", "taskId")
            End If
            'Add asyncop to dictionary
            TaskTokens(TaskID) = AsyncOp
        End SyncLock
        ' Start the asynchronous operation.
        WorkerDelegate = New WorkerEventHandler(AddressOf DoWork)
        WorkerDelegate.BeginInvoke(AsyncOp, OnCompletedDelegate, Nothing, Nothing)

    End Sub

    Private Sub DoWork(ByVal AsyncOp As AsyncOperation, ByVal MyCompletionProcedureDelegate As SendOrPostCallback)

        ' This method performs the actual computation.
        ' It is executed on the worker thread.
        Dim TheResult As Object = Nothing
        Dim CurrentResults As Object = Nothing
        Dim AnError As Exception = Nothing

        'Do the work here
        Dim ctrLoop As Integer
        For ctrLoop = 1 To 10
            'Get the current results
            CurrentResults = New Object
            'Create eventargs
            Dim ProgressEventArgs As New Method_V3ProgressChangedEventArgs( _
                ctrLoop * 10, CurrentResults, AsyncOp.UserSuppliedState)
            'Post the progress message
            AsyncOp.Post(OnProgressReportDelegate, ProgressEventArgs)
        Next
        TheResult = New Object

        'Create a state object and set it's values
        Dim MyState As New Method_V3State(TheResult, AnError, AsyncOp)

        'Create eventargs based on mystate
        Dim MyEventArgs As New Method_V3CompletedEventArgs(TheResult, AnError, False, MyState)

        ' In this case, don't allow cancellation, as the method is 
        ' about to raise the completed event.
        SyncLock TaskTokens.SyncRoot
            TaskTokens.Remove(AsyncOp.UserSuppliedState)
        End SyncLock

        'Post a completion message
        AsyncOp.PostOperationCompleted(OnCompletedDelegate, MyEventArgs)

        'Note that after the call to PostOperationCompleted, AsyncOp
        ' is no longer usable, and any attempt to use it will cause.
        ' an exception to be thrown.

    End Sub

    ' This method is invoked via the AsyncOperation object,
    ' so it is guaranteed to be executed on the correct thread.
    Protected Sub OnMethod_V3AsyncCompleted(ByVal UserState As Object) 'MS convention is that this sub must be protected and start with On

        Dim MyEventArgs As Method_V3CompletedEventArgs = CType(UserState, Method_V3CompletedEventArgs)
        RaiseEvent OnMethod_V3Completed(Me, MyEventArgs)

    End Sub

    ' This method cancels a pending asynchronous operation.
    Public Sub CancelAsync(ByVal TaskId As Object)

        'Lock the thread
        SyncLock TaskTokens.SyncRoot
            'Get object from dictionary entry identified by TaskID
            Dim objTemp As Object = TaskTokens(TaskId)
            'Did we get an object?
            If Not (objTemp Is Nothing) Then
                'Convert object to AsyncOp
                Dim AsyncOp As AsyncOperation = CType(objTemp, AsyncOperation)

                'Prepare some empty return values
                Dim TheResults As Object = Nothing
                Dim AnException As Exception = Nothing

                'Make sure we return canceled as true
                Dim canceled As Boolean = True

                'Roll all the bits into an eventargs object
                Dim NewEventArgs As New Method_V3CompletedEventArgs( _
                    TheResults, AnException, canceled, AsyncOp.UserSuppliedState)

                'The AsyncOp object is responsible for 
                ' marshalling the call to the proper 
                ' thread or context.
                AsyncOp.PostOperationCompleted(OnCompletedDelegate, NewEventArgs)

            End If
        End SyncLock

    End Sub

    ' This method is invoked via the AsyncOperation object,
    ' so it is guaranteed to be executed on the correct thread.
    Protected Sub OnMethod_V3ReportProgress(ByVal State As Object) 'MS convention is that this sub must be protected and start with On

        Dim MyEventArgs As Method_V3ProgressChangedEventArgs = CType(State, Method_V3ProgressChangedEventArgs)
        RaiseEvent OnMethod_V3ProgressChanged(MyEventArgs)

    End Sub
End Class

#Region "Supporting Classes for Pattern_AsyncMethodFull"
Public Class Method_V3CompletedEventArgs
    Inherits System.ComponentModel.AsyncCompletedEventArgs

    Private MyResults As Object

    Public Sub New(ByVal NewResults As Object, ByVal NewException As Exception, ByVal Canceled As Boolean, ByVal NewState As Object)
        MyBase.new(NewException, canceled, NewState)
        MyResults = NewResults
    End Sub

    Public ReadOnly Property Results() As Object
        Get
            Return MyResults
        End Get
    End Property
End Class

Public Class Method_V3ProgressChangedEventArgs
    Inherits System.ComponentModel.ProgressChangedEventArgs

    Private MyResults As Object

    Public Sub New(ByVal ProgressPercentage As Integer, ByVal NewResults As Object, ByVal NewState As Object)
        MyBase.new(ProgressPercentage, NewState)
        MyResults = NewResults
    End Sub

    Public ReadOnly Property Results() As Object
        Get
            Return MyResults
        End Get
    End Property
End Class

Friend Class Method_V3State
    Public Results As Object = Nothing
    Public AsyncOp As AsyncOperation = Nothing
    Public AnException As Exception = Nothing

    Public Sub New(ByVal NewResults As Object, ByVal NewException As Exception, ByVal NewAsyncOp As AsyncOperation)
        Me.Results = NewResults
        Me.AnException = NewException
        Me.AsyncOp = NewAsyncOp
    End Sub
End Class
#End Region
