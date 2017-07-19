Imports System.ServiceModel
Imports System.Diagnostics.CodeAnalysis

''' <summary>
''' WCF proxies do not clean up properly if they throw an exception. This class contains methods which ensure that the service proxy is handled correctly.
''' Note that you must not call TService.Close() or TService.Abort() within the 'action' lambda sub/func.
''' <code>
''' 'Original code with Using block
''' Using service As New WCF_MessageService.MessageClient
'''     For Each msg As Message In Messages
'''         If service.SendMessage(msg) = True Then
'''             Me.Messages.Items.Remove(msg)
'''         End If
'''     Next
''' End Using
''' 'Replacement code with a lambda function and an inline procedure
''' ServiceHelper.Using(Of WCF_MessageService.MessageClient)(
'''     Sub(service)
'''         For Each msg As Message In Messages
'''             If service.Action(msg) = True Then
'''                 Me.Messages.Items.Remove(msg)
'''             End If
'''         Next
'''     End Sub)
''' </code>
''' The 'payload block' of code that does the actual work is, in the original version, wrapped with a Using block. 
''' The Microsoft recommendation wraps it in a Try...Catch block and the ServiceHelper wraps it in a lambda function that is called by the Using method(s).
''' </summary>
''' <remarks>
''' Original idea: http://nimtug.org/blogs/damien-mcgivern/archive/2009/05/26/wcf-communicationobjectfaultedexception-quot-cannot-be-used-for-communication-because-it-is-in-the-faulted-state-quot-messagesecurityexception-quot-an-error-occurred-when-verifying-security-for-the-message-quot.aspx
''' </remarks>
<ExcludeFromCodeCoverage> Public NotInheritable Class ServiceHelper
    ''' <summary>
    ''' Using method that takes a generic service type and a lambda sub.
    ''' </summary>
    ''' <typeparam name="ServiceT">The type of the service to use</typeparam>
    ''' <param name="action">Lambda of the Sub to perform with the service</param>
    Public Shared Sub [Using](Of ServiceT As {ICommunicationObject, IDisposable, New})(action As Action(Of ServiceT))
        Dim service = New ServiceT()
        Dim serviceclosed As Boolean = False
        Try
            action(service)
            If service.State <> CommunicationState.Faulted Then
                service.Close()
                serviceclosed = True
            End If
        Finally
            If serviceclosed = False Then
                service.Abort()
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Using method that takes a generic service type and a lambda sub which includes a parameter.
    ''' </summary>
    ''' <typeparam name="ServiceT">The type of the service to use.</typeparam>
    ''' <typeparam name="TParameter">The type of the input parameter.</typeparam>
    ''' <param name="action">Lambda of the Sub to perform with the service.</param>
    ''' <param name="input">The input value.</param>
    Public Shared Sub [Using](Of ServiceT As {ICommunicationObject, IDisposable, New}, TParameter)(input As TParameter, action As Action(Of ServiceT, TParameter))
        Dim service = New ServiceT()
        Dim serviceclosed As Boolean = False
        Try
            action(service, input)
            If service.State <> CommunicationState.Faulted Then
                service.Close()
                serviceclosed = True
            End If
        Finally
            If serviceclosed = False Then
                service.Abort()
            End If
        End Try
    End Sub
    ''' <summary>
    ''' Using method that takes a generic service type and a lambda function with a return value.
    ''' </summary>
    ''' <typeparam name="ServiceT">The type of the service to use.</typeparam>
    ''' <typeparam name="TResult">The type of the return value.</typeparam>
    ''' <param name="action">Lambda of the Function to perform with the service.</param>
    ''' <returns>The return value.</returns>
    Public Shared Function [Using](Of ServiceT As {ICommunicationObject, IDisposable, New}, TResult)(action As Func(Of ServiceT, TResult)) As TResult
        Dim service = New ServiceT()
        Dim serviceclosed As Boolean = False
        Dim result As TResult
        Try
            result = action(service)
            If service.State <> CommunicationState.Faulted Then
                service.Close()
                serviceclosed = True
            End If
        Finally
            If serviceclosed = False Then
                service.Abort()
            End If
        End Try
        Return result
    End Function
    ''' <summary>
    ''' Using method that takes a generic service type and a lambda function with an input parameter and a return value.
    ''' </summary>
    ''' <typeparam name="ServiceT">The type of the service to use.</typeparam>
    ''' <typeparam name="TParameter">The type of the input parameter.</typeparam>
    ''' <typeparam name="TResult">The type of the return value.</typeparam>
    ''' <param name="action">Lambda of the Function to perform with the service.</param>
    ''' <param name="input">The input value.</param>
    ''' <returns>The return value.</returns>
    Public Shared Function [Using](Of ServiceT As {ICommunicationObject, IDisposable, New}, TParameter, TResult)(input As TParameter, action As Func(Of ServiceT, TParameter, TResult)) As TResult
        Dim service = New ServiceT()
        Dim serviceclosed As Boolean = False
        Dim result As TResult
        Try
            result = action(service, input)
            If service.State <> CommunicationState.Faulted Then
                service.Close()
                serviceclosed = True
            End If
        Finally
            If serviceclosed = False Then
                service.Abort()
            End If
        End Try
        Return result
    End Function

End Class

#Region "Example Usage"

'Namespace WCF_ServiceReference

'#Region "Support classes - collapse this"
'    ''' <summary>
'    ''' This is a stand in for the class that would be created when adding a WCF Service Reference to a project. Here it is mocked by a namespace and a class.
'    ''' </summary>
'    Friend Class MessageClient
'        Implements IDisposable, ICommunicationObject

'#Region "IDisposable Support"
'        Private disposedValue As Boolean ' To detect redundant calls

'        ' IDisposable
'        Protected Overridable Sub Dispose(disposing As Boolean)
'            If Not Me.disposedValue Then
'                If disposing Then
'                    ' TODO: dispose managed state (managed objects).
'                End If

'                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'                ' TODO: set large fields to null.
'            End If
'            Me.disposedValue = True
'        End Sub

'        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
'        'Protected Overrides Sub Finalize()
'        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'        '    Dispose(False)
'        '    MyBase.Finalize()
'        'End Sub

'        ' This code added by Visual Basic to correctly implement the disposable pattern.
'        Public Sub Dispose() Implements IDisposable.Dispose
'            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'            Dispose(True)
'            GC.SuppressFinalize(Me)
'        End Sub
'#End Region

'#Region "ICommunicationObject"
'        Function SendMessage(item As Message) As Boolean
'            Return True
'        End Function

'        Public Sub Abort() Implements ICommunicationObject.Abort

'        End Sub

'        Public Function BeginClose(callback As AsyncCallback, state As Object) As IAsyncResult Implements ICommunicationObject.BeginClose

'        End Function

'        Public Function BeginClose(timeout As TimeSpan, callback As AsyncCallback, state As Object) As IAsyncResult Implements ICommunicationObject.BeginClose

'        End Function

'        Public Function BeginOpen(callback As AsyncCallback, state As Object) As IAsyncResult Implements ICommunicationObject.BeginOpen

'        End Function

'        Public Function BeginOpen(timeout As TimeSpan, callback As AsyncCallback, state As Object) As IAsyncResult Implements ICommunicationObject.BeginOpen

'        End Function

'        Public Sub Close() Implements ICommunicationObject.Close

'        End Sub

'        Public Sub Close(timeout As TimeSpan) Implements ICommunicationObject.Close

'        End Sub

'        Public Event Closed(sender As Object, e As EventArgs) Implements ICommunicationObject.Closed

'        Public Event Closing(sender As Object, e As EventArgs) Implements ICommunicationObject.Closing

'        Public Sub EndClose(result As IAsyncResult) Implements ICommunicationObject.EndClose

'        End Sub

'        Public Sub EndOpen(result As IAsyncResult) Implements ICommunicationObject.EndOpen

'        End Sub

'        Public Event Faulted(sender As Object, e As EventArgs) Implements ICommunicationObject.Faulted

'        Public Sub Open() Implements ICommunicationObject.Open

'        End Sub

'        Public Sub Open(timeout As TimeSpan) Implements ICommunicationObject.Open

'        End Sub

'        Public Event Opened(sender As Object, e As EventArgs) Implements ICommunicationObject.Opened

'        Public Event Opening(sender As Object, e As EventArgs) Implements ICommunicationObject.Opening

'        Public ReadOnly Property State As CommunicationState Implements ICommunicationObject.State
'            Get

'            End Get
'        End Property
'#End Region

'    End Class
'    ''' <summary>
'    ''' Empty class to provide something to work on in the 'payload'.
'    ''' </summary>
'    ''' <remarks></remarks>
'    Friend Class Message
'    End Class
'#End Region

'    ''' <summary>
'    ''' Various examples of how to use the ServiceHelper.
'    ''' </summary>
'    Friend Class Example

'        Private MessageList As List(Of Message)
'        ''' <summary>
'        ''' Typical Using block with a WCF Service that throws errors
'        ''' </summary>
'        Sub ExampleWithUsingBlock()
'            Using service As New WCF_ServiceReference.MessageClient
'                'This is the 'payload' code block:
'                For Each item As Message In MessageList
'                    If service.SendMessage(item) = True Then
'                        MessageList.Remove(item)
'                    End If
'                Next
'            End Using
'        End Sub
'        ''' <summary>
'        ''' Reworked WCF Service call as per Microsoft recommendation 'Do not use Using block for service calls'. See: http://msdn.microsoft.com/en-us/library/aa354510.aspx for details
'        ''' </summary>
'        Sub ExampleWithTryCatchBlock()
'            Dim service As New WCF_ServiceReference.MessageClient
'            'Exceptions that are thrown from communication methods on a Windows Communication Foundation (WCF) client are either expected or unexpected. 
'            Try
'                'Here is the 'payload' code block again
'                For Each item As Message In MessageList
'                    If service.SendMessage(item) = True Then
'                        MessageList.Remove(item)
'                    End If
'                Next
'                service.Close()
'            Catch timex As TimeoutException
'                'Expected exceptions from communication methods on a WCF client include TimeoutException, CommunicationException, and any derived class of CommunicationException. These indicate a problem during communication that can be safely handled by aborting the WCF client and reporting a communication failure. Because external factors can cause these errors in any application, correct applications must catch these exceptions and recover when they occur.
'                service.Abort()
'            Catch commex As CommunicationException
'                'Code that calls a client communication method must catch the TimeoutException and CommunicationException. One way to handle such errors is to abort the client and report the communication failure.
'                service.Abort()
'            Catch ex As Exception
'                'Typically there is no useful way to handle unexpected errors, so typically you should not catch them when calling a WCF client communication method.
'                'Throw
'            Finally
'                service.Dispose()
'            End Try
'        End Sub
'        ''' <summary>
'        ''' Implementation with ServiceHelper class, Using generic function, and a lambda Sub
'        ''' </summary>
'        Sub ExampleWithServiceHelperSub()
'            ServiceHelper.Using(Of WCF_ServiceReference.MessageClient)(
'             Sub(service)
'                 'Here is the 'payload' code block again
'                 For Each item As Message In MessageList 'Note that MessageList is in scope because it is part of the containing class
'                     If service.SendMessage(item) = True Then
'                         MessageList.Remove(item)
'                     End If
'                 Next
'             End Sub)
'        End Sub
'        ''' <summary>
'        ''' Implementation with ServiceHelper class, Using generic function, and a lambda Sub that includes a parameter
'        ''' </summary>
'        Sub ExampleWithServiceHelperSubWithParameter()
'            Dim messages As New List(Of Message)
'            ServiceHelper.Using(Of WCF_ServiceReference.MessageClient, List(Of Message))(messages,
'             Sub(service, msglist)
'                 'Here is the 'payload' code block again
'                 For Each item As Message In msglist 'Note that we pass in the message list as a parameter and thus do not rely on the containing class's property
'                     If service.SendMessage(item) = True Then
'                         MessageList.Remove(item)
'                     End If
'                 Next
'             End Sub)
'        End Sub
'        ''' <summary>
'        ''' Implementation with ServiceHelper class, Using generic function, and a lambda Function that returns a result
'        ''' </summary>
'        Sub ExampleWithServiceHelperFunction()
'            Dim result As Boolean = ServiceHelper.Using(Of WCF_ServiceReference.MessageClient, Boolean)(
'            Function(service) As Boolean
'                'Here is the 'payload' code block again
'                For Each item As Message In MessageList 'Note that MessageList is in scope because it is part of the containing class
'                    If service.SendMessage(item) = True Then
'                        MessageList.Remove(item)
'                    End If
'                Next
'                Return True
'            End Function)
'        End Sub
'        ''' <summary>
'        ''' Implementation with ServiceHelper class, Using generic function, and a lambda Function that includes a parameter and returns a result
'        ''' </summary>
'        Sub ExampleWithServiceHelperFunctionWithParameter()
'            Dim messages As New List(Of Message)
'            Dim result As Boolean = ServiceHelper.Using(Of WCF_ServiceReference.MessageClient, List(Of Message), Boolean)(messages,
'            Function(service, msglist) As Boolean
'                'Here is the 'payload' code block again
'                For Each item As Message In msglist 'Note that we pass in the message list as a parameter and thus do not rely on the containing class's property
'                    If service.SendMessage(item) = True Then
'                        MessageList.Remove(item)
'                    End If
'                Next
'                Return True
'            End Function)
'        End Sub

'    End Class
'End Namespace
#End Region
