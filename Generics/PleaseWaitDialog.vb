Public Class PleaseWaitDialog
    ''' <summary>
    ''' Delegate for callback function
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Delegate Function PleaseWaitCallback() As Boolean

    Private myCallbackDelegate As PleaseWaitCallback
    Private myTimeout As Integer
    Private myTimerCycle As Integer
    ''' <summary>
    ''' Shows the please wait dialog while calling the callback function regularly as determined by the IntervalMilliseconds. 
    ''' Will close when the callback function returns true or after approximately TimeoutMilliseconds have elapsed
    ''' </summary>
    ''' <param name="CallbackFunction"></param>
    ''' <param name="TimeoutMilliseconds"></param>
    ''' <param name="IntervalMilliseconds"></param>
    ''' <remarks></remarks>
    Public Sub Setup(ByRef CallbackFunction As PleaseWaitCallback, ByVal TimeoutMilliseconds As UShort, ByVal IntervalMilliseconds As UShort)
        Try
            myCallbackDelegate = CallbackFunction
            myTimeout = TimeoutMilliseconds
            myTimerCycle = IntervalMilliseconds
            'Fire up the timer
            myTimer.Interval = myTimerCycle
            myTimer.Enabled = True
            'showdialog - this will block until the timer or the form calls close
            Me.ShowDialog()
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            'Ignore
        End Try
    End Sub

    Private Sub picDistraction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picDistraction.Click
        Try
            Me.Close()
        Catch ex As Exception
            'Ignore
        End Try
    End Sub
    ''' <summary>
    ''' Every time this comes round invoke the callback function and check the result, if the result is true or we have run out of time close the dialog
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub myTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles myTimer.Tick
        Try
            Dim Result As Boolean = myCallbackDelegate.Invoke
            myTimeout = myTimeout - myTimerCycle
            If Result = True Or myTimeout < 0 Then
                myTimer.Enabled = False
                Me.Close()
            End If
        Catch ex As Exception
            'Ignore
        End Try
    End Sub
End Class