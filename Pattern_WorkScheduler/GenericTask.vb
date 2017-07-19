﻿Public Class GenericTask(Of S As {IState(Of O), New}, O)
    Implements ITask(Of S, O)

    Public Event Completed(ByVal Task As ITask(Of S, O)) Implements ITask(Of S, O).Completed

    Private MyState As S
    Private CompleteFlag As Boolean = False

    Public Sub New()
        MyState = New S
    End Sub

    Public Property State() As S Implements ITask(Of S, O).State
        Get
            Return MyState
        End Get
        Set(ByVal value As S)
            value = MyState
        End Set
    End Property

    Public Sub RunTask() Implements ITask(Of S, O).RunTask
        'Do something
        'CODE_HERE
        'Once we are finished tell our consumer, passing a reference to this task
        CompleteFlag = True
        RaiseEvent Completed(Me)
    End Sub

#Region "Properties"
    Public ReadOnly Property Complete() As Boolean Implements ITask(Of S, O).Complete
        Get
            Return CompleteFlag
        End Get
    End Property

    Public Property Description() As String Implements ITask(Of S, O).Description
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property

    Public Property Name() As String Implements ITask(Of S, O).Name
        Get
            Return Me.ToString
        End Get
        Set(ByVal value As String)

        End Set
    End Property
#End Region
End Class
