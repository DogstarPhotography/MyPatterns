Imports System.ComponentModel
Imports MyPatterns.TypeCodes
Imports MyPatterns.CategoryCodes

Public Enum TypeCodes
    <Description("Type A")> TypeA = 0
    <Description("Type B")> TypeB
    <Description("Type C")> TypeC
End Enum

Public Enum CategoryCodes
    <Description("General")> General = 0
    <Description("Specific")> Specific
    <Description("Other")> Other
End Enum

Public NotInheritable Class Constants

#Region "Resources"

    Private Shared myTypes As Dictionary(Of TypeCodes, ConstantItem) = Nothing

    Public Shared ReadOnly Property Types() As Dictionary(Of TypeCodes, ConstantItem)
        Get
            If myTypes Is Nothing Then
                myTypes = New Dictionary(Of TypeCodes, ConstantItem)
                BuildTypes(myTypes)
            End If
            Return myTypes
        End Get
    End Property

    Private Shared Sub BuildTypes(ByRef dict As Dictionary(Of TypeCodes, ConstantItem))
        With dict
            .Add(TypeCodes.TypeA, New ConstantItem(TypeCodes.TypeA.Description, "Type A are..."))
            .Add(TypeCodes.TypeB, New ConstantItem(TypeCodes.TypeB.Description, "Type B are..."))
            .Add(TypeCodes.TypeC, New ConstantItem(TypeCodes.TypeC.Description, "Type C are..."))
        End With
    End Sub

#End Region

#Region "Categories"

    Private Shared myCategories As Dictionary(Of CategoryCodes, ConstantItem) = Nothing

    Public Shared ReadOnly Property Categories() As Dictionary(Of CategoryCodes, ConstantItem)
        Get
            If myCategories Is Nothing Then
                myCategories = New Dictionary(Of CategoryCodes, ConstantItem)
                BuildCategories(myCategories)
            End If
            Return myCategories
        End Get
    End Property

    Private Shared Sub BuildCategories(ByRef dict As Dictionary(Of CategoryCodes, ConstantItem))
        With dict
            .Add(CategoryCodes.General, New ConstantItem(CategoryCodes.General.Description, "General category"))
            .Add(CategoryCodes.Specific, New ConstantItem(CategoryCodes.Specific.Description, "Specific category"))
            .Add(CategoryCodes.Other, New ConstantItem(CategoryCodes.Other.Description, "Other category"))
        End With
    End Sub

#End Region

End Class

Public NotInheritable Class ConstantItem

#Region "Constructors"
    ''' <summary>
    ''' Default constructor.
    ''' </summary>
    Public Sub New()
        'Do nothing
    End Sub
    ''' <summary>
    ''' Simple constructor.
    ''' </summary>
    Sub New(value As String)
        Me.Name = value
        Me.Description = value
    End Sub
    ''' <summary>
    ''' Proper constructor.
    ''' </summary>
    Sub New(name As String, description As String)
        Me.Name = name
        Me.Description = description
    End Sub

#End Region

    Property Name As String
    Property Description As String

    ''' <summary>
    ''' See http://stackoverflow.com/questions/293215/default-properties-in-vb-net
    ''' </summary>
    Public Shared Widening Operator CType(value As String) As ConstantItem
        Return New ConstantItem(value)
    End Operator
    ''' <summary>
    ''' See http://stackoverflow.com/questions/293215/default-properties-in-vb-net
    ''' </summary>
    Public Shared Widening Operator CType(value As ConstantItem) As String
        Return value.Name
    End Operator

End Class

Public Class ConstantsExample

    Public Sub UseConstant()

        Dim value As String = Constants.Types(TypeA)
        Dim category As String = Constants.Categories(General)

    End Sub

End Class

