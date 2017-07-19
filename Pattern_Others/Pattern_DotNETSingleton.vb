''' <summary>
''' Singleton using .NET framework features
''' </summary>
''' <remarks>This implementation relies on features on the .NET framework, see 'Exploring the Singleton Design Pattern' in MSDN for details</remarks>
Public NotInheritable Class Pattern_DotNETSingleton

    Private Sub New()
        'Prevent anyone from instantiating this class by making New private
    End Sub

    Public Shared ReadOnly Instance As Pattern_DotNETSingleton = New Pattern_DotNETSingleton

    'Public Function GetSingleton() As Pattern_Singleton
    '    Return New Pattern_Singleton
    'End Function

End Class
