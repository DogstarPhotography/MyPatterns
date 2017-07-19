Imports System.Reflection

Public Class Binder

    Public Event Message(ByVal Content As String)

    Public Sub Test(ByVal Filename As String)

        'Create object
        Dim Target As Assembly = Assembly.GetExecutingAssembly() '= Assembly.LoadFrom(Filename)

        'List components
        Dim Modules() As [Module] = Target.GetModules()
        RaiseEvent Message("Modules")
        For Each curMod As [Module] In Modules
            RaiseEvent Message("Module - " & curMod.ToString)
        Next

        Dim Types() As [Type] = Target.GetTypes()
        RaiseEvent Message("Types")
        For Each curType As [Type] In Types
            RaiseEvent Message("Type - " & curType.ToString)
        Next

        Dim ExportedTypes() As [Type] = Target.GetExportedTypes()
        RaiseEvent Message("Exported Types")
        For Each curType As [Type] In ExportedTypes
            RaiseEvent Message("Type - " & curType.ToString)
        Next

    End Sub
End Class
