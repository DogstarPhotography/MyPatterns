Imports System.Web.Mvc
Imports System.Web.Mvc.Html
Imports System.Runtime.CompilerServices

''' <summary>
''' Helpers that implement Bootstrap features.
''' </summary>
''' <remarks>
''' Requires an entry in the Views/web.config file: &lt;add namespace="ASPNET_MVC_Patterns.BootstrapHelpers" /&gt;.
''' See http://stackoverflow.com/questions/14723010/prove-me-wrong-vb-net-htmlhelper-extension-method-not-working-in-mvc-4-with-vs
''' </remarks>
Public Module BootstrapHelpers

    ''' <summary>
    ''' Create a form with the form-inline class attribute.
    ''' </summary>
    ''' <param name="helper">HtmlHelper</param>
    ''' <returns></returns>
    <Extension()>
    Public Function BeginInlineForm(ByVal helper As HtmlHelper) As MvcForm
        Return helper.BeginForm(Nothing, Nothing, FormMethod.Post, New With {Key .[class] = "form-inline"})
    End Function
    ''' <summary>
    ''' Create a form with the form-horizontal class attribute.
    ''' </summary>
    ''' <param name="helper">HtmlHelper</param>
    ''' <returns></returns>
    <Extension()>
    Public Function BeginHorizontalForm(ByVal helper As HtmlHelper) As MvcForm
        Return helper.BeginForm(Nothing, Nothing, FormMethod.Post, New With {Key .[class] = "form-horizontal"})
    End Function

End Module
