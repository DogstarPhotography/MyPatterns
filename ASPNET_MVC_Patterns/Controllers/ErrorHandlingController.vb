Imports System.Web.Mvc
Imports System.Web.Routing

Namespace Controllers
    ''' <summary>
    ''' Demonstrates various error handling options.
    ''' </summary>
    ''' <remarks>
    ''' Reference: http://www.codeguru.com/csharp/.net/net_asp/mvc/handling-errors-in-asp.net-mvc-applications.htm but note that there is a major flaw with the code sample for OnException - see below for working code
    ''' </remarks>
    Public Class ErrorHandlingController
        Inherits Controller
        ''' <summary>
        ''' start here
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function Index() As ActionResult
            Return View()
        End Function

#Region "Alert Messages"
        ''' <summary>
        ''' This method throws an error that will be displayed on the page thanks to some code and HTML in the Layout.
        ''' </summary>
        Function TestDangerMessage() As ActionResult
            Try
                Throw New Exception("This is an error message from an exception thrown in the TestDangerMessage method of the ErrorHandling Controller.")
            Catch ex As Exception
                TempData.Add("Alert_Danger_Message", ex.Message)
            End Try
            Return View("Index")
        End Function
        ''' <summary>
        ''' This method throws an error that will be displayed on the page thanks to some code and HTML in the Layout.
        ''' </summary>
        Function TestWarningMessage() As ActionResult
            Try
                Throw New Exception("This is an error message from an exception thrown in the TestWarningMessage method of the ErrorHandling Controller.")
            Catch ex As Exception
                TempData.Add("Alert_Warning_Message", ex.Message)
            End Try
            Return View("Index")
        End Function
        ''' <summary>
        ''' This method throws an error that will be displayed on the page thanks to some code and HTML in the Layout.
        ''' </summary>
        Function TestInfoMessage() As ActionResult
            Try
                Throw New Exception("This is an error message from an exception thrown in the TestInfoMessage method of the ErrorHandling Controller.")
            Catch ex As Exception
                TempData.Add("Alert_Info_Message", ex.Message)
            End Try
            Return View("Index")
        End Function
        ''' <summary>
        ''' This method throws an error that will be displayed on the page thanks to some code and HTML in the Layout.
        ''' </summary>
        Function TestSuccessMessage() As ActionResult
            Try
                Throw New Exception("This is an error message from an exception thrown in the TestSuccessMessage method of the ErrorHandling Controller.")
            Catch ex As Exception
                TempData.Add("Alert_Success_Message", ex.Message)
            End Try
            Return View("Index")
        End Function

#End Region
        ''' <summary>
        ''' This method throws an error that will be handled by the OnException method below.
        ''' </summary>
        Function TestOnException() As ActionResult
            Throw New Exception("This is an error message from an exception thrown in the TestOnException method of the ErrorHandling Controller.")
        End Function
        ''' <summary>
        ''' This method throws an error that is handled by the default error handler, it requires an error page in the shared views folder.
        ''' </summary>
        <HandleError> Function TestHandleError() As ActionResult
            Throw New Exception("This is an error message from an exception thrown in the TestHandleError method of the ErrorHandling Controller.")
        End Function
        ''' <summary>
        ''' Display an exception.
        ''' </summary>
        Function DisplayError() As ActionResult
            'Fetch the stored data
            Dim ex As Exception = CType(TempData.Item("OnException_Exception"), Exception)
            Dim controller As String = TempData.Item("OnException_Controller").ToString
            Dim action As String = TempData.Item("OnException_Action").ToString
            'Display
            Return View(ex)
        End Function
        ''' <summary>
        ''' The controller can implement a handler that will catch unhandled errors and deal with them.
        ''' </summary>
        Protected Overrides Sub OnException(filterContext As ExceptionContext)
            MyBase.OnException(filterContext)
            'If the exception is unhandled we can deal with it
            If filterContext.ExceptionHandled = False Then
                'Make sure we mark this exception as handled!
                filterContext.ExceptionHandled = True
                'Get the controller, action, and exception making sure we do not have null values
                Dim controller As String = If(filterContext.RouteData.Values("controller").ToString, "")
                Dim action As String = If(filterContext.RouteData.Values("action").ToString, "")
                Dim exc As Exception = If(filterContext.Exception, New Exception("Unknown error."))
                'We need to store the data temporarily
                TempData.Clear()
                TempData.Add("OnException_Exception", exc)
                TempData.Add("OnException_Controller", controller)
                TempData.Add("OnException_Action", action)
                'TODO: This is a good place to log any error
                'The OnException sub cannot return a view or redirect to an action so we must fill out the ExceptionContext's Result property instead
                filterContext.Result = RedirectToAction("DisplayError")
            End If
        End Sub

    End Class
End Namespace