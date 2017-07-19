@Code
    Dim current As String = ViewContext.RouteData.Values("controller").ToString()
End Code

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>@ViewBag.Title - My ASP.NET Application</title>
    <link href="~/Content/Site.css" rel="stylesheet" type="text/css" />
    <link href="~/Content/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <script src="~/Scripts/modernizr-2.6.2.js"></script>

</head>
<!-- We need to add some padding here to move the content down to account for the size of the navbar-->
<body style="padding-top: 60px;">
    <div class="navbar navbar-inverse navbar-fixed-top">
        <div class="container-fluid">
            <div class="navbar-header">
                <!-- The next element creates the three line button when the navbar is collapsed, don't remove the span elements -->
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                @*@Html.ActionLink("Application Name", "Index", "Home", New With {.area = ""}, New With {.class = "navbar-brand"})*@
            </div>
            <div class="navbar-collapse collapse">
                @Html.Action("NavBar", "SharedViews", New With {.Current = current})
            </div>
        </div>
    </div>

    <div class="container body-content">
        @* Alert Messages to User *@
        @If TempData("Alert_Success_Message") IsNot Nothing Then
            @<div class="alert alert-success" role="alert">@TempData("Alert_Success_Message")</div>
        End If
        @If TempData("Alert_Info_Message") IsNot Nothing Then
            @<div class="alert alert-info" role="alert">@TempData("Alert_Info_Message")</div>
        End If
        @If TempData("Alert_Warning_Message") IsNot Nothing Then
            @<div class="alert alert-warning" role="alert">@TempData("Alert_Warning_Message")</div>
        End If
        @If TempData("Alert_Danger_Message") IsNot Nothing Then
            @<div class="alert alert-danger" role="alert">@TempData("Alert_Danger_Message")</div>
        End If
        @* Main Content *@
        @RenderBody()
        <hr />
        @* 
            Validation Summary within dismissable Bootstrap Alert 
            Simply call ModelState.AddModelError within controller to add items
        *@
        @If ViewData.ModelState.Any(x => x.Value.Errors.Any()) Then
            @<div class="alert alert-danger">
                @<a class="close" data-dismiss="alert">&times;</a>
                @Html.ValidationSummary()
            @</div>
        End If
        <footer>
            <div class="well well-sm">
                <p>&copy; @DateTime.Now.Year - Robin G Brown</p>
            </div>
        </footer>
    </div>

    <script src="~/Scripts/jquery-2.1.3.min.js"></script>
    <script src="~/Scripts/bootstrap.min.js"></script>

    @RenderSection("Scripts", False)

</body>

</html>
