@Code
    ViewData("Title") = "Index"
End Code

<h2>Alerts</h2>

<div class="well">The following links display a Bootstrap Alert message when handling an error in the controller action.</div>

<p>
    @Html.ActionLink("Test the Danger Message error handling", "TestDangerMessage")
</p>
<p>
    @Html.ActionLink("Test the Warning Message error handling", "TestWarningMessage")
</p>
<p>
    @Html.ActionLink("Test the Info Message error handling", "TestInfoMessage")
</p>
<p>
    @Html.ActionLink("Test the Success Message error handling", "TestSuccessMessage")
</p>

<h2>Errors</h2>

<div class="well">The following links use various error handlers.</div>

<p>
    @Html.ActionLink("Test the OnException Error Handler", "TestOnException")
</p>
<p>
    @Html.ActionLink("Test the HandleError error handling", "TestHandleError")
</p>