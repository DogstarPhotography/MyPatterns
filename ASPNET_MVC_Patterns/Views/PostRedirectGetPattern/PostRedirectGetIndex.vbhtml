@ModelType ASPNET_MVC_Patterns.AjaxFormsViewModel

<h2>Post-Redirect-Get Pattern Example</h2>

@Using Html.BeginForm("PostData", "PostRedirectGetPattern")
    @<table>
        <tr><td>@Html.DisplayFor(Function(m) Model.Name)</td></tr>
        <tr><td><input type="hidden" id="Name" name="Name" value="@Model.Name" /></td></tr>
        <tr><td><input type="hidden" id="Description" name="Description" value="@Model.Description" /></td></tr>
        <tr><td><input type="hidden" id="Value" name="Value" value="@Model.Value" /></td></tr>
    </table>
    @:<input type="submit" />
End Using


