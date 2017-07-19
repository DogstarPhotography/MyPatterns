@ModelType FormModel

@Code
    ViewData("Title") = "Forms"
End Code

@section ExtendedNavbar
    @Html.Raw(Model.Menu)
End Section

@html.raw(Model.Content)

@section Scripts
    @Html.Raw(Model.State)
    @Html.Raw(Model.Script)
End Section