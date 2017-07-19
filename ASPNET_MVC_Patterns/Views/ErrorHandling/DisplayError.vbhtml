@ModelType  Exception

@Code
    ViewData("Title") = "DisplayError"
End Code

<h2>Display Error</h2>

<table class="table table-bordered">
    <tr class="row">
        <td>Description</td>
    </tr>
    <tr class="row">
        <td>@Model.Message</td>
    </tr>
</table>
