@ModelType ASPNET_MVC_Patterns.CustomTemplatesViewModel

@Code
    ViewData("Title") = "Custom Templates"
End Code

<h2>Custom Templates</h2>
@*
    We build our Bootstrap layout in this page: one row with two medium sized columns
*@
<div class="row">
    <div class="col-md-4">
        @* The Html.Action passes the data to the controller which returns a partial view that uses a template to render the html for the model *@
        @Html.Action("DataList", "CustomTemplates", New With {.data = Model.TableA})
    </div>
    <div class="col-md-4">
        @* The Html.Action passes the data to the controller which returns a partial view that uses a template to render the html for the model *@
        @Html.Action("DataItems", "CustomTemplates", New With {.data = Model.TableB})
    </div>
</div>
