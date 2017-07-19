@ModelType IEnumerable(Of DataItem)

<table class="table">
    <tr class="row">
        <th>Name</th>
        <th>Description</th>
    </tr>
    @*
        The Html.DisplayFor(Function(m) m) will use the Views/Shared/DisplayTemplates/DataItem.vbhtml template to display each of the items in the model.
        By convention we don't need to use a For Each construct here as MVC will work out that the model is of type IEnumerable and automatically iterate through each item
    *@
    @Html.DisplayFor(Function(m) m)
</table>
