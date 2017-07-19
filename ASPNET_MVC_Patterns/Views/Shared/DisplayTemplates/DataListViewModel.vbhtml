@ModelType DataListViewModel

<!-- DisplayFor Template for DataListViewModel -->
<table class="table">
    <tr class="row">
        <th>Name</th>
        <th>Description</th>
    </tr>
    @For Each item As DataItem In Model.Items
        @<tr class="row">
            <td>@item.Name</td>
            <td>@item.Description</td>
        </tr>
    Next
</table>
