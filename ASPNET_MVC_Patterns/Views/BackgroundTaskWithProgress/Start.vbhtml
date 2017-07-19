@ModelType Models.TaskRequestViewModel

<h2>Start</h2>

<div>
    @Using Html.Beginform("TaskInitiation", "BackgroundTaskWithProgress")

    @<table class="table">
        <tr>
            <th>Name</th>
            <th>Description</th>
            <th>Select</th>
        </tr>
        @For idx = 0 To Model.Tasks.Count - 1
            Dim index As Integer = idx
            @* For each task we add a row with the name, description, and a checkbox.
                We MUST use hidden input fields and pass the data by index in order to 
                output html that will post and correctly create a TaskRequestViewModel during model binding on the return trip*@
            @<tr>
                <td>
                    @Html.DisplayFor(Function(m) m.Tasks.Item(index).Name)
                    @Html.HiddenFor(Function(m) m.Tasks.Item(index).Name)
                </td>
                <td>
                    @Html.DisplayFor(Function(m) m.Tasks.Item(index).Description)
                    @Html.HiddenFor(Function(m) m.Tasks.Item(index).Name)
                </td>
                <td>
                    @Html.CheckBoxFor(Function(m) m.Tasks.Item(index).Selected)
                </td>
            </tr>
        Next
        <tr>
            <td></td>
            <td></td>
            <td><button type="submit">Start</button></td>
        </tr>
    </table>
    end using
</div>

