@ModelType Models.TaskProgressViewModel

<table class="table">
    <tr>
        <th>Name</th>
        <th>Status</th>
        <th>Progress</th>
    </tr>
    @For Each item As Models.Task In Model.Tasks
        @<tr>
            <td>
                @item.Name
            </td>
            <td>
                @item.Status
            </td>
            <td>
                <div class="progress" style="width: 100px">
                    <div id="barProgress" class="progress-bar" role="progressbar" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100" style="width: @item.Progress%;">
                        @item.Progress%
                    </div>
                </div>
            </td>
        </tr>
    Next
    @* If all the processing is finished we should display a completion option to the user. *@
    @If Model.Finished = True Then
        @<tr>
            <td colspan="3">All tasks completed: @Html.ActionLink("Return to Start", "Start", "BackgroundTaskWithProgress")</td>
        </tr>
    End If
</table>





