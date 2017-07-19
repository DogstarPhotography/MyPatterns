@Code
    ViewData("Title") = "Home"
End Code

<div class="panel panel-default">
    <div class="panel-heading">
        <h3 class="panel-title">Home</h3>
    </div>
    <div class="panel-body">
        <ul class="list-group">
            <li class="list-group-item">@Html.ActionLink("Forms Demonstration", "Index", "Forms")</li>
            <li class="list-group-item">@Html.ActionLink("Library Demonstration", "Index", "Library")</li>
            <li class="list-group-item">@Html.ActionLink("Task List Demonstration", "Index", "TaskList")</li>
        </ul>
    </div>
</div>
