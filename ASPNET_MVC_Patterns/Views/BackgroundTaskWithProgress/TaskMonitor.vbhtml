
<h2>Task Monitor</h2>

<div id="TaskProgress">
    @Html.Action("GetTaskProgress")
</div>


@Section Scripts
<script type="text/javascript">
    // Get the current status every second
    $(document).ready(function () {
        setInterval(FetchStatus, 1000);
    });
    //Load the given div with the progress of the task
    function FetchStatus() {
        $('#TaskProgress').load('/BackgroundTaskWithProgress/GetTaskProgress');
    }
</script>
End Section