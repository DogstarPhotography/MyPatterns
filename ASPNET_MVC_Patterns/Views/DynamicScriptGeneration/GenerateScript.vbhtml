@ModelType ASPNET_MVC_Patterns.DynamicScriptGenerationData

<script>
    // Dynamically generated script for model: @Model.ID
    // Note that this script needs to be placed _after_ the jquery include

    // Bind the click event of our dynamically generated button to a function
    $(document).ready(function () {
        $("#btnSubmit_@Model.Id").on("click", buttonalert_@Model.Id());
    });

    // This function could be replaced by any other function. e.g a file upload with progress routine
    function buttonalert_@Model.Id() {
        alert("Hello from @Model.Name");
    }
</script>
