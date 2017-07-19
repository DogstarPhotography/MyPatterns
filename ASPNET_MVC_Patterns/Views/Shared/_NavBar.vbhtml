@ModelType ASPNET_MVC_Patterns.NavBarViewModel

<ul class="nav navbar-nav">
    @If Model.Current = "ajaxforms" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Ajax Forms", "AjaxFormsIndex", "AjaxForms")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Ajax Forms", "AjaxFormsIndex", "AjaxForms")</li>
    End If
    @If Model.Current = "backgroundtaskwithprogress" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Background Task with Progress", "Start", "BackgroundTaskWithProgress")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Background Task with Progress", "Start", "BackgroundTaskWithProgress")</li>
    End If
    @If Model.Current = "fileupload" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("File Upload", "Index", "FileUpload")</li>
    Else
        @<li role="presentation">@Html.ActionLink("File Upload", "Index", "FileUpload")</li>
    End If
    @If Model.Current = "postredirectgetpattern" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Post Redirect Get Pattern", "PostRedirectGetIndex", "PostRedirectGetPattern")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Post Redirect Get Pattern", "PostRedirectGetIndex", "PostRedirectGetPattern")</li>
    End If
    @If Model.Current = "xhrupload" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("XMLHttpRequest Upload", "Index", "XHRUpload")</li>
    Else
        @<li role="presentation">@Html.ActionLink("XMLHttpRequest Upload", "Index", "XHRUpload")</li>
    End If
    @If Model.Current = "bootstrapmodal" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Bootstrap Modal", "Index", "BootstrapModal")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Bootstrap Modal", "Index", "BootstrapModal")</li>
    End If
    @If Model.Current = "dynamicscriptgeneration" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Dynamic Script Generation", "Index", "DynamicScriptGeneration")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Dynamic Script Generation", "Index", "DynamicScriptGeneration")</li>
    End If
    @If Model.Current = "customtemplates" Then
        @<li role="presentation" class="disabled">@Html.ActionLink("Custom Templates", "Index", "CustomTemplates")</li>
    Else
        @<li role="presentation">@Html.ActionLink("Custom Templates", "Index", "CustomTemplates")</li>
    End If
</ul>


