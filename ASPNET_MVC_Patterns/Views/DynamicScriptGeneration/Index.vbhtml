@ModelType ASPNET_MVC_Patterns.DynamicScriptGenerationViewModel
@Code
    ViewData("Title") = "Dynamic Script Generation"
End Code

<h2>Index</h2>

@For Each item As DynamicScriptGenerationData In Model.Data
    Html.RenderPartial("DisplayData", item)
Next

@Section Scripts
    @For Each item As DynamicScriptGenerationData In Model.Data
        Html.RenderPartial("GenerateScript", item)
    Next
End Section
