@ModelType Knockout_MVC_Patterns.BookModel

<div class="container">
    <h2>Edit</h2>

    @Using (Html.BeginForm("Edit", "Library", FormMethod.Post, New With {.class = "form-horizontal", .role = "Form"}))
        @Html.AntiForgeryToken()

            @<h4>BookModel</h4>
            @<hr />
            @Html.ValidationSummary(True, "", New With {.class = "text-danger"})
            @Html.HiddenFor(Function(model) model.Id)

            @<div class="form-group">
                @Html.LabelFor(Function(model) model.Title, htmlAttributes:=New With {.class = "control-label col-md-2"})
                <div class="col-md-10">
                    @Html.EditorFor(Function(model) model.Title, New With {.htmlAttributes = New With {.class = "form-control"}})
                    @Html.ValidationMessageFor(Function(model) model.Title, "", New With {.class = "text-danger"})
                </div>
            </div>

            @<div class="form-group">
                @Html.LabelFor(Function(model) model.Author, htmlAttributes:=New With {.class = "control-label col-md-2"})
                <div class="col-md-10">
                    @Html.EditorFor(Function(model) model.Author, New With {.htmlAttributes = New With {.class = "form-control"}})
                    @Html.ValidationMessageFor(Function(model) model.Author, "", New With {.class = "text-danger"})
                </div>
            </div>

            @<div class="form-group">
                @Html.LabelFor(Function(model) model.Year, htmlAttributes:=New With {.class = "control-label col-md-2"})
                <div class="col-md-10">
                    @Html.EditorFor(Function(model) model.Year, New With {.htmlAttributes = New With {.class = "form-control"}})
                    @Html.ValidationMessageFor(Function(model) model.Year, "", New With {.class = "text-danger"})
                </div>
            </div>

            @<div class="form-group">
                <div class="col-md-offset-2 col-md-10">
                    <input type="submit" value="Save" class="btn btn-default" />
                </div>
            </div>
        
    End Using

    <div>
        @Html.ActionLink("Back to List", "Index")
    </div>

    @Section Scripts
        @Scripts.Render("~/bundles/jqueryval")
    End Section
</div>