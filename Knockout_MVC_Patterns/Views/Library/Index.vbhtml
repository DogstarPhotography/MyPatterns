@ModelType Knockout_MVC_Patterns.Librarymodel

@* Typical MVC Code *@
<div class="container">
    <h2 style="text-align: center">@model.Name</h2>
    <table class="table table-bordered table-striped table-condensed table-hover">
        <thead>
            <tr>
                <th>Title</th>
                <th>Author</th>
                <th>Years</th>
                <th>Actions</th>
            </tr>
        </thead>
        <tbody>
            @For Each book In Model.GetBooks()
                @<tr>
                    <td>@book.Title</td>
                    <td>@book.Author</td>
                    <td>@book.Year</td>
                    <td>
                        @Html.ActionLink("Edit", "Edit", New With {.id = book.Id}, New With {.class = "btn btn-primary btn-xs"})
                        @Html.ActionLink("Remove", "Remove", New With {.id = book.Id}, New With {.class = "btn btn-primary btn-xs"})
                    </td>
                </tr>
            Next
        </tbody>
    </table>
    @Html.ActionLink("Add new book", "Add", Nothing, New With {.class = "btn btn-primary"})
</div>

<hr />

@* Knockout Version *@
<div class="container">
    <h2 style="text-align: center"><span data-bind="text: Name"></span></h2>
    <table class="table table-bordered table-striped table-condensed table-hover">
        <thead>
            <tr>
                <th>Title</th>
                <th>Author</th>
                <th>Years</th>
                <th>Actions</th>
            </tr>
        </thead>
        <tbody data-bind="foreach: Books">
                <tr>
                    <td data-bind="text: Title"></td>
                    <td data-bind="text: Author"></td>
                    <td data-bind="text: Year"></td>
                    <td>
                        <!-- Note that $root here is a reference to the ViewModel 
                            so $root.edit calls the edit function of the root object 
                            of the current item in the foreach -->
                        <a href="#" data-bind="click: $root.edit" class="btn btn-primary btn-xs">Edit</a>
                        <a href="#" data-bind="click: $root.remove" class="btn btn-primary btn-xs">Remove</a>
                    </td>
                </tr>
        </tbody>
    </table>
    <a href="#" data-bind="click: add" class="btn btn-primary">Add new book</a>
</div>

@Section Scripts

<script type="text/javascript">

    var libraryViewModel = function () {
        var self = this;
        self.Name = ko.observable();
        self.Books = ko.observableArray();

        // Initial Data - these ajax calls will run once when the object is created
        // Note that we use Ajax calls here is they integrate well with ASP.NET MVC
        $.ajax({
            url: '@Url.Action("koGetName")',
            cache: false,
            type: 'GET',
            contentType: 'application/json; charset=utf-8',
            data: {},
            success: function (data) {
                self.Name(data);
            }
        });
        $.ajax({
            url: '@Url.Action("koGetBooks")',
            cache: false,
            type: 'GET',
            contentType: 'application/json; charset=utf-8',
            data: {},
            success: function (data) {
                self.Books(data);
            }
        });

        // Remove
        self.remove = function (book) {
            var id = book.Id;
            $.ajax({
                url: '@Url.Action("koRemove")',
                cache: false,
                type: 'GET',
                contentType: 'application/json; charset=utf-8',
                data: JSON.stringify({ id: id }),
                dataType: 'json',
                success: function (data) {
                    self.Books(data);
                }
            });
        };

        // Add
        self.add = function () {
            $.ajax({
                url: '@Url.Action("koAdd")',
                cache: false,
                type: 'GET',
                contentType: 'application/json; charset=utf-8',
                data: {},
                success: function (data) {
                    self.Books(data);
                }
            });
        };

        // Edit
        self.edit = function (book) {
            var id = book.Id;
            location.href = "Library/Edit/" + id;
        };

    };

    // Applying Bindings
    ko.applyBindings(new libraryViewModel());

</script>


End Section
