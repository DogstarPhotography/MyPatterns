@ModelType FileUploadViewModel

@Code
    ViewData("Title") = "Index"
End Code

<h2>Index</h2>

<p>Demonstrates the use of an XMLHttpRequest to upload data and provide feedback for a progress bar.</p>

@Using Html.BeginForm("StoreFormData", "XHRUpload", FormMethod.Post, New With {.EncType = "multipart/form-data", .[class] = "form-horizontal"})
    @:<div class="form-group">
        @:<label for="Name">Name</label>
        @:<input name="Name" class="form-control" id="Name" type="text" value="">
    @:</div>
    @:<div class="form-group">
        @:<label for="Description">Description</label>
        @:<input name="Description" class="form-control" id="Description" type="text" value="">
    @:</div>
    @:<div class="form-group">
        @:<label for="File">File</label>
        @:<input type="file" name="File" id="File" class="form-control" /><br />
    @:</div>
    @:<input type="submit" name="Submit" id="Submit" value="Upload" />
End Using

<p></p>

<div class="progress">
    <div class="progress-bar" role="progressbar" aria-valuenow="60" aria-valuemin="0" aria-valuemax="100" style="width: 0%;">0%</div>
</div>


@Section Scripts

    <script type="text/javascript">

        // Bind the click events of the buttons to the function that initiate the upload
        $(document).ready(function () {
            "use strict";
            $("#Submit").bind("click", UploadFormData);
        });


        function UploadFormData() {
            "use strict";

            // Assemble the form data to match the properties of the FileUploadViewModel
            var body = new FormData();
            // Name
            var name = $("#Name").val();
            body.append("Name", name);
            // Description
            var description = $("#Description").val();
            body.append("Description", description);
            // File
            var file = $("#File")[0].files[0];
            body.append("File", file);

            // Create an XMLHttpRequest and add event handlers
            var xhr = new XMLHttpRequest();
            // Upload Progress Event Handler
            xhr.upload.onprogress = function (evt) {
                if (evt.lengthComputable) {
                    //update the progress bar
                    var percentComplete = Math.floor((evt.loaded / evt.total) * 100);
                    $(".progress-bar").css({ width: percentComplete + "%" });
                    $(".progress-bar").html(percentComplete + "%");
                }
            };
            // File Uploaded Event Handler
            xhr.upload.onload = function () {
                $(".progress-bar").css({ width: "100%" });
                $(".progress-bar").html("100%");
                if (xhr.readyState==4 && xhr.status==200)
                {
                    window.location.href = "/AjaxForms/AjaxFormsIndex";
                }
            };
            // Error Event Handler
            xhr.upload.onerror = function () {
                alert("Error");
            };            

            // Send the form data to the controller action
            xhr.open("POST", "/XHRUpload/StoreFormData/", true);
            xhr.send(body);

        }

    </script>
End Section
