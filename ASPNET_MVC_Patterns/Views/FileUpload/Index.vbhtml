@ModelType FileUploadViewModel

@Code
    ViewData("Title") = "Index"
End Code

<h2>Single File Upload</h2>

@Using Html.BeginForm("StoreFile", "Upload", FormMethod.Post, New With {.EncType = "multipart/form-data"})
    @:<input type="file" name="FileUpload1" /><br />
    @:<input type="submit" name="Submit" id="Submit" value="Upload" />
End Using

<h2>Single File Upload with Additional Data</h2>

@Using Html.BeginForm("StoreFormData", "Upload", FormMethod.Post, New With {.EncType = "multipart/form-data"})
    @Html.EditorFor(Function(m) m.Name)
    @Html.EditorFor(Function(m) m.Description)
    @:<input type="file" id="file" name="file" /><br />
    @:<input type="submit" name="Submit" id="Submit" value="Upload" />
End Using