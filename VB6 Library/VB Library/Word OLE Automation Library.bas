Attribute VB_Name = "modWordOLEAutomation"
Option Explicit
'---------------------------------------------------------
'   WORD fuctions and calls set up by
'               Mark Matthews,
'               Aug-Sept 1995.
'   Word.bas - a set of common functions and the required
'   declarations for Word OLE Automation. These were
'   derived using Word 6.
'   Date : 01/08/95 onwards
'----------------------------------------------------------------

'Define the object to use everywhere:
Global objWordBasic As Object

'Constants for the paper tray printer settings:
Global Const BINPAPERCASSETTE = 0
Global Const BINMANUALFEED = 1
Global Const BINCONTROLPANEL = 2

Function AttatchToWordFax() As Integer
'this routine should make sure any subsequently Faxed documents (using
'SendWordFax) are sent With formatting in place.
'Note that is also affects the other settings listed below.

    On Error Resume Next

    AttatchToWordFax = objWordBasic.ToolsOptionGeneral(1, 0, 0, 0, 1, 1, 1, 1, 1, 6, 1, -1)

    On Error GoTo 0

'The FULL macro code is:
'ToolsOptionsGeneral .Pagination = 1, .WPHelp = 0, .WPDocNavKeys = 0,
'   .BlueScreen = 0, .ErrorBeeps = 1, .Effects3d = 1, .UpdateLinks = 1,
'   .SendMailAttach = 1, .RecentFiles = 1, .RecentFileCount = "6",
'   .Units = 1, .ButtonFieldClicks = -1

End Function

Sub ChangeWordDirectory(strNewDir As String)

    On Error Resume Next

    objWordBasic.ChDefaultDir strNewDir, 0
    
    On Error GoTo 0
    
'FileOpen .Name = "BLNK000.DOC", .ConfirmConversions = 0, .ReadOnly = 0, .AddToMru = 0, .PasswordDoc = "", .PasswordDot = "", .Revert = 0, .WritePasswordDoc = "", .WritePasswordDot = ""
'
'FileSaveAs .Name = "\\primus\DEVELOP\MARK\VERSION3\BLNK000.DOT", .Format = 1, .LockAnnot = 0, .Password = "", .AddToMru = 1, .WritePassword = "", .RecommendReadOnly = 0, .EmbedFonts = 0, .NativePictureFormat = 0, .FormsData = 0
'
'FileSaveAs .Name = "BLNK000.DOT", .Format = 1, .LockAnnot = 0, .Password = "", .AddToMru = 1, .WritePassword = "", .RecommendReadOnly = 0, .EmbedFonts = 0, .NativePictureFormat = 0, .FormsData = 0
'
'FileClose
'
'FileNew .Template = "Blnk000", .NewTemplate = 0
'
'FileSaveAs .Name = "trad0.doc", .Format = 0, .LockAnnot = 0, .Password = "", .AddToMru = 1, .WritePassword = "", .RecommendReadOnly = 0, .EmbedFonts = 0, .NativePictureFormat = 0, .FormsData = 0
'
'FileClose

End Sub

Sub CloseWordDocument()
'this routine will close the open word document

'the Full macro code was:
'DocClose
    
    On Error Resume Next
    
    objWordBasic.DocClose

    On Error GoTo 0

End Sub

Sub CloseWordFile()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'Closes the currently open Word file -this
'will invoke as FileSaveAs if not done already
    
    On Error Resume Next

    objWordBasic.FileClose

    On Error GoTo 0

End Sub

Sub CloseWordHeaderFooter()

   On Error Resume Next
   
   objWordBasic.CloseViewHeaderFooter

   On Error GoTo 0
'ViewHeader
'GoToHeaderFooter
'EditClear -22
'CloseViewHeaderFooter

End Sub

Sub CopyWordToClipboard()
'this routine will copyt all of the highlighted text in
'a word document:

    On Error Resume Next

    objWordBasic.EditCopy

    On Error GoTo 0

'the full macro code for this statement was:
'EditCopy

End Sub

Sub CreateWordBookmark(strBookmarkName As String)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'The strBookmark must be exactly the same as the bookmark
'in the word document
    On Error Resume Next

    objWordBasic.EditGoTo (strBookmarkName)

    On Error GoTo 0



End Sub

Function CreateWordObject() As Integer
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will create the Object so it can be used:

    On Error GoTo CreateWordObject_ErrorHandler

    Set objWordBasic = CreateObject("Word.Basic")
    CreateWordObject = True

    Exit Function

CreateWordObject_ErrorHandler:
    Set objWordBasic = Nothing
    CreateWordObject = False
    MsgBox "Error: " & Err & " " & Error$(Err), 16, "ERROR!!"
    Exit Function

End Function

Sub CutWordFromClipboard()
'this routine will cut all of the highlighted text into
'the clipboard:

    On Error Resume Next

    objWordBasic.EditCut

    On Error GoTo 0

'the full macro code for this statement was:
'EditCut

End Sub

Sub FaxWordDoc()
'this routine will make word send a fax message through
'MS Mail. this will be the active document

'the Full macro code was:
'FileSendMail
'DocClose
    
    On Error GoTo FaxWordDoc_ErrorHandler
    
    objWordBasic.FileSendMail

    Exit Sub

FaxWordDoc_ErrorHandler:
    MsgBox "Error in trying to Fax Word Document:" & Chr$(10) & "#" & Err & ": " & Error$(Err)
    Resume Next
    Exit Sub

End Sub

Function FindFoundWordDoc() As Integer
    'this function returns -1 if the text was found, and 0 if it was not
                      '   TRUE                          FALSE
'An EditFindFound() Example is in WordBasic Help
Dim intWordResult As Integer
    On Error Resume Next

    intWordResult = objWordBasic.EditFindFound()
    FindFoundWordDoc = intWordResult

    On Error GoTo 0

End Function

Sub FindWordText(strFindText As String)
    On Error Resume Next

    objWordBasic.EditFind (strFindText)

    On Error GoTo 0



Rem do match case
'EditFind .Find = "dolly", .Direction = 0, .MatchCase = 1, .WholeWord = 0,
'.PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 1
'LineUp 33

Rem don't match case
'EditFind .Find = "dolly", .Direction = 0, .MatchCase = 0, .WholeWord = 0,
'.PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 1

Rem EditFind .Find = "INSERTCOURSEFEE", .Direction = 0, .MatchCase = 1,
'.WholeWord = 1, .PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 2
    
End Sub

Function FindWordTextMatchCase(strFindText As String) As Integer
    'sub to find MATCHED CASE text
    
    On Error GoTo FindWordTextMatchCase_ErrorHandler
    FindWordTextMatchCase = objWordBasic.EditFind(strFindText, 0, 1, 0, 0, 0, 0, 1)
    
    Exit Function

FindWordTextMatchCase_ErrorHandler:
    MsgBox "Error " & Err & " " & Error$(Err), MB_ICONSTOP, "In FindWordTextMatchCase"
    Exit Function

Rem do match case
'EditFind .Find = "dolly", .Direction = 0, .MatchCase = 1, .WholeWord = 0,
'.PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 1
'LineUp 33

Rem don't match case
'EditFind .Find = "dolly", .Direction = 0, .MatchCase = 0, .WholeWord = 0,
'.PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 1
Rem EditFind .Find = "INSERTCOURSEFEE", .Direction = 0, .MatchCase = 1,
'.WholeWord = 1, .PatternMatch = 0, .SoundsLike = 0, .Format = 0, .Wrap = 2
End Function

Sub FormatSmallWordFooter()

'WORD BASIC IS:
'FormatStyle .Name = "Footer", .BasedOn = "Normal", .NextStyle = "Footer", .Type = 0, .AddToTemplate = 0, .Define
'FormatDefineStyleFont .Points = "8", .Underline = 0, .Color = 0, .Strikethrough = 0, .Superscript = 0, .Subscript = 0, .Hidden = 0, .SmallCaps = 0, .AllCaps = 0, .Spacing = "", .Position = "", .Kerning = 0, .KerningMin = "", .Tab = "0", .Font = "Times New Roman", .Bold = 0, .Italic = 0
'Style "Footer"

    On Error Resume Next
    
     objWordBasic.FormatStyle "Footer", "Normal", "Footer", 0, 0
     objWordBasic.FormatDefineStyleFont "8", 0, 0, 0, 0, 0, 0, 0, 0, "", "", 0, "", "0", "Times New Roman", 0, 0
     objWordBasic.Style "Footer"
    
    On Error GoTo 0

End Sub

Sub GotoNextWordPage()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will go to the next current document page

    On Error Resume Next

    objWordBasic.EditGoTo ("")

    On Error GoTo 0

'EditGoTo.Destination = ""      this should be next Page
'EditGoTo.Destination = "-"     this should be prev page

End Sub

Sub GotoPreviousWordPage()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will go to the PREVIOUS current document page

    On Error Resume Next

    objWordBasic.EditGoTo ("-")

    On Error GoTo 0

'EditGoTo.Destination = ""      this should be next Page
'EditGoTo.Destination = "-"     this should be prev page


End Sub

Sub GotoStartOfWordDoc()
'routine to send the cursor to the very begining of a document
    On Error Resume Next
 
    objWordBasic.StartOfDocument

    On Error GoTo 0
 
End Sub

Sub GotoTopWordDoc()
'sub to goto the very top of a word document
    On Error Resume Next
    
    objWordBasic.EditGoTo "l0"
    
    On Error GoTo 0

'EditGoTo .Destination = "l0"
End Sub

Function GotoWordBookmark(strBookmark As String) As Integer
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'The strBookmark must be exactly the same as the bookmark
'in the word document
    On Error Resume Next

    objWordBasic.EditGoTo (strBookmark)

    On Error GoTo 0

End Function

Sub InsertToSmallWordFooter(ByVal strInsertThisText As String)

    'to set the footer to font of TimesnewRoman, 6 point text.
    On Error Resume Next

    objWordBasic.ViewHeader
    objWordBasic.GoToHeaderFooter
    objWordBasic.Font "Times New Roman"
    objWordBasic.FontSize 6
    objWordBasic.Insert strInsertThisText
    objWordBasic.CloseViewHeaderFooter

    On Error GoTo 0

End Sub

Sub InsertToWordFooter(strInsertText As String)
'This sub will open the Word document Footer and place the text in it,
'on the first line down.

    On Error Resume Next
    Call ViewWordHeaderFooter
    Call ToggleWordHeaderFooter  'to move into Footer
    Call MoveDownWordLines(1)
    Call InsertWordText(strInsertText)
    Call CloseWordHeaderFooter
    On Error GoTo 0

End Sub

Sub InsertToWordHeader(strTextToInsert As String)
'This sub will open the Word document Header and place the text in it
'on the first line.

    On Error Resume Next
    Call ViewWordHeaderFooter
    Call InsertWordText(strTextToInsert)
    Call CloseWordHeaderFooter
    On Error GoTo 0

End Sub

Sub InsertWordFile(strFileName As String)
'sub to insert the whole of a Word file into a document.
'this requires the full file name And Path.

    On Error Resume Next

    objWordBasic.InsertFile (strFileName)

    On Error GoTo 0

'the macro code is:
'InsertFile.Name = "C:\PRICES\PRICES.DOC"

End Sub

Sub InsertWordPageBreak()
'sub to insert a page break into a document on the
'current cursor line

    On Error Resume Next

    objWordBasic.InsertBreak (0)

    On Error GoTo 0

'the full macro code for this is:
'InsertBreak .Type = 0
End Sub

Function InsertWordPageNumbers(intType As Integer, intPosition As Integer, intFirstPage As Integer) As Integer
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This inserts page numbering into the document
'InsertPageNumbers .Type = 1, .Position = 1, .FirstPage = 1
'To number the first page we have: FirstPage is 1=true(yes), 0=false(no)
'type refers to: 0=top page, 1 = bottom
'Position refers to the aligning
'   0=left,1=centre,2=right,3=inside,4=outside

    On Error Resume Next

    InsertWordPageNumbers = objWordBasic.InsertPageNumbers(intType, intPosition, intFirstPage)

    On Error GoTo 0

End Function

Sub InsertWordText(strNewText As String)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will insert the strNewText at the Current cursor
'location within the current Word document.
    On Error Resume Next

    objWordBasic.Insert (strNewText)

    On Error GoTo 0


End Sub

Sub KillWordObject()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will close word and free all resouces, and make the
'object null, so it must be Set again to use it.

    On Error Resume Next

    Set objWordBasic = Nothing

    On Error GoTo 0


End Sub

Sub MaximizeWord()
'Sub by Mark Matthews.
'This will maximize the application:

    On Error Resume Next

    objWordBasic.AppMaximize

    On Error GoTo 0

End Sub

Sub MinimizeWord()
'Sub by Mark Matthews.
'This will minimize the application:

    On Error Resume Next

    objWordBasic.AppMinimize

    On Error GoTo 0


End Sub

Sub MoveDownWordLines(intLineCount As Integer)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This moves DOWN the specified number of lines in the open
'Word document

    On Error Resume Next

    objWordBasic.LineDown (intLineCount)

    On Error GoTo 0


End Sub

Sub MoveLeftWordCharacters(intCharCount As Integer)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This moves Left the specified number of characters
'in the open Word document

    On Error Resume Next

    objWordBasic.Charleft (intCharCount)

    On Error GoTo 0

End Sub

Sub MoveRightWordCharacters(intCharCount As Integer)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This moves Right the specified number of characters
'in the open Word document

    On Error Resume Next

    objWordBasic.CharRight (intCharCount)

    On Error GoTo 0

End Sub

Sub MoveUpWordLines(intLineCount As Integer)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This moves up the specified number of lines in the open
'Word document

    On Error Resume Next

    objWordBasic.LineUp (intLineCount)

    On Error GoTo 0

End Sub

Sub MoveWordPages(intPages As Integer)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will move the specified number of document pages
' +ve numbers take you further down the document(->End)
' -ve numbers take you towards the begining.

Dim strPages
    
    On Error Resume Next

    strPages = Str$(intPages)
    objWordBasic.EditGoTo (strPages)

    On Error GoTo 0

'EditGoTo.Destination = ""      this should be next Page
'EditGoTo.Destination = "-"     this should be prev page


End Sub

Sub NewDefaultWordFile()
    'routine to open a new default word file.
    
    On Error Resume Next

    objWordBasic.FileNewDefault
 
    On Error GoTo 0

End Sub

Function NewWordFile(strTemplateName As String) As Integer
    'Routine to open a new word file based on a template.
    
Dim intNWFResult As Integer

    'On Error GoTo NewWordFile_ErrorHandler
    On Error Resume Next

    strTemplateName = Trim$(strTemplateName)
    'Debug.Print "-----NewWordFile----------"
    'Debug.Print "strTemplateName >" & strTemplateName
    'Debug.Print "Before: intNWFResult= " & intNWFResult

    intNWFResult = objWordBasic.FileNew(strTemplateName, 0)

    'Debug.Print "After : intNWFResult= " & intNWFResult
    NewWordFile = intNWFResult
     
    On Error GoTo 0
    
    Exit Function

'The exact macro code from word is:
'FileNew .Template = "Letter1", .NewTemplate = 0
'FileNew .Template = "p:\trad\working\blnk002.dot", .NewTemplate = 0

NewWordFile_ErrorHandler:
    MsgBox "Error in Opening New Word File from Template; " & Chr$(10) & Err & ": " & Error$(Err)
    Resume Next

End Function

Sub OpenWordFile(strFileName As String)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will open a Word file with the default settings.
'In full this is:
' object.fileopen("filename",0,0,0,"","",0,"","")
'-but the above full version invariably return errors.

'NOTE:FILENAME IS FULL DIRECTORY PATH
    
    On Error Resume Next

    objWordBasic.FileOpen (strFileName)

    On Error GoTo 0

End Sub

Sub PasteWordFromClipboard()
'this routine will paste all of the clipboard text into
'a word document:

    On Error Resume Next

    objWordBasic.EditPaste

    On Error GoTo 0

'the full macro code for this statement was:
'EditPaste

End Sub

Sub PrintWordFilD_Details(strPages As String, strNumber As String, strFileName As String)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'FilePrint .AppendPrFile = 0, .Range = "4", .PrToFileName = "",
'   .From = "", .To = "", .Type = 0, .NumCopies = "2",
'   .Pages = "1,3-4", .Order = 0, .PrintToFile = 0, .Collate = 1,
'   .FileName = ""
Dim intResult As Integer
    On Error Resume Next

    intResult = objWordBasic.fileprint(0, "4", "", "", "", 0, strNumber, strPages, 0, 0, 1, strFileName)

    On Error GoTo 0

End Sub

Sub PrintWordFile()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will print, to the default Word printer, the
'whole of the currently open Word document.
'The Literal Word macro is:
'FilePrint .AppendPrFile = 0, .Range = "4", .PrToFileName = "",
'   .From = "", .To = "", .Type = 0, .NumCopies = "1",
'   .Pages = "1", .Order = 0, .PrintToFile = 0, .Collate = 1,
'   .FileName = ""
    
    On Error Resume Next

    objWordBasic.fileprint

    On Error GoTo 0

End Sub

Sub SaveALLWordFiles()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will save ALL the currently open word documents to
'their default paths and names: this will invoke SaveAS
'WITHIN word if any are new documents.

    On Error Resume Next

    objWordBasic.FileSaveAll

    On Error GoTo 0

End Sub

Sub SaveAsWordFile(strFileName As String)
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will save the currently open word document to its
'NEW path and name

    On Error Resume Next
    strFileName = Trim$(strFileName)
    objWordBasic.FileSaveAs (strFileName)

    On Error GoTo 0

'FileSaveAs .Name = "TEMP1.DOC", .Format = 0,
'   .LockAnnot = 0, .Password = "", .AddToMru = 1,
'   .WritePassword = "", .RecommendReadOnly = 0,
'   .EmbedFonts = 0, .NativePictureFormat = 0,
'   .FormsData = 0

'FileSave

'FileOpen .Name = "BLNK000.DOC", .ConfirmConversions = 0, .ReadOnly = 0, .AddToMru = 0, .PasswordDoc = "", .PasswordDot = "", .Revert = 0, .WritePasswordDoc = "", .WritePasswordDot = ""

'FileSaveAll

'ChDefaultDir "\\PRIMUS\PROJECTS\TRAD\DAYBOOK\", 0
'FileSaveAs .Name = "DOC24.DOC", .Format = 0, .LockAnnot = 0, .Password = "", .AddToMru = 1, .WritePassword = "", .RecommendReadOnly = 0, .EmbedFonts = 0, .NativePictureFormat = 0, .FormsData = 0
'FileClose
'

End Sub

Sub SaveWordFile()
'---------------------------------------------------------
'   WORD fuctions and calls set up by Mark Matthews
'This will save the currently open word document to its
'default path and name

    On Error Resume Next

    objWordBasic.FileSave

    On Error GoTo 0

'FileSaveAs .Name = "TEMP1.DOC", .Format = 0, .LockAnnot = 0, .Password = "", .AddToMru = 1, .WritePassword = "", .RecommendReadOnly = 0, .EmbedFonts = 0, .NativePictureFormat = 0, .FormsData = 0
'FileSave
'FileOpen .Name = "BLNK000.DOC", .ConfirmConversions = 0, .ReadOnly = 0, .AddToMru = 0, .PasswordDoc = "", .PasswordDot = "", .Revert = 0, .WritePasswordDoc = "", .WritePasswordDot = ""
'FileSaveAll

End Sub

'
Sub SelectAllWordDoc()
'this routine will select all of a word document:

    On Error Resume Next

    objWordBasic.EditSelectAll

    On Error GoTo 0

'the full macro code for this statement was:
'EditSelectAll
End Sub

Function SetWordPrinterBin(intType As Integer) As Integer
'This function will alter the Word Default Printer Tray
'to the following values:
'BINPAPERCASSETTE = 0        = "Paper Cassette"
'BINMANUALFEED = 1           = "Manual Feed"
'BINFROMCONTROLPANEL = 2     = "From Control Panel"

Dim strBin As String
    
    On Error Resume Next
     
    Select Case intType
        Case BINPAPERCASSETTE
            strBin = "Paper Cassette"
        Case BINMANUALFEED
            strBin = "Manual Feed"
        Case BINCONTROLPANEL
            strBin = "From Control Panel"
        Case Else
            Exit Function
    End Select

    SetWordPrinterBin = objWordBasic.ToolsOptionsPrint(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, strBin)

    On Error GoTo 0

'The full macro code was:
'ToolsOptionsPrint .Draft = 0, .Reverse = 0, .UpdateFields = 0,
'.Summary = 0, .ShowCodes = 0, .Annotations = 0, .ShowHidden = 0,
'.EnvFeederInstalled = 0, .WidowControl = 1, .DfltTrueType = 1,
'.UpdateLinks = 0, .Background = 0, .DrawingObjects = 1, .FormsData = 0,
'.DefaultTray = "Paper Cassette"

'ToolsOptionsPrint .Draft = 0, .Reverse = 0, .UpdateFields = 0,
'.Summary = 0, .ShowCodes = 0, .Annotations = 0, .ShowHidden = 0,
'.EnvFeederInstalled = 0, .WidowControl = 1, .DfltTrueType = 1,
'.UpdateLinks = 0, .Background = 0, .DrawingObjects = 1, .FormsData = 0,
'.DefaultTray = "Manual Feed"

'ToolsOptionsPrint .Draft = 0, .Reverse = 0, .UpdateFields = 0,
'.Summary = 0, .ShowCodes = 0, .Annotations = 0, .ShowHidden = 0,
'.EnvFeederInstalled = 0, .WidowControl = 1, .DfltTrueType = 1,
'.UpdateLinks = 0, .Background = 0, .DrawingObjects = 1, .FormsData = 0,
'.DefaultTray = "From Control Panel"

End Function

Sub ToggleWordHeaderFooter()

   On Error Resume Next
   
   objWordBasic.GoToHeaderFooter

   On Error GoTo 0
'ViewHeader
'GoToHeaderFooter
'LineDown 1
'EditClear
'EndOfLine
'EditClear
'EditClear
'EditClear
'EditClear -22
'CloseViewHeaderFooter
End Sub

Sub ViewWordHeaderFooter()

   On Error Resume Next
   
   objWordBasic.ViewHeader

   On Error GoTo 0
'ViewHeader
'GoToHeaderFooter
'LineDown 1
'EditClear
'EndOfLine
'EditClear
'EditClear
'EditClear
'EditClear -22
'CloseViewHeaderFooter

End Sub

Sub zzzprotomacro6()
'Sub MAIN

Rem goto next line:
'EditGoTo.Destination = "l"

'EditGoTo.Destination = "l"
'EditGoTo.Destination = "l"

Rem goto next page:
'EditGoTo.Destination = ""

Rem goto previous page:
'EditGoTo.Destination = "-"

'EditGoTo.Destination = "l"
'EditGoTo.Destination = "l"
'EditGoTo.Destination = "l"

Rem goto a bookmark:by label ONLY
'EditGoTo.Destination = "bookmark_01"
'EditGoTo.Destination = "bookmark_02"
'EditGoTo.Destination = "-"
End Sub

Sub zzzprotomacro7()
'Sub MAIN
'FileSaveAs .Name = "TRAD002.DOC", .Format = 0,
'    .LockAnnot = 0, .Password = "", .AddToMru = 1,
'    .WritePassword = "", .RecommendReadOnly = 0,
'    .EmbedFonts = 0, .NativePictureFormat = 0,
'    .FormsData = 0
'FileSave
'End Sub
Rem so use FileSaveAs("docname",0,0,"",1,"",0,0,0,0)
End Sub

Sub zzzprotomacro8()
Rem change default directory
'ChDefaultDir "C:\PROTOTYP\TRADDOTS\", 0
Rem to open a file:
'FileOpen .Name = "BLNK003.DOC", .ConfirmConversions = 0, .ReadOnly = 0,
'       .AddToMru = 0, .PasswordDoc = "", .PasswordDot = "",
'       .Revert = 0, .WritePasswordDoc = "", .WritePasswordDot = ""

Rem object.fileopen("filename",0,0,0,"","",0,"","")


'VLine 5
'LineDown 7
'LineUp 1
'EditBookmark .Name = "bookmark_1", .SortBy = 0, .Add
'LineDown 1
'EditBookmark .Name = "bookmark_2", .SortBy = 0, .Add
'LineDown 5
'CharRight 5
'CharLeft 1
'Insert " "
'EditBookmark .Name = "bookmark_3", .SortBy = 0, .Add
'VLine 35
'FileClose

End Sub

