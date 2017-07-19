Attribute VB_Name = "modStringLib"
'Created by Robin G Brown
'-----------------------------------------------------------------------------
'   This module contains code for various string handling functions
'-----------------------------------------------------------------------------
'   IMPORTANT!
'   Certain string functions that deal with CSV data have been transferred
'   to the CSV handling library file
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "StringHandlingCode"
Private lngCounter As Long
Private lngReturn As Long
'Constants for EncryptDecrypt
Global Const conEncrypt = 0
Global Const conDecrypt = 1
'Constants for ExtractEncodedData
Global Const conEncodeNumber = 1
Global Const conEncodeString = 2
'---Spoiler Information for CSV String Handling - RGB
'      String
'       AB,CDE,FG
'       123456789
'         x   y z
'      z = Len(String)
'      x = Instr(String, ",")
'      y = Instr(x + 1, String, ",")
'      Len("CDE") = y - x - 1
'      "AB" = Left$(String, x - 1)
'      "FG" = Right$(String, z - y)
'      "CDE" = Mid$(String, x + 1, y - x - 1)
'---End of Spoiler Information for String Handling
'
'Public Function ImmediateErrorStringFunction() As String
''Created by Robin G Brown, INSERT_DATE
''DESCRIPTION_OF_FUNCTIONALITY_HERE
''Function Declares
'Const conFunction = "ImmediateErrorStringFunction"
'    'Error Trap
'    On Error GoTo ImmediateErrorStringFunction_ErrorHandler
'    'CODE_HERE
'    ImmediateErrorStringFunction = "ImmediateErrorStringFunction"
'Exit Function
''Error Handler
'ImmediateErrorStringFunction_ErrorHandler:
'    Select Case Err.Number
'    'Case ERROR_CODE_HERE
'        'ERROR_HANDLING_CODE_HERE
'    Case Else
'        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
'    End Select
'    ImmediateErrorStringFunction = ""
'    Exit Function
'End Function

Public Function StripQuotes(ByVal strDirty As String) As String
'Created by Robin G Brown, 8/9/97
'Strip out any " or ' from a string, replacing them with spaces
'Function Declares
Const conFunction = "StripQuotes"
Dim strClean As String
Dim strChar As String * 1
    'Error Trap
    On Error GoTo StripQuotes_ErrorHandler
    strClean = ""
    For lngCounter = 1 To Len(strDirty)
        strChar = Mid$(strDirty, lngCounter, 1)
        Select Case Asc(strChar)
        Case 34, 39
            strClean = strClean & " "
        Case Else
            strClean = strClean & strChar
        End Select
    Next
    StripQuotes = strClean
Exit Function
'Error Handler
StripQuotes_ErrorHandler:
    StripQuotes = ""
    Exit Function
End Function

Public Function ReplaceStringInString(ByVal strSourceField As String, ByVal strSearch As String, ByVal strReplace As String) As String
'Created by Robin G Brown, 1/9/97
'Replace all occurences of strSearch in strSourceField with strReplace
'Function Declares
Const conFunction = "ReplaceStringInString"
Dim lngStartPos As Long
Dim lngEndPos As Long
Dim strFirstBit As String
Dim strLastBit As String
    'Error Trap
    On Error GoTo ReplaceStringInString_ErrorHandler
    If strSearch <> strReplace Then
        'Find first occurence
        lngStartPos = InStr(strSourceField, strSearch)
        Do While lngStartPos > 0
            'Replace current occurence
            lngEndPos = lngStartPos + Len(strSearch)
            'Cut out pre-replace bit
            strFirstBit = Left$(strSourceField, lngStartPos - 1)
            'Cut out post-replace bit
            strLastBit = Right$(strSourceField, Len(strSourceField) - lngEndPos + 1)
            'Construct changed string
            strSourceField = strFirstBit & strReplace & strLastBit
            'Find next occurence
            lngStartPos = InStr(strSourceField, strSearch)
        Loop
    End If
    ReplaceStringInString = strSourceField
Exit Function
'Error Handler
ReplaceStringInString_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ReplaceStringInString = ""
    Exit Function
End Function

Public Function ConvertStringToDouble(ByVal strValue As String) As Double
'Created by Robin G Brown, 19/8/97
'Take strValue and make sure it ends up as a double
'Function Declares
Const conFunction = "ConvertStringToDouble"
    'Error Trap
    On Error GoTo ConvertStringToDouble_ErrorHandler
    If IsNumeric(strValue) Then
        ConvertStringToDouble = CDbl(strValue)
        Exit Function
    Else
        If InStr(strValue, "%") > 0 Then
            ConvertStringToDouble = CDbl(Val(strValue))
            Exit Function
        End If
    End If
Exit Function
'Error Handler
ConvertStringToDouble_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        Err.Raise Err.Number, conModule & ":" & conFunction, "Unexpected error : " & Err.Description
    End Select
    ConvertStringToDouble = 0#
    Exit Function
End Function

Public Function ReplaceCharacters(ByVal strSourceField As String, ByVal strSearch As String, ByVal strReplace As String) As String
'Created by Robin G Brown, 6/8/97
'Read strSourceField and replace each occurence of strSearch with strReplace, _
    strReplace must be larger then or equal in size to strSearch, _
    otherwise not all characters in strSourceField will be replaced
'Function Declares
Const conFunction = "ReplaceCharacters"
Dim lngXCharPos As Long
Dim lngZLen As Long
Dim lngSearchLen As Long
    'Error Trap
    On Error GoTo ReplaceCharacters_ErrorHandler
    lngSearchLen = Len(strSearch)
    If Not IsNull(strSourceField) Then
        lngXCharPos = InStr(1, strSourceField, strSearch)
        Do While lngXCharPos > 0
            Mid(strSourceField, lngXCharPos, lngSearchLen) = strReplace
            lngXCharPos = InStr(lngXCharPos + 1, strSourceField, strSearch)
        Loop
    End If
    ReplaceCharacters = strSourceField
Exit Function
'Error Handler
ReplaceCharacters_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ReplaceCharacters = strSourceField
    Exit Function
End Function

Public Function IsNumberString(ByVal strToCheck As String) As Boolean
'Created by Robin G Brown
'Reads a string, returns true if it is a valid simple number, false otherwise
'Nasty characters: see below.
'Function Declares
Dim lngCharPos As Long
    'Error Trap
    On Error Resume Next
    lngCharPos = 1
    If Not IsNull(strToCheck) Then
        Do While lngCharPos <= Len(strToCheck)
            'First of all, use only a single character to check
            Select Case Mid$(strToCheck, lngCharPos, 1)
            Case ""
                IsNumberString = True
            Case "0" To "9"
                IsNumberString = True
            Case ".", "+", "-", ","
                'Some punctuation is allowed in numbers
                IsNumberString = True
            Case "a" To "z"
                IsNumberString = False
                Exit Function
            Case "A" To "Z"
                IsNumberString = False
                Exit Function
            Case "_", "-"
                IsNumberString = False
                Exit Function
            Case Else
                'Includes " ", :,;'#~/\?|!£$%^&*()+=¬`@[]{}, etc.
                IsNumberString = False
                Exit Function
            End Select
            lngCharPos = lngCharPos + 1
        Loop
    End If
End Function

Public Function IsDigitString(ByVal strToCheck As String) As Boolean
'Created by Robin G Brown
'Reads a string, returns true if it consists only of numbers 0-9, false otherwise
'Nasty characters: see below.
'Function Declares
Dim lngCharPos As Long
    'Error Trap
    On Error Resume Next
    lngCharPos = 1
    If Not IsNull(strToCheck) Then
        Do While lngCharPos <= Len(strToCheck)
            'First of all, use only a single character to check
            Select Case Mid$(strToCheck, lngCharPos, 1)
            Case ""
                IsDigitString = True
            Case "0" To "9"
                IsDigitString = True
            Case Else
                'Includes " ", :,;'#~/\?|!£$%^&*()+=¬`@[]{}, etc.
                IsDigitString = False
                Exit Function
            End Select
            lngCharPos = lngCharPos + 1
        Loop
    End If
End Function

Public Function TruncateLeft(ByVal strTruncate As String, ByVal lngDigits As Long) As String
'Created by Robin G Brown, 11/7/97
'Chop off a number of characters on the left side of strTruncate
'Function Declares
Const conFunction = "TruncateLeft"
    'Error Trap
    On Error GoTo TruncateLeft_ErrorHandler
    If lngDigits > Len(strTruncate) Then
        TruncateLeft = ""
    Else
        TruncateLeft = Right$(strTruncate, Len(strTruncate) - lngDigits)
    End If
Exit Function
'Error Handler
TruncateLeft_ErrorHandler:
    TruncateLeft = ""
    Exit Function
End Function

Public Function TruncateRight(ByVal strTruncate As String, ByVal lngDigits As Long) As String
'Created by Robin G Brown, 11/7/97
'Chop off a number of characters on the right side of strTruncate
'Function Declares
Const conFunction = "TruncateRight"
    'Error Trap
    On Error GoTo TruncateRight_ErrorHandler
    If lngDigits > Len(strTruncate) Then
        TruncateRight = ""
    Else
        TruncateRight = Left$(strTruncate, Len(strTruncate) - lngDigits)
    End If
Exit Function
'Error Handler
TruncateRight_ErrorHandler:
    TruncateRight = ""
    Exit Function
End Function

Public Function TruncateEnds(ByVal strTruncate As String, ByVal lngDigits As Long) As String
'Created by Robin G Brown, 11/7/97
'Chop off a number of characters on either side of strTruncate
'Function Declares
Const conFunction = "TruncateEnds"
    'Error Trap
    On Error GoTo TruncateEnds_ErrorHandler
    If lngDigits > (Len(strTruncate) / 2) Then
        TruncateEnds = ""
    Else
        TruncateEnds = Mid$(strTruncate, lngDigits + 1, Len(strTruncate) - (lngDigits * 2))
    End If
Exit Function
'Error Handler
TruncateEnds_ErrorHandler:
    TruncateEnds = ""
    Exit Function
End Function

Public Function ExtractEncodedData(ByVal strTextData As String, ByVal lngStart As Long, _
    ByVal lngLength As Long, ByVal lngDataType As Long) As String
'Created by Robin G Brown, 7/7/97
'Extract text from a string and convert to an encoded data type, _
    returns "" on error
'Function Declares
Const conFunction = "ExtractEncodedData"
Dim strExtract As String
    'Error Trap
    On Error GoTo ExtractEncodedData_ErrorHandler
    strExtract = Mid$(strTextData, lngStart, lngLength)
    Select Case lngDataType
    Case conEncodeNumber
        'Return a number
        ExtractEncodedData = strExtract
    Case Else
        'Return a string
        ExtractEncodedData = Chr$(34) & strExtract & Chr$(34)
    End Select
Exit Function
'Error Handler
ExtractEncodedData_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    ExtractEncodedData = ""
    Exit Function
End Function

Public Function AlphaFirstString(ByVal strAlphaOne As String, ByVal strAlphaTwo As String) As String
'Created by Robin G Brown, 27/5/97
'Return the string which is alphabetically before the other
'Function Declares
Const conFunction = "AlphaFirstString"
    'Error Trap
    On Error GoTo AlphaFirstString_ErrorHandler
    Select Case StrComp(strAlphaOne, strAlphaTwo, vbBinaryCompare)
    Case -1
        AlphaFirstString = strAlphaOne
    Case 0
        AlphaFirstString = strAlphaOne
    Case 1
        AlphaFirstString = strAlphaTwo
    Case Else
        AlphaFirstString = Null
    End Select
Exit Function
'Error Handler
AlphaFirstString_ErrorHandler:
    AlphaFirstString = Null
    Exit Function
End Function

Public Function AlphaLastString(ByVal strAlphaOne As String, ByVal strAlphaTwo As String) As String
'Created by Robin G Brown, 27/5/97
'Return the string which is alphabetically after the other
'Function Declares
Const conFunction = "AlphaLastString"
    'Error Trap
    On Error GoTo AlphaLastString_ErrorHandler
    Select Case StrComp(strAlphaOne, strAlphaTwo, vbBinaryCompare)
    Case -1
        AlphaLastString = strAlphaTwo
    Case 0
        AlphaLastString = strAlphaOne
    Case 1
        AlphaLastString = strAlphaOne
    Case Else
        AlphaLastString = Null
    End Select
Exit Function
'Error Handler
AlphaLastString_ErrorHandler:
    AlphaLastString = Null
    Exit Function
End Function

Public Function GetNthWord(ByVal strSourceField As String, ByVal lngNthWord As Long) As String
'Created by Robin G Brown, 20/5/97
'Get the Nth word from a string of words seperated by whitespace
'Function Declares
Const conFunction = "GetNthWord"
Dim lngXPos As Long
Dim lngYPos As Long
Dim lngCurrentWord As Long
    'Error Trap
    On Error GoTo GetNthWord_ErrorHandler
    'Set lngXPos to skip over whitespace at start of strSourceField
    lngXPos = FindNextCharacter(strSourceField, 1) - 1
    lngCurrentWord = 1
    Do While lngCurrentWord < lngNthWord
        'Set lngXPos
        lngXPos = FindWhiteSpace(strSourceField, lngXPos + 1)
        lngXPos = FindNextCharacter(strSourceField, lngXPos + 1) - 1
        lngCurrentWord = lngCurrentWord + 1
    Loop
    'Set lngYPos
    lngYPos = FindWhiteSpace(strSourceField, lngXPos + 1)
    If lngYPos > 0 Then
        GetNthWord = Mid$(strSourceField, lngXPos + 1, lngYPos - lngXPos - 1)
    Else
        GetNthWord = Right$(strSourceField, Len(strSourceField) - lngXPos)
    End If
Exit Function
'Error Handler
GetNthWord_ErrorHandler:
    GetNthWord = ""
    Exit Function
End Function

Public Function CharactersInString(ByVal strStringToSearch As String, ByVal strCharacter As String) As Long
'Created by Robin G Brown, 20/5/97
'Returns the number of selected characters in a string
'Function Declares
Const conFunction = "CharactersInString"
Dim strChar As String * 1
Dim lngSearchPos As String
Dim lngCharsFound As Long
    'Error Trap
    On Error GoTo CharactersInString_ErrorHandler
    strChar = strCharacter
    lngCharsFound = 0
    lngSearchPos = InStr(1, strStringToSearch, strChar)
    Do While lngSearchPos > 0
        lngCharsFound = lngCharsFound + 1
        lngSearchPos = InStr(lngSearchPos + 1, strStringToSearch, strChar)
    Loop
    CharactersInString = lngCharsFound
Exit Function
'Error Handler
CharactersInString_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        'MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conFunction
    End Select
    CharactersInString = -1
    Exit Function
End Function

Public Function StripCRLF(ByVal strDirty As String) As String
'Created by Robin G Brown
'Strip out any CRs or LFs from a string, replacing them with spaces
'Function Declares
Const conFunction = "StripCRLF"
Dim strClean As String
Dim strChar As String * 1
    'Error Trap
    On Error GoTo StripCRLF_ErrorHandler
    strClean = ""
    For lngCounter = 1 To Len(strDirty)
        strChar = Mid$(strDirty, lngCounter, 1)
        Select Case Asc(strChar)
        Case 10, 13
            strClean = strClean & " "
        Case Else
            strClean = strClean & strChar
        End Select
    Next
    StripCRLF = strClean
Exit Function
'Error Handler
StripCRLF_ErrorHandler:
    StripCRLF = ""
    Exit Function
End Function

Public Function CheckArrayForStrings(ByRef strArray() As String) As Boolean
'Created by Robin G Brown, 18/7/97
'Check an array to see if there is any string data in it
'Function Declares
Const conFunction = "CheckArrayForStrings"
Dim lngArraycounter As Long
    'Error Trap
    On Error GoTo CheckArrayForStrings_ErrorHandler
    For lngArraycounter = LBound(strArray, 1) To UBound(strArray, 1)
        If strArray(lngArraycounter) <> "" Then
            CheckArrayForStrings = True
            Exit Function
        End If
    Next
    CheckArrayForStrings = False
Exit Function
'Error Handler
CheckArrayForStrings_ErrorHandler:
    CheckArrayForStrings = False
    Exit Function
End Function

Public Function IsCleanString(ByVal strToCheck As String) As Boolean
'Created by Robin G Brown
'Reads strToCheck, if there are no 'nasty' characters in it, returns true
'Nasty characters: see below.
'Function Declares
Dim lngCharPos As Long
    'Error Trap
    On Error Resume Next
    lngCharPos = 1
    If Not IsNull(strToCheck) Then
        Do While lngCharPos <= Len(strToCheck)
            'First of all, use only a single character to check
            Select Case Mid$(strToCheck, lngCharPos, 1)
            Case ""
                IsCleanString = True
            Case "0" To "9"
                IsCleanString = True
            Case "a" To "z"
                IsCleanString = True
            Case "A" To "Z"
                IsCleanString = True
            Case "_", "-"
                IsCleanString = True
            Case Else
                'Includes " ", :,;'#~/\?|!£$%^&*()+=¬`@[]{}, etc.
                IsCleanString = False
                Exit Function
            End Select
            lngCharPos = lngCharPos + 1
        Loop
    End If
End Function

Public Function WordMatch(ByVal strSearch As String, ByVal strFind As String) As Long
'Created by Robin G Brown, 18/2/97
'Searches for exact words in strings _
    i.e. searches strSearch for strFind, strFind must be surrounded by whitespace _
    and must match exactly.
'E.G. _
    MsgBox WordMatch("here little fish thing", "fish") _
    will display 13 _
    but MsgBox WordMatch("here little fishthing", "fish") _
    will display 0 _
    and MsgBox WordMatch("here little fishy fish thing", "fish") _
    will display 19
'Function Declares
Const conFunction = "WordMatch"
Dim strSubSearch As String
Dim lngLenFind As Long
Dim lngSearchPos As Long
Dim strCharBefore As String * 1
Dim strCharAfter As String * 1
    'Error Trap
    On Error GoTo WordMatch_ErrorHandler
    lngLenFind = Len(strFind)
    lngSearchPos = 1
    Do While lngSearchPos < Len(strSearch)
        strSubSearch = Mid$(strSearch, lngSearchPos, lngLenFind)
        If StrComp(strSubSearch, strFind) = 0 Then
            'Check for whitespace and check for beginning of strSearch
            If lngSearchPos = 1 Then
                strCharBefore = " "
            Else
                strCharBefore = Mid$(strSearch, lngSearchPos - 1, 1)
            End If
            strCharAfter = Mid$(strSearch, lngSearchPos + lngLenFind, 1)
            If IsWhiteSpace(strCharBefore) And IsWhiteSpace(strCharAfter) Then
                'Found whitespace on both sides, so get out of here!
                WordMatch = lngSearchPos
                Exit Function
            End If
        End If
        lngSearchPos = lngSearchPos + 1
    Loop
    'Match not found so return 0
    WordMatch = 0
On Error GoTo 0
Exit Function
'Error Handler
WordMatch_ErrorHandler:
    On Error Resume Next
    'Return 0 on error as per InStr
    WordMatch = 0
    Exit Function
End Function

Public Function IsWhiteSpace(ByVal strToCheck As String) As Boolean
'Created by Robin G Brown, 18/2/97
'Reads input, if it is entirely whitespace, returns true
'Whitespace : Anything that is not either a letter or a number, null or zero length _
    thus all punctuation, etc. is included.
'Function Declares
    'Error Trap
    On Error Resume Next
    IsWhiteSpace = False
    If Not IsNull(strToCheck) Then
        'First of all, use only a single character to check
        Select Case Left$(strToCheck, 1)
        Case "0" To "9"
            IsWhiteSpace = False
        Case "a" To "z"
            IsWhiteSpace = False
        Case "A" To "Z"
            IsWhiteSpace = False
        Case ""
            IsWhiteSpace = False
        Case "_", "-"
            'Underscores and dashes may not usually be considered as whitespace
            IsWhiteSpace = False
        Case "\", "/"
            'Slashes may not usually be considered as whitespace
            IsWhiteSpace = False
        Case Else
            IsWhiteSpace = True
        End Select
    End If
    On Error GoTo 0
End Function

Public Function GetLastCharPos(ByVal strSourceField As String, ByVal strCharacter As String) As Long
'Created by Robin G Brown
'Takes a string and finds the position of the last given character in it, _
    will return 0 if character is not found or if more than one character is entered _
    or if there is an error
'Function Declares
Const conFunction = "GetLastCharPos"
Dim lngAnother As Long
Dim lngStart As Long
Dim lngLast As Long
    'Error Trap
    On Error GoTo GetLastCharPos_ErrorHandler
    GetLastCharPos = 0
    If Len(strCharacter) > 1 Then
        Exit Function
    End If
    'find first position
    lngStart = 1
    lngAnother = InStr(lngStart, strSourceField, strCharacter, 1)
    lngLast = lngAnother
    Do While lngAnother >= lngStart
        lngLast = lngAnother
        lngStart = lngAnother + 1
        lngAnother = InStr(lngStart, strSourceField, strCharacter, 1)
    Loop
    GetLastCharPos = lngLast
Exit Function
'Error Handler
GetLastCharPos_ErrorHandler:
    GetLastCharPos = 0
    Exit Function
End Function

Public Function GetLastSlashPos(ByVal strSourceField As String) As Long
'Created by Robin G Brown
'Takes a string and finds the position of the last '/' in it
'Function Declares
Const conFunction = "GetLastSlashPos"
    'Error Trap
    On Error GoTo GetLastSlashPos_ErrorHandler
    GetLastSlashPos = GetLastCharPos(strSourceField, "\")
Exit Function
'Error Handler
GetLastSlashPos_ErrorHandler:
    GetLastSlashPos = 0
    Exit Function
End Function

Public Function FindWhiteSpace(ByVal strSearch As String, Optional ByVal varStartPosition As Variant) As Long
'Created by Robin G Brown
'Returns the position of the next 'whitespace' character
Dim lngCharCounter As Long
Dim lngChar As String * 1
    On Error GoTo FindWhiteSpace_ErrorHandler
    If IsMissing(varStartPosition) Then
        varStartPosition = 1
    End If
    If VarType(varStartPosition) <> vbLong Then
        varStartPosition = 1
    End If
    For lngCharCounter = varStartPosition To Len(strSearch)
        lngChar = Mid$(strSearch, lngCharCounter, 1)
        If IsWhiteSpace(lngChar) = True Then
            FindWhiteSpace = lngCharCounter
            Exit Function
        End If
    Next
FindWhiteSpace_ErrorHandler:
    FindWhiteSpace = 0
    Exit Function
End Function

Public Function FindNextCharacter(ByVal strSearch As String, Optional ByVal varStartPosition As Variant) As Long
'Created by Robin G Brown
'Returns the position of the next 'non-whitespace' character
Dim lngCharCounter As Long
Dim lngChar As String * 1
    On Error GoTo FindNextCharacter_ErrorHandler
    If IsMissing(varStartPosition) Then
        varStartPosition = 1
    End If
    If VarType(varStartPosition) <> vbLong Then
        varStartPosition = 1
    End If
    For lngCharCounter = varStartPosition To Len(strSearch)
        lngChar = Mid$(strSearch, lngCharCounter, 1)
        If IsWhiteSpace(lngChar) = False Then
            FindNextCharacter = lngCharCounter
            Exit Function
        End If
    Next
FindNextCharacter_ErrorHandler:
    FindNextCharacter = 0
    Exit Function
End Function

Public Function IsChar(ByVal strTest As String) As Long
'Created by Robin G Brown, 18/4/95
'Returns true if first character of test string is an alpha character
On Error GoTo IsChar_ErrorHandler
    strTest = Left$(strTest, 1)
    'A-Z=65-90, a-z=97-122
    If (Asc(strTest) > 64 And Asc(strTest) < 91) _
    Or (Asc(strTest) > 96 And Asc(strTest) < 123) _
    Then
        IsChar = True
    End If
Exit Function
IsChar_ErrorHandler:
    IsChar = False
    Exit Function
End Function

Public Function ProfileToString(ByVal strTarget As String) As String
'Created by Robin G Brown, 18/4/95
'Changes fixed length string to variable length string
Dim strReturn As String
Dim strExchange As String
Dim lngCounter As Long
    On Error GoTo ProfileToString_ErrorHandler
    strReturn = ""
    lngCounter = 1
    Do
        strExchange = Mid$(strTarget, lngCounter, 1)
        If Asc(strExchange) < 1 Or Asc(strExchange) > 255 Then Exit Do
        strReturn = strReturn & strExchange
        lngCounter = lngCounter + 1
    Loop
    ProfileToString = strReturn
Exit Function
ProfileToString_ErrorHandler:
    ProfileToString = ""
    Exit Function
End Function

Public Function FilterCharacterKey(ByVal KeyAsciiCode As Long) As Long
'Created by Robin G Brown, 15/11/95
'This function takes an ASCII code and returns either the ASCII code or 0 _
    depending on wither the ASCII code is for a character, digit or punctuation
'Return value is set to 0 if this code should not be present _
    in a 'numbers only' textbox, string, etc.
'This function can be copied and modified to allow for filtering of other characters - RGB/15/4/97
    On Error Resume Next
    Select Case KeyAsciiCode
    Case 8, 9, 10, 32
        'Backspace, tab, linefeed and space
        'These should never be altered
        FilterCharacterKey = KeyAsciiCode
    Case 33, 38, 40 To 43, 45 To 47, 60, 61, 62, 124
        'Mathematical and logical stuff
        FilterCharacterKey = KeyAsciiCode
    Case 34 To 39, 44, 58, 59, 63, 64, 91 To 96, 123, 125, 126
        'Punctuation stuff
        FilterCharacterKey = 0
    Case 48 To 57
        'Numbers
        FilterCharacterKey = KeyAsciiCode
    Case 65 To 90, 97 To 122
        'lowercase and uppercase characters
        FilterCharacterKey = 0
    Case Else
        'all unsupported characters and 'unusual ones'
        FilterCharacterKey = 0
    End Select
    On Error GoTo 0
End Function

Public Function EncryptDecrypt(ByVal strTarget As String, ByVal lngAction As Long) As String
'Created by Robin G Brown
'Encrypts or Decrypts a string using a simple encryption algorithm
'strTarget is plain or Encrypted string, _
    not too long!
'lngAction conEncrypt = 0, conDecrypt = 1
'Sub Declares
Const conSub = "EncryptDecrypt"
    'Error Trap
Dim strReturn As String
Dim strExchange As String
Dim lngCounter As Long
    On Error GoTo EncryptDecrypt_ErrorHandler
    strReturn = ""
    If lngAction = 0 Then
        'conEncrypt Routine
        For lngCounter = 1 To Len(strTarget)
            strExchange = Mid$(strTarget, lngCounter, 1)
            strReturn = strReturn & Chr$(Asc(strExchange) + lngCounter)
            'strReturn = strReturn & Chr$((Asc(strExchange) + lngCounter) Mod 26)
        Next
    Else
        'conDecrypt Routine
        For lngCounter = 1 To Len(strTarget)
            strExchange = Mid$(strTarget, lngCounter, 1)
            strReturn = strReturn & Chr$(Asc(strExchange) - lngCounter)
            'strReturn = strReturn & Chr$((Asc(strExchange) - lngCounter) Mod 26)
        Next
    End If
EncryptDecrypt_ErrorHandler:
    EncryptDecrypt = strReturn
    Exit Function
End Function

