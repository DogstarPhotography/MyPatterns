VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4185
   ClientLeft      =   7770
   ClientTop       =   1455
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4185
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAction 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   3300
      TabIndex        =   2
      Top             =   3660
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   3660
      Width           =   1455
   End
   Begin TabDlg.SSTab tabOptions 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   6059
      _Version        =   327680
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Default Behaviour"
      TabPicture(0)   =   "Options.frx":0000
      Tab(0).ControlCount=   2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOptions"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraPresentation"
      Tab(0).Control(1).Enabled=   0   'False
      TabCaption(1)   =   "AutoArrange"
      TabPicture(1)   =   "Options.frx":001C
      Tab(1).ControlCount=   7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picWindow(1)"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "picWindow(2)"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "picWindow(3)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "picWindow(4)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "picWindow(5)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "picWindow(6)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "chkOption(1)"
      Tab(1).Control(6).Enabled=   -1  'True
      Begin VB.Frame fraPresentation 
         Caption         =   "Presentation"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   4395
         Begin VB.OptionButton optTheme 
            Caption         =   "Theme 4"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton optTheme 
            Caption         =   "Theme 3"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   900
            Width           =   2055
         End
         Begin VB.OptionButton optTheme 
            Caption         =   "Theme 2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton optTheme 
            Caption         =   "Theme 1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Auto Arrange"
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   11
         Top             =   2940
         Width           =   1635
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":0038
         Height          =   2475
         Index           =   6
         Left            =   -71580
         ScaleHeight     =   2415
         ScaleWidth      =   975
         TabIndex        =   10
         Top             =   420
         Width           =   1035
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":047A
         Height          =   1215
         Index           =   5
         Left            =   -72720
         ScaleHeight     =   1155
         ScaleWidth      =   975
         TabIndex        =   9
         Top             =   1680
         Width           =   1035
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":08BC
         Height          =   1215
         Index           =   4
         Left            =   -72720
         ScaleHeight     =   1155
         ScaleWidth      =   975
         TabIndex        =   8
         Top             =   420
         Width           =   1035
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":0CFE
         Height          =   1215
         Index           =   3
         Left            =   -73800
         ScaleHeight     =   1155
         ScaleWidth      =   975
         TabIndex        =   7
         Top             =   1680
         Width           =   1035
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":1140
         Height          =   1215
         Index           =   2
         Left            =   -73800
         ScaleHeight     =   1155
         ScaleWidth      =   975
         TabIndex        =   6
         Top             =   420
         Width           =   1035
      End
      Begin VB.PictureBox picWindow 
         DragIcon        =   "Options.frx":1582
         Height          =   2475
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   2415
         ScaleWidth      =   975
         TabIndex        =   5
         Top             =   420
         Width           =   1035
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4395
         Begin VB.CheckBox chkOption 
            Caption         =   "Show Scratchpad"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1995
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Robin G Brown, 28th April 1997
'-----------------------------------------------------------------------------
'   This form contains code for _
    showing options for the behaviour of the program
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---All Declares
Private Const conModule = "frmOptions"
Private intCounter As Integer
Private intReturn As Integer
'Declare constants for boilerplate code
Private Const conCentering = True    'causes form to center on screen when shown, check form_resize event

Private Sub chkOption_Click(Index As Integer)
'Created by Robin G Brown, 29/4/97
'Enable/Disable the Auto Arrange tab depending on the setting of checkbox(1)
    'Error trap
    On Error Resume Next
    Select Case Index
    Case 1
        If chkOption(1).Value = 1 Then
            For intCounter = 1 To 6
                picWindow(intCounter).Enabled = True
            Next
        Else
            For intCounter = 1 To 6
                picWindow(intCounter).Enabled = False
            Next
        End If
    Case Else
        'Do nothing
    End Select
End Sub

Private Sub cmdAction_Click(Index As Integer)
'Created by Robin G Brown, 28/4/97
'respond to the Ok or Cancel actions
'Sub Declares
Const conSub = "ImmediateErrorSub"
    'Error Trap
    On Error GoTo ImmediateErrorSub_ErrorHandler
    Select Case Index
    Case 0
        'OK
        'Carry out any actions
        'Set the frmPosition array up, checking that the layout is OK first
        If picWindow(1).Tag <> "empty" And picWindow(6).Tag <> "empty" Then
            MsgBox "One of the sides on the Auto Arrange layout must be empty, please correct."
            Exit Sub
        End If
        'Checkboxes
        If chkOption(0).Value = 1 Then
            frmMaster.mnuWindowScratchpad.Checked = True
        Else
            frmMaster.mnuWindowScratchpad.Checked = False
        End If
        'Resize the scratchpad, now that the option has been set
        Call mdiTreeChild.ArrangeSourceAndScratch
        'Theme options
        For intCounter = 1 To 4
            If optTheme(intCounter).Value = True Then
                Call ChangeTheme(intCounter)
                Exit For
            End If
        Next
        'Sort out the autoarrange
        If chkOption(1).Value = 1 Then
            frmMaster.mnuWindowAutoArrange.Checked = True
            'Arrange the windows if autoarrange is on
            For intCounter = 1 To 6
                Select Case picWindow(intCounter).Tag
                Case "tree"
                    Set frmPosition(intCounter) = mdiTreeChild
                Case "height"
                    Set frmPosition(intCounter) = mdiRowSelector
                Case "width"
                    Set frmPosition(intCounter) = mdiColSelector
                Case "depth"
                    Set frmPosition(intCounter) = mdiPageSelector
                Case "report"
                    Set frmPosition(intCounter) = mdiReportGrid
                Case Else
                    Set frmPosition(intCounter) = Nothing
                End Select
            Next
            'Do the actual arranging and Make sure the timer is on
            Call frmMaster.ArrangeChildren(conAutoArrange)
            frmMaster.tmrAuto.Enabled = True
        Else
            'Set the menu and timer off
            frmMaster.mnuWindowAutoArrange.Checked = False
            frmMaster.tmrAuto.Enabled = False
        End If
        'Then unload this form
        Unload Me
    Case 1
        'Cancel
        Unload Me
    Case Else
        'Do nothing
    End Select
Exit Sub
'Error Handler
ImmediateErrorSub_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub

Private Sub Form_Load()
'Created by Robin G Brown, 28/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Load"
    'Error Trap
    On Error GoTo Form_Load_ErrorHandler
    'Set up the form
    'Checkboxes
    If frmMaster.mnuWindowScratchpad.Checked = True Then
        chkOption(0).Value = 1
    Else
        chkOption(0).Value = 0
    End If
    If frmMaster.mnuWindowAutoArrange.Checked = True Then
        chkOption(1).Value = 1
        For intCounter = 1 To 6
            picWindow(intCounter).Enabled = True
        Next
    Else
        chkOption(1).Value = 0
        For intCounter = 1 To 6
            picWindow(intCounter).Enabled = False
        Next
    End If
    'Set up the picture boxes
    For intCounter = 1 To 6
        If Not frmPosition(intCounter) Is Nothing Then
            Select Case frmPosition(intCounter).Caption
            Case "height"
                picWindow(intCounter).Picture = frmMaster.imlDimension.ListImages(1).ExtractIcon
            Case "width"
                picWindow(intCounter).Picture = frmMaster.imlDimension.ListImages(2).ExtractIcon
            Case "depth"
                picWindow(intCounter).Picture = frmMaster.imlDimension.ListImages(3).ExtractIcon
            Case Else
                picWindow(intCounter).Picture = frmPosition(intCounter).Icon
            End Select
            picWindow(intCounter).Tag = frmPosition(intCounter).Caption
        Else
            picWindow(intCounter).Tag = "empty"
        End If
    Next
Exit Sub
'Error Handler
Form_Load_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Private Sub Form_Resize()
'Created by Robin G Brown, 28/4/97
'Default behaviour
'Sub Declares
Const conSub = "Form_Resize"
    'Error Trap
    On Error GoTo Form_Resize_ErrorHandler
    'Boilerplate code to center form on screen
    If WindowState > 0 Or conCentering <> True Then
        Exit Sub
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    'CODE_HERE
Exit Sub
'Error Handler
Form_Resize_ErrorHandler:
    Select Case Err.Number
    'Case ERROR_CODE_HERE
        'ERROR_HANDLING_CODE_HERE
    Case Else
        MsgBox "Unexpected error : " & Err.Description, vbExclamation, conModule & " : " & conSub
    End Select
    Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Created by Robin G Brown, 28/4/97
'Default behaviour
'Function Declares
Const conFunction = "Form_Unload"
Dim lngTempValue As Long
    'Errors 'disabled' to ensure that form unloads
    On Error Resume Next
    'CODE_HERE
End Sub

Private Sub picWindow_DragDrop(Index As Integer, Source As Control, X As Single, y As Single)
'Created by Robin G Brown, 29/4/97
'Move the pictures and tags around
'Sub Declares
Dim strTag As String
Dim imgPicture As Picture
    'Error Trap
    'On Error Resume Next
    Set imgPicture = picWindow(Index).Picture
    strTag = picWindow(Index).Tag
    Set picWindow(Index).Picture = Source.Picture
    picWindow(Index).Tag = Source.Tag
    Set Source.Picture = imgPicture
    Source.Tag = strTag
End Sub

Private Sub picWindow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'Created by Robin G Brown, 29/4/97
'Start the drag operation
'Sub Declares
    'Error Trap
    On Error Resume Next
    If Button = vbLeftButton Then
        picWindow(Index).DragIcon = picWindow(Index).Picture
        picWindow(Index).Drag vbBeginDrag
    End If
End Sub
