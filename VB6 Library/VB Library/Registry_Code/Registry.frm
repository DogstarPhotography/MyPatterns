VERSION 4.00
Begin VB.Form frmRegistry 
   Caption         =   "Registry Example Project"
   ClientHeight    =   6645
   ClientLeft      =   1770
   ClientTop       =   1545
   ClientWidth     =   6030
   Height          =   7050
   Icon            =   "Registry.frx":0000
   Left            =   1710
   LinkTopic       =   "Form1"
   MouseIcon       =   "Registry.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   6645
   ScaleWidth      =   6030
   Top             =   1200
   Width           =   6150
   Begin TabDlg.SSTab tabRegistry 
      Height          =   5355
      Left            =   120
      TabIndex        =   2
      Tag             =   "tabRegistry"
      Top             =   780
      Width           =   5835
      _Version        =   65536
      _ExtentX        =   10292
      _ExtentY        =   9446
      _StockProps     =   15
      Caption         =   "API"
      TabsPerRow      =   2
      Tab             =   1
      TabOrientation  =   0
      Tabs            =   2
      Style           =   0
      TabMaxWidth     =   1764
      TabHeight       =   529
      MousePointer    =   99
      MouseIcon       =   "Registry.frx":0614
      TabCaption(0)   =   "VB4"
      Tab(0).ControlCount=   12
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdRegistry(0)"
      Tab(0).Control(1)=   "chkRegistry(0)"
      Tab(0).Control(2)=   "txtRegistry(6)"
      Tab(0).Control(3)=   "txtRegistry(5)"
      Tab(0).Control(4)=   "txtRegistry(4)"
      Tab(0).Control(5)=   "txtRegistry(3)"
      Tab(0).Control(6)=   "txtRegistry(2)"
      Tab(0).Control(7)=   "lblRegistry(7)"
      Tab(0).Control(8)=   "lblRegistry(6)"
      Tab(0).Control(9)=   "lblRegistry(5)"
      Tab(0).Control(10)=   "lblRegistry(4)"
      Tab(0).Control(11)=   "lblRegistry(3)"
      TabCaption(1)   =   "API"
      Tab(1).ControlCount=   10
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblRegistry(0)"
      Tab(1).Control(1)=   "lblRegistry(1)"
      Tab(1).Control(2)=   "lblRegistry(2)"
      Tab(1).Control(3)=   "lblRegistry(8)"
      Tab(1).Control(4)=   "txtRegistry(0)"
      Tab(1).Control(5)=   "cboRegistry(0)"
      Tab(1).Control(6)=   "txtRegistry(1)"
      Tab(1).Control(7)=   "chkRegistry(1)"
      Tab(1).Control(8)=   "cmdRegistry(1)"
      Tab(1).Control(9)=   "txtRegistry(7)"
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   300
         MouseIcon       =   "Registry.frx":092E
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   3120
         Width           =   5115
      End
      Begin VB.CommandButton cmdRegistry 
         Caption         =   "&Quick"
         Height          =   375
         Index           =   1
         Left            =   4860
         TabIndex        =   22
         Top             =   420
         Width           =   855
      End
      Begin VB.CommandButton cmdRegistry 
         Caption         =   "&Quick"
         Height          =   375
         Index           =   0
         Left            =   -70140
         TabIndex        =   21
         Top             =   420
         Width           =   855
      End
      Begin VB.CheckBox chkRegistry 
         Caption         =   "Lock Root Key and Key Boxes"
         Height          =   255
         Index           =   1
         Left            =   1440
         MouseIcon       =   "Registry.frx":0C38
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   660
         Value           =   2  'Grayed
         Width           =   2835
      End
      Begin VB.CheckBox chkRegistry 
         Caption         =   "Lock AppName, Section and  Key Boxes"
         Height          =   255
         Index           =   0
         Left            =   -73740
         MouseIcon       =   "Registry.frx":0F42
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   600
         Value           =   2  'Grayed
         Width           =   3255
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   -74640
         MouseIcon       =   "Registry.frx":124C
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Tag             =   "txtRegistry"
         Top             =   4860
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   -74640
         MouseIcon       =   "Registry.frx":1556
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Tag             =   "txtRegistry"
         Top             =   1500
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   -74640
         MouseIcon       =   "Registry.frx":1860
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Tag             =   "txtRegistry"
         Top             =   2340
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   -74640
         MouseIcon       =   "Registry.frx":1B6A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Tag             =   "txtRegistry"
         Top             =   3180
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   -74640
         MouseIcon       =   "Registry.frx":1E74
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Tag             =   "txtRegistry"
         Top             =   4020
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   300
         MouseIcon       =   "Registry.frx":217E
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   2280
         Width           =   5115
      End
      Begin VB.ComboBox cboRegistry 
         Height          =   300
         Index           =   0
         ItemData        =   "Registry.frx":2488
         Left            =   300
         List            =   "Registry.frx":249B
         MouseIcon       =   "Registry.frx":2500
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "cboRegistry"
         Top             =   1440
         Width           =   5115
      End
      Begin VB.TextBox txtRegistry 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   300
         MouseIcon       =   "Registry.frx":280A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3960
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter 'Key' name To Use"
         Height          =   315
         Index           =   8
         Left            =   300
         MouseIcon       =   "Registry.frx":2B14
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Tag             =   "RegistryLabel"
         Top             =   2700
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Default to use with GetSetting:"
         Height          =   315
         Index           =   7
         Left            =   -74640
         MouseIcon       =   "Registry.frx":2E1E
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Tag             =   "RegistryLabel"
         Top             =   4440
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter ApplicationName To Use:"
         Height          =   315
         Index           =   6
         Left            =   -74640
         MouseIcon       =   "Registry.frx":3128
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Tag             =   "RegistryLabel"
         Top             =   1080
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Section To Use:"
         Height          =   315
         Index           =   5
         Left            =   -74640
         MouseIcon       =   "Registry.frx":3432
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Tag             =   "RegistryLabel"
         Top             =   1920
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Key To Use"
         Height          =   315
         Index           =   4
         Left            =   -74640
         MouseIcon       =   "Registry.frx":373C
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Tag             =   "RegistryLabel"
         Top             =   2760
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Value to send to Registry:"
         Height          =   315
         Index           =   3
         Left            =   -74640
         MouseIcon       =   "Registry.frx":3A46
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Tag             =   "RegistryLabel"
         Top             =   3600
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Section Name(s) To Use"
         Height          =   315
         Index           =   2
         Left            =   300
         MouseIcon       =   "Registry.frx":3D50
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Tag             =   "RegistryLabel"
         Top             =   1860
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter Value to send to Registry:"
         Height          =   315
         Index           =   1
         Left            =   300
         MouseIcon       =   "Registry.frx":405A
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Tag             =   "RegistryLabel"
         Top             =   3540
         Width           =   5115
      End
      Begin VB.Label lblRegistry 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select Root Key:"
         Height          =   315
         Index           =   0
         Left            =   300
         MouseIcon       =   "Registry.frx":4364
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "RegistryLabel"
         Top             =   1080
         Width           =   5115
      End
   End
   Begin ComctlLib.StatusBar stbrRegistry 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Tag             =   "stbrRegistry"
      Top             =   6270
      Width           =   6030
      _Version        =   65536
      _ExtentX        =   10636
      _ExtentY        =   661
      _StockProps     =   68
      AlignSet        =   -1  'True
      MouseIcon       =   "Registry.frx":466E
      MousePointer    =   99
      Style           =   1
      SimpleText      =   "Select A Function Set, Type Some Text, And Push Some Buttons."
      i1              =   "Registry.frx":4988
   End
   Begin ComctlLib.Toolbar tlbrRegistry 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6030
      _Version        =   65536
      _ExtentX        =   10636
      _ExtentY        =   1085
      _StockProps     =   96
      BorderStyle     =   1
      ImageList       =   ""
      MousePointer    =   99
      MouseIcon       =   "Registry.frx":4AFB
      ButtonWidth     =   1005
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      NumButtons      =   14
      i1              =   "Registry.frx":4E15
      i2              =   "Registry.frx":4FC4
      i3              =   "Registry.frx":518B
      i4              =   "Registry.frx":5346
      i5              =   "Registry.frx":5509
      i6              =   "Registry.frx":56B0
      i7              =   "Registry.frx":586B
      i8              =   "Registry.frx":5A26
      i9              =   "Registry.frx":5BE9
      i10             =   "Registry.frx":5DAC
      i11             =   "Registry.frx":5F70
      i12             =   "Registry.frx":612C
      i13             =   "Registry.frx":62F0
      i14             =   "Registry.frx":64A0
      AlignSet        =   -1  'True
      Wrappable       =   0   'False
   End
End
Attribute VB_Name = "frmRegistry"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'declare variables for boilerplate code
Dim Centering As Integer
Dim Floating As Integer
'end of boilerplate

'flag for API state
'Dim intAPIKeyOpen As Integer

'constants for checkboxes
Const UNCHECKED = 0
Const CHECKED = 1
Const GREYED = 2
Const CHK_VB4 = 0
Const CHK_API = 1

'flag to track current tab
'Dim intTabShowing As Integer
'constants for tab control
Const TAB_VB4 = 0
Const TAB_API = 1

'constants for Menu Buttons
Const BUTTON_GAP = (1 Or 5 Or 13)
Const BUTTON_SAVE = 2
Const BUTTON_GET = 3
Const BUTTON_DEL_VB4 = 4
Const BUTTON_CREATE = 6
Const BUTTON_OPEN = 7
Const BUTTON_SET = 8
Const BUTTON_QUERY = 9
Const BUTTON_DEL_API = 10
Const BUTTON_CLOSE = 11
Const BUTTON_EXIT = 14
Const BUTTON_OTHER = 12

'constsants for combo boxes
Const CBO_API = 0
Const CB0_HKEY_USERS = 0
Const CB0_HKEY_CURRENT_USER = 1
Const CB0_HKEY_LOCAL_MACHINE = 2
Const CB0_HKEY_CLASSES_ROOT = 3
Const MATCH_VB4 = "Software\VB and VBA Program Settings\"

'constants for Textboxes
Const TXT_API_VALUE = 0
Const TXT_API_KEY = 1
Const TXT_API_KEYNAME = 7
Const TXT_VB4_SETTING = 2
Const TXT_VB4_KEY = 3
Const TXT_VB4_SECTION = 4
Const TXT_VB4_APPNAME = 5
Const TXT_VB4_DEFAULT = 6

Private Sub cboRegistry_Click(Index As Integer)
'MRM Jan 96
'to set the values to match the VB ones
    On Error Resume Next
    
    Select Case Index
        
        Case CBO_API
            'make text in boxes match
            If cboRegistry(CBO_API).ListIndex = cboRegistry(CBO_API).ListCount - 1 Then
                txtRegistry(TXT_API_KEY).Text = MATCH_VB4
                cboRegistry(CBO_API).ListIndex = cboRegistry(CBO_API).ItemData(cboRegistry(CBO_API).ListIndex)
            
                If txtRegistry(TXT_VB4_APPNAME) <> "" And txtRegistry(TXT_VB4_SECTION) <> "" And txtRegistry(TXT_VB4_KEY) <> "" Then
                    txtRegistry(TXT_API_KEY).Text = txtRegistry(TXT_API_KEY).Text & txtRegistry(TXT_VB4_APPNAME) & "\" & txtRegistry(TXT_VB4_SECTION)
                    txtRegistry(TXT_API_KEYNAME).Text = txtRegistry(TXT_VB4_KEY)
                End If
            End If
            
        Case Else
            'do nothing
            
    End Select
    
    Me.Refresh
    
    On Error GoTo 0
    

End Sub

Private Sub chkRegistry_Click(Index As Integer)
'MRM Jan 96
'to lock out the text so you don't accidentally change the values
Dim intLockedBoxes As Boolean
    
    On Error Resume Next
    
    If chkRegistry(Index).Value = CHECKED Then
        intLockedBoxes = True
    Else
        intLockedBoxes = False
    End If
    
    Select Case Index
        Case CHK_API
            txtRegistry(TXT_API_KEY).Locked = intLockedBoxes
            txtRegistry(TXT_API_KEYNAME).Locked = intLockedBoxes
            cboRegistry(CBO_API).Enabled = Not (intLockedBoxes)
            
        Case CHK_VB4
            txtRegistry(TXT_VB4_KEY).Locked = intLockedBoxes
            txtRegistry(TXT_VB4_SECTION).Locked = intLockedBoxes
            txtRegistry(TXT_VB4_APPNAME).Locked = intLockedBoxes
        
        Case Else
            'do nothing
    End Select
    
    Me.Refresh
    
    On Error GoTo 0

End Sub

Private Sub cmdRegistry_Click(Index As Integer)
'MRM Jan 96
'makes api and vb4 entries the same: both buttons to do the same
    On Error Resume Next
    
    txtRegistry(TXT_API_VALUE).Text = "Test For Fish"
    txtRegistry(TXT_API_KEY).Text = ""
    txtRegistry(TXT_VB4_SETTING).Text = "Test For Fish"
    txtRegistry(TXT_VB4_KEY).Text = "MyKey"
    txtRegistry(TXT_VB4_SECTION).Text = "MySection"
    txtRegistry(TXT_VB4_APPNAME).Text = "MyApp"
    txtRegistry(TXT_VB4_DEFAULT).Text = "Haddock And Kippers"
    
    cboRegistry(CBO_API).ListIndex = cboRegistry(CBO_API).ListCount - 1
    
    On Error GoTo 0


End Sub

Private Sub Form_Activate()
Dim lngTempValue As Long

    On Error GoTo Form_Activate_ErrorHandler
    'this section of code causes window to be 'always on top' until it is 'sunk' with another call to floatwindow, see form_unload event
    If Floating = True Then
        Me.Show
        lngTempValue = Screen.ActiveForm.hWnd
        Call FloatWindow(lngTempValue, FLOAT)
    End If

Exit Sub
Form_Activate_ErrorHandler:
    MsgBox "Unexpected Error - " & Error$(Err), vbInformation, "Error"
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'the next line must be commented out to disable calling SOS4WIN
    'Call S4WVB(KeyCode)

End Sub


Private Sub Form_Load()

    On Error GoTo Form_Load_ErrorHandler
    'Set variables for boilerplate code comment out to prevent
    Centering = True    'causes form to center on screen when shown, check form_resize event
    Floating = False    'causes form to be 'always on top', check form_activate and form_unload events for other bits
    'To enable SOS4WIN to be called follow these instructions (refer to SOS4WIN manual pp.77-79)
    'Add this line to the keydown event for any objects that will call help, rmember that the forms keypreview property must be false for this to work
    'Call S4WVB(KeyCode)
    'End of boilerplate

    tabRegistry.Tab = TAB_VB4
    cboRegistry(CBO_API).ListIndex = 0
    cboRegistry(CBO_API).Refresh
    
    txtRegistry(TXT_API_VALUE).Text = ""
    txtRegistry(TXT_API_KEY).Text = ""
    txtRegistry(TXT_VB4_SETTING).Text = ""
    txtRegistry(TXT_VB4_KEY).Text = ""
    txtRegistry(TXT_VB4_SECTION).Text = ""
    txtRegistry(TXT_VB4_APPNAME).Text = ""
    
    Me.Refresh
    Beep
    
Exit Sub
Form_Load_ErrorHandler:
    MsgBox "Unexpected Error - " & Error$(Err), vbInformation, "Error"
    Exit Sub

End Sub

Private Sub Form_Resize()

    'boilerplate code to center form on screen
    On Error GoTo Form_Resize_ErrorHandler
    If WindowState > 0 Or Centering <> True Then
        Exit Sub
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    'end of boilerplate

Exit Sub
Form_Resize_ErrorHandler:
    MsgBox "Unexpected Error - " & Error$(Err), vbInformation, "Error"
    Exit Sub


End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim lngTempValue As Long

    'this section of code should also be used whenever the form is hidden
    On Error Resume Next
    If Floating <> False Then
        lngTempValue = Screen.ActiveForm.hWnd
        Call FloatWindow(lngTempValue, SINK)
    End If
    On Error GoTo 0

End Sub


Private Sub tabRegistry_Click(PreviousTab As Integer)
'MRM Jan 96
'To Enable and Disable the Toolbar Buttons
'BUTTON_GAP = (1 Or 5 Or 11 Or 12)
Dim strTYPE As String
Dim btnToolbarButton As Button

    'On Error Resume Next
    
    Select Case tabRegistry.Tab
        Case TAB_API
            'reset menu enabled properties
            strTYPE = "API"
        Case Else
            'reset menu enabled properties
            strTYPE = "VB4"
    End Select
    
    For Each btnToolbarButton In tlbrRegistry.Buttons
        If InStr(1, btnToolbarButton.Tag, strTYPE, 1) > 0 Then
            btnToolbarButton.Visible = True
        Else
            btnToolbarButton.Visible = False
        End If
    Next
    
    tlbrRegistry.Buttons(BUTTON_EXIT).Visible = True
    
    Me.Refresh
    
    On Error GoTo 0

End Sub

Private Sub tlbrRegistry_ButtonClick(ByVal Button As Button)
'MRM JAN 96
'To call the correct function for each button that is pressed.
Dim intResult As Integer
Dim lngNewAPIHandle As Long
Dim strReturnMessage As String
Dim strReturnKeyValue As String
Dim strOtherMachineName As String

    On Error GoTo tlbrRegistry_BC_Errorhandler
    
    'Set intResult to <> TRUE OR FALSE
    intResult = 2

    'MsgBox Button.Index & " " & Button.Tag

    'Do intialisations
    Select Case Button.Index
        'check if api used
        Case BUTTON_OTHER, BUTTON_SET, BUTTON_CREATE, BUTTON_OPEN, BUTTON_QUERY, BUTTON_DEL_API, BUTTON_CLOSE
            If Not (txtRegistry(TXT_API_KEY) <> "") Then GoTo tlbrRegistry_BC_EmptyBoxes:
    
            'get handle for api
            Select Case cboRegistry(CBO_API).ItemData(cboRegistry(CBO_API).ListIndex)
                Case CB0_HKEY_USERS
                    lngNewAPIHandle = HKEY_USERS
                    
                Case CB0_HKEY_CURRENT_USER
                    lngNewAPIHandle = HKEY_CURRENT_USER
                    
                Case CB0_HKEY_LOCAL_MACHINE
                    lngNewAPIHandle = HKEY_LOCAL_MACHINE
                    
                Case CB0_HKEY_CLASSES_ROOT
                    lngNewAPIHandle = HKEY_CLASSES_ROOT
                    
                Case Else
                    'error: invalid entry
                    GoTo tlbrRegistry_BC_InvalidAPIHandle
            End Select
            
            'check the vaue
            If Not (lngNewAPIHandle = HKEY_CLASSES_ROOT Or lngNewAPIHandle = HKEY_CURRENT_USER Or lngNewAPIHandle = HKEY_LOCAL_MACHINE Or lngNewAPIHandle = HKEY_USERS) Then
tlbrRegistry_BC_InvalidAPIHandle:
                MsgBox "You have somehow choosen an invalid Root Key. Please choose another", vbInformation, "Registry Example Project"
                Exit Sub
            End If
            
        Case BUTTON_SAVE, BUTTON_GET, BUTTON_DEL_VB4
            'VB4 is used
            If Not (txtRegistry(TXT_VB4_APPNAME) <> "" And txtRegistry(TXT_VB4_SECTION) <> "" And txtRegistry(TXT_VB4_KEY) <> "") Then
tlbrRegistry_BC_EmptyBoxes:
                MsgBox "Please fill in all the needed Textboxes first!", vbInformation, "Registry Example Project"
                Exit Sub
            End If
        
        Case Else 'inc. BUTTON_EXIT
            'do nothing
            
    End Select
    
    Select Case Button.Index
        Case BUTTON_SAVE
            'USE: Function WrappedSaveSetting(ByVal strAppName As String, ByVal strSectionName As String, ByVal strKeyName As String, ByVal strSetting As String, ByRef strGiveMessage As String) As Int
            intResult = WrappedSaveSetting(txtRegistry(TXT_VB4_APPNAME).Text, txtRegistry(TXT_VB4_SECTION).Text, txtRegistry(TXT_VB4_KEY).Text, txtRegistry(TXT_VB4_SETTING).Text, strReturnMessage)
            
        Case BUTTON_GET
            'USE: Function WrappedGetSetting(ByVal strAppName As String, ByVal strSectionName As String, ByVal strKeyName As String, ByRef strGiveMessage As String, Optional ByVal strDefault As Variant) As Str
            'strTemp = txtRegistry(TXT_VB4_DEFAULT).Text
            intResult = WrappedGetSetting(txtRegistry(TXT_VB4_APPNAME).Text, txtRegistry(TXT_VB4_SECTION).Text, txtRegistry(TXT_VB4_KEY).Text, strReturnMessage, txtRegistry(TXT_VB4_DEFAULT).Text)
            
        Case BUTTON_DEL_VB4
            'USE: Function WrappedDeleteSetting(ByVal strAppName As String, ByVal strSectionName As String, ByRef strGiveMessage As String, Optional ByVal strKeyName As Variant) As Int
            If Not (txtRegistry(TXT_VB4_DEFAULT) <> "") Then GoTo tlbrRegistry_BC_EmptyBoxes
            intResult = WrappedDeleteSetting(txtRegistry(TXT_VB4_APPNAME).Text, txtRegistry(TXT_VB4_SECTION).Text, strReturnMessage, txtRegistry(TXT_VB4_KEY).Text)
            
        Case BUTTON_CREATE
            'USE: Function APICreateKey(ByVal lngRootKey As Long, ByVal strFullKeyName As String, ByRef strGiveMessage As String) As Int
            'and open new key: lngNewAPIHandle
            intResult = APICreateKey(lngNewAPIHandle, txtRegistry(TXT_API_KEY).Text, strReturnMessage)
        
        Case BUTTON_OPEN
            'USE: Function APIOpenKey(ByVal lngRootKey As Long, ByVal strFullKeyName As String, ByRef strGiveMessage As String) As Int
            'and open new key: lngNewAPIHandle
            intResult = APIOpenKey(lngNewAPIHandle, txtRegistry(TXT_API_KEY).Text, strReturnMessage)
            
        Case BUTTON_SET
            'USE: Function APISetKey(ByVal strMyValueName As String, ByVal lngDataType As Long, ByVal strMyValue As String, ByVal lngLength As Long, ByRef strGiveMessage As String) As Integer
            If Not (txtRegistry(TXT_API_VALUE) <> "") Then GoTo tlbrRegistry_BC_EmptyBoxes
            intResult = APISetKey(txtRegistry(TXT_API_KEYNAME).Text, REG_SZ, txtRegistry(TXT_API_VALUE).Text, Len(txtRegistry(TXT_API_VALUE).Text) + 1, strReturnMessage)
        
        Case BUTTON_QUERY
            'USE: Function APIQueryKey(ByVal strMyValueName As String, ByVal strMyValue As String, ByRef strGiveMessage As String) As Int
            intResult = APIQueryKey(txtRegistry(TXT_API_KEYNAME).Text, strReturnKeyValue, strReturnMessage)
            
        Case BUTTON_DEL_API
            'USE: Function APIDeleteKey(ByVal strMyValueName As String, ByRef strGiveMessage As String) As Int
            intResult = APIDeleteKey(txtRegistry(TXT_API_KEYNAME).Text, strReturnMessage)
            
        Case BUTTON_CLOSE
            'USE: Function APICloseKey(ByRef strGiveMessage As String) As Int
            intResult = APICloseKey(strReturnMessage)
            
        Case BUTTON_OTHER
            'USE: Function APIConnectOtherRegistry(ByVal lngRootKey As Long, ByVal strMachineName As String, ByRef strGiveMessage As String) As Integer
            strOtherMachineName = InputBox("Please enter the machine name to connect to. For this one, enter 'Me' or press Cancel.", "Registry Example Project", "Me!")
            If InStr(1, strOtherMachineName, "ME", 1) = 1 Or strOtherMachineName = "" Then strOtherMachineName = vbNullString
            
            intResult = APIConnectOtherRegistry(lngNewAPIHandle, strOtherMachineName, strReturnMessage)
            
        Case BUTTON_EXIT
            'USE: NONE
            End
            
        Case BUTTON_GAP
            'USE: NONE
            'do nothing
            
        Case Else 'always use this!!!
            'do nothing: msgbox in case to flag user
            MsgBox Button.ToolTipText, vbInformation, "Registry Example Project"
            
    End Select
    
    If ((Button.Index = BUTTON_OPEN Or Button.Index = BUTTON_CREATE) And intResult = True) Or (Button.Index = BUTTON_CLOSE And intResult = False) Then
        'disable Create, Open, enable Close
        tlbrRegistry.Buttons(BUTTON_OPEN).Enabled = False
        tlbrRegistry.Buttons(BUTTON_CREATE).Enabled = False
        tlbrRegistry.Buttons(BUTTON_CLOSE).Enabled = True
    ElseIf ((Button.Index = BUTTON_OPEN Or Button.Index = BUTTON_CREATE) And intResult = False) Or (Button.Index = BUTTON_CLOSE And intResult = True) Then
        'enable Create, Open, disable Close
        tlbrRegistry.Buttons(BUTTON_OPEN).Enabled = True
        tlbrRegistry.Buttons(BUTTON_CREATE).Enabled = True
        tlbrRegistry.Buttons(BUTTON_CLOSE).Enabled = False
    End If
 
    If strReturnMessage <> "IGNORE" Then MsgBox "Action Returned The Following Value:" & vbCrLf & strReturnMessage, vbInformation, "Registry Example Project"
    
    Exit Sub

tlbrRegistry_BC_Errorhandler:
    MsgBox "Error #" & Err & " in ToolBar Click: " & Error, vbInformation, "Registry Example Project"
    Exit Sub


End Sub


