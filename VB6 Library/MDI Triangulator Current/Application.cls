VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Created by Robin G Brown, 2nd May 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    the TriReport 'Application' class, which is used by other applications _
    to interface with this application using ActiveX
'   IMPORTANT - the object created with this class is used to control _
    the application via ActiveX, so nothing from inside the application _
    should make calls to this object
'   Use app as a prefix for this class
'-----------------------------------------------------------------------------
'---Set Options
Option Explicit
Option Base 0
Option Compare Text
'---Events
'Instances of this class will need to be declared WithEvents if events are used
'EVENT_NAME_HERE is used to inform the object user _
    that EVENT_FUNCTIONALITY_HERE
'Public Event EVENT_NAME_HERE(ARGUMENTS_HERE)
'---All Declares
Private Const conModule = "Application"
'Public properties
'Private properties
Private intCounter As Integer
Private intReturn As Integer

Private Sub Class_Initialize()
'Created by Robin G Brown, 2/5/97
'Initialize the application when it is created as an object by another application, _
     without showing any forms
'Sub Declares
    'Error Trap
    On Error Resume Next
    'ApplicationInitialise is always the first code to run
    Call ApplicationInitialise
    'Load the report form but don't show it
    Load frmMaster
    'We can now manipulate the report form without it being displayed
End Sub

'-----------------------------------------------------------------------------
'   Public Functions and Methods
'-----------------------------------------------------------------------------

Public Sub Visible(ByVal booState As Boolean)
'Created by Robin G Brown, 2/5/97
'Show the master form...
'Sub Declares
    'Error Trap
    On Error Resume Next
    If booState = True Then
        frmMaster.Show
    Else
        frmMaster.Hide
    End If
End Sub

Public Sub FileNew()
'Created by Robin G Brown, 15/9/97
'Create a new report or graph
    Err.Raise vbObjectError, App.ProductName, "FileNew procedure not currently implemented."
End Sub

Public Sub FileOpen()
'Created by Robin G Brown, 15/9/97
'Open a report or graph file
    Err.Raise vbObjectError, App.ProductName, "FileOpen procedure not currently implemented."
End Sub

Public Sub FileSave()
'Created by Robin G Brown, 15/9/97
'Save the selected report or graph
    Err.Raise vbObjectError, App.ProductName, "FileSave procedure not currently implemented."
End Sub

'-----------------------------------------------------------------------------
'   Private Functions and Methods
'-----------------------------------------------------------------------------

'-----------------------------------------------------------------------------
'   Properties
'-----------------------------------------------------------------------------

'Public Property Get Application() As Application
''Created by Robin G Brown, 7/5/97
''Copied from 'Root Objects' topic in VB books online, _
'    returns a reference to the application object
''Sub Declares
'    'Error Trap
'    On Error Resume Next
'    Set Application = Me
'End Property

Private Sub Class_Terminate()
'Created by Robin G Brown, 2/5/97
'Get rid of any objects in memory and shut down
'Sub Declares
    'Error Trap
    On Error Resume Next
    'Close the application
    If Not frmMaster Is Nothing Then
        Unload frmMaster
        Set frmMaster = Nothing
    End If
End Sub



