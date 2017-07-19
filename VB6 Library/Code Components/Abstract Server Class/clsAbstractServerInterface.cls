VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IServerScriptAbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created by Robin G Brown, 4th August 1997
'-----------------------------------------------------------------------------
'   This class module contains code for _
    an abstract class interface for server objects _
    to be used by the script server engine
'-----------------------------------------------------------------------------

Public Function ServerInitialize(ByVal strDataFilename As String, ByVal strDestinationFilename As String, Optional varOptions As Variant) As Long

End Function

Public Function ServerProcess() As Long

End Function

Public Function ServerClose() As Long

End Function





