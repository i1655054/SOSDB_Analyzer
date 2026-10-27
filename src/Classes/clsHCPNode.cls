VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHCPNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NodeID As Long

Public ParentID As Long

Public PrevNodeID As Long

Public NextNodeID As Long

Public Level As Long

Public NodeType As String

Public BranchType As String

Public Text As String

Public SourceText As String

Public SourceLine As String

'--------------------------------
' ã∆ñ±ê‡ñæóp
'--------------------------------
Public BusinessText As String

Public PosX As Double
Public PosY As Double
