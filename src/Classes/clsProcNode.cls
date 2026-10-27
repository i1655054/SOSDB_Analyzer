VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ノードID
Public NodeID As Long

' 親ノードID
Public ParentID As Long

Public PrevNodeID As Long

Public NextNodeID As Long

' プロシージャ名
Public ProcName As String

Public CallName As String

' モジュール名
Public ModuleName As String

' ノード種別
Public NodeType As String

' 表示文字列
Public NodeText As String

' ソース行番号
Public LineNo As Long

' ネストレベル
Public Level As Long

' 分岐種別
Public BranchType As String

' ソース行
Public SourceLine As String
