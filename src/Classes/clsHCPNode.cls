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
'==================================================
' HCPノード
'==================================================
'
' @JPName
' HCPノード
'
' @Category
' HCP
'
' @Summary
' HCPチャートを構成するノード情報を保持する
'
' @Remarks
' ProcNodeから変換されたノード情報および
' 描画位置情報を管理する
'

' ノードID
Public NodeID As Long

' 親ノードID
Public ParentID As Long

' 前ノードID
Public PrevNodeID As Long

' 次ノードID
Public NextNodeID As Long

' 階層
Public Level As Long

' ノード種別
Public NodeType As String

' 分岐種別(TRUE/FALSE)
Public BranchType As String

' HCP表示文字列
Public Text As String

' 元ソースコード
Public SourceText As String

' ソース行番号
Public SourceLine As String

' HCPコメント
Public CommentText As String

'--------------------------------
' 業務説明用
'--------------------------------

' 業務説明文
Public BusinessText As String

' 描画X座標
Public PosX As Double

' 描画Y座標
Public PosY As Double

