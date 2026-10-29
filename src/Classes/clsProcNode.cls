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

'=================================================
' Procノード
'=================================================
'
' @JPName
' Procノード
'
' @Category
' Model
'
' @Summary
' ProcAnalyzerで生成する
' 制御フローノード情報を保持する
'
' @Remarks
' Procedure内の各処理行を
' ノード形式で管理する
'

' ノードID
Public NodeID As Long

' 親ノードID
Public ParentID As Long

' 前ノードID
Public PrevNodeID As Long

' 次ノードID
Public NextNodeID As Long

' プロシージャ名
Public ProcName As String

' 呼出先プロシージャ名
Public CallName As String

' モジュール名
Public ModuleName As String

' ノード種別
Public NodeType As String

' 表示文字列
Public NodeText As String

' 論理行番号
Public LineNo As Long

' ネストレベル
Public Level As Long

' 分岐種別(TRUE/FALSE)
Public BranchType As String

' ソースコード行
Public SourceLine As String
