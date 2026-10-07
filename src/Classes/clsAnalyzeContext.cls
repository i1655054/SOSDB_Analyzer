VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnalyzeContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==================================================
' 解析コンテキスト
'==================================================
'
' @JPName
' 解析コンテキスト
'
' @Category
' Context
'
' @Summary
' 解析対象ブックおよび管理ブック、
' ツール情報を保持する
'
' @Remarks
' 解析処理間で共有する実行コンテキスト
'

' 解析対象ブック
Public TargetBook As Workbook

' 管理ブック
Public ManagementBook As Workbook

' ツール情報
Public Tool As clsToolInfo

'Public ToolID As Long

' @JPName
' コンテキスト妥当性確認
'
' @Category
' Context
'
' @Input
' なし
'
' @Output
' IsValid(Boolean)
'
' @Summary
' 解析コンテキストが有効か判定する
'
' @Remarks
' 対象ブックおよび管理ブックの
' 設定有無を確認する
'
Public Function IsValid() As Boolean

    IsValid = _
        Not TargetBook Is Nothing _
        And _
        Not ManagementBook Is Nothing _
        And _
        Not Tool Is Nothing

End Function

