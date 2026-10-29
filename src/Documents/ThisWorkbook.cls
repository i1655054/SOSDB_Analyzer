VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' @JPName
' Workbook起動
'
' @Category
' Workbook
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' Analyzer起動時の初期化処理を実行する
'
' @Remarks
' Analyzer初期化処理を呼び出し
' メインメニューを表示する
'
Private Sub Workbook_Open()

    InitializeAnalyzer

End Sub

' @JPName
' Workbook終了
'
' @Category
' Workbook
'
' @Input
' Cancel(Boolean)
'
' @Output
' なし
'
' @Summary
' Workbook終了時の後処理を実行する
'
' @Remarks
' Window状態を復元し
' 登録済みショートカットを解除する
' Cancelは使用しない
'
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error Resume Next
    
    RestoreMainWindowState
    
    UnRegisterShortcut

End Sub

