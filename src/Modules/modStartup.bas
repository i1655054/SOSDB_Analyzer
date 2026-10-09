Attribute VB_Name = "modStartup"
Option Explicit

' @JPName
' Analyzer初期化
'
' @Category
' Common
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' Analyzer起動時の初期設定を行う
'
' @Remarks
' ショートカット登録および
' PathChart設定読込を実施する
'
Public Sub InitializeAnalyzer()

    RegisterShortcut

    LoadPathChartConfig
    
    SaveMainWindowState

    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    
    ResizeMainWindow

    ' フォームを表示したまま他を操作できる
    frmMainMenu.Show vbModeless

End Sub

