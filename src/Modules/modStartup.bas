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
' Analyzer起動時の初期化処理を実行する
'
' @Remarks
' ショートカット登録
' PathChart設定読込
' メインシート表示
' メインメニュー表示を実施する
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

