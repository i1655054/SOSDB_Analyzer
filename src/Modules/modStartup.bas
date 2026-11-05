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

    '# Analyzer環境初期化
    '# PathChartショートカット登録
    RegisterShortcut

    '# PathChart設定読込
    LoadPathChartConfig
    
    '# Excel画面状態保存
    SaveMainWindowState

    '# メイン画面表示
    '# メインシート表示
    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    
    '# メイン画面サイズ調整
    ResizeMainWindow

    ' フォームを表示したまま他を操作できる
    '# メインメニュー表示
    frmMainMenu.Show vbModeless

End Sub

