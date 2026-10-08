Attribute VB_Name = "modStartup"
Option Explicit

'==================================================
' 起動時Windowサイズ
'==================================================

' 起動時画面幅
Public gMainOldWidth As Double

' 起動時画面高さ
Public gMainOldHeight As Double

' 起動時Left
Public gMainOldLeft As Double

' 起動時Top
Public gMainOldTop As Double

' 起動時旧画面状態
Public gMainOldState As XlWindowState

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

