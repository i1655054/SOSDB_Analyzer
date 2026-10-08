Attribute VB_Name = "modStartup"
Option Explicit

'==================================================
' Main Tool Book サイズ
'==================================================

' Main画面幅
Private Const MAIN_WIDTH  As Long = 400

' Main画面高さ
Private Const MAIN_HEIGHT As Long = 250

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

    ThisWorkbook.Worksheets(SHEET_MAIN).Activate
    
    With ThisWorkbook.Windows(1)

        .WindowState = xlNormal

        .Top = 0
        .Left = 0

        .Width = MAIN_WIDTH
        .Height = MAIN_HEIGHT

    End With

    ' フォームを表示したまま他を操作できる
    frmMainMenu.Show vbModeless

End Sub

