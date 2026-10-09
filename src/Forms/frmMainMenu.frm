VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMainMenu 
   Caption         =   "解析ツール"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2280
   OleObjectBlob   =   "frmMainMenu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' @JPName
' メインメニュー初期化
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' メインメニュー画面を初期化する
'
Private Sub UserForm_Initialize()

    Me.Caption = "解析ツール"

End Sub

' @JPName
' ソース管理起動
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ソース管理画面を表示する
'
Private Sub cmdExport_Click()

    Unload Me

    ' フォームを表示したまま他を操作できる
    frmExportTool.Show vbModeless
    
End Sub

' @JPName
' 関数管理起動
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 対象ツール選択画面を表示する
'
Private Sub cmdFunction_Click()

    Unload Me

    ' フォームを表示したまま他を操作できる
    frmTargetTool.Show vbModeless
    
End Sub

' @JPName
' メインメニュー終了
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' メインメニュー画面を閉じる
'
Private Sub cmdClose_Click()

    RestoreMainWindowState
    
    Unload Me

End Sub

' @JPName
' メインメニュー終了制御
'
' @Category
' UI
'
' @Input
' Cancel(Integer)
' CloseMode(Integer)
'
' @Output
' なし
'
' @Summary
' メインメニュー終了時の制御を行う
'
' @Remarks
' ×ボタン押下時は
' PathChart画面を再表示する
'
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        RestoreMainWindowState
    
    End If

End Sub

