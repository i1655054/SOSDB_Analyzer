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
' @Summary
' ソース管理画面を表示する
'
Private Sub cmdExport_Click()

    frmExportTool.Show vbModal

End Sub

' @JPName
' 関数管理起動
'
' @Category
' UI
'
' @Summary
' 対象ツール選択画面を表示する
'
Private Sub cmdFunction_Click()

    frmTargetTool.Show vbModeless

End Sub

' @JPName
' メインメニュー終了
'
' @Category
' UI
'
' @Summary
' メインメニュー画面を閉じる
'
Private Sub cmdClose_Click()

    Unload Me

End Sub


