Attribute VB_Name = "modNavigator"
Option Explicit

'==================================================
' Navigator管理
'==================================================

' @JPName
' Navigator表示
'
' @Category
' Navigator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' Navigator画面を表示する
'
' @Remarks
' 既に表示中の場合は再表示しない
'
Public Sub ShowNavigator()

    If Not frmNavigator.Visible Then
        frmNavigator.Show vbModeless
    End If

End Sub

' @JPName
' Navigator更新
'
' @Category
' Navigator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 現在アクティブなシートに合わせて
' Navigatorの選択状態を更新する
'
' @Remarks
' Navigator表示中のみ実行する
'
Public Sub UpdateNavigator()

    If frmNavigator.Visible Then

        frmNavigator.SelectCurrentSheet

    End If

End Sub
