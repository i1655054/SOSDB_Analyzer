Attribute VB_Name = "modHistory"
Option Explicit

'==================================================
' 履歴管理
'
' ナビゲータで利用するシート遷移履歴を管理する
'==================================================

' 遷移履歴
Public History As Collection

' 現在履歴位置
Public CurrentPos As Long

' 履歴移動中フラグ
Public IsHistoryMoving As Boolean

' @JPName
' 履歴初期化
'
' @Category
' Navigator
'
' @Summary
' シート遷移履歴を初期化する
'
Public Sub InitHistory()

    Set History = New Collection
    CurrentPos = 0

End Sub

' @JPName
' 履歴追加
'
' @Category
' Navigator
'
' @Input
' SheetName(String)
'
' @Output
' なし
'
' @Summary
' 指定シートを履歴へ追加する
'
' @Remarks
' 非表示シートおよび重複履歴は追加しない
'
Public Sub AddHistory(ByVal SheetName As String)

    Dim ws As Worksheet

    Set ws = Worksheets(SheetName)

    If ws.Visible <> xlSheetVisible Then Exit Sub
    
    ' メニューシートは履歴対象外
    If ws.Name = "Sheet1" Then Exit Sub

    If History Is Nothing Then
        Set History = New Collection
    End If

    If History.Count > 0 Then
        If History(History.Count) = SheetName Then Exit Sub
    End If

    History.Add SheetName

    CurrentPos = History.Count

End Sub

' @JPName
' 前履歴取得
'
' @Category
' Navigator
'
' @Output
' SheetName(String)
'
' @Summary
' 一つ前の履歴シート名を取得する
'
Public Function GetPreviousSheet() As String

    If History Is Nothing Then
        Exit Function
    End If

    If CurrentPos <= 1 Then
        Exit Function
    End If

    CurrentPos = CurrentPos - 1

    GetPreviousSheet = History(CurrentPos)

End Function

' @JPName
' 次履歴取得
'
' @Category
' Navigator
'
' @Output
' SheetName(String)
'
' @Summary
' 一つ後の履歴シート名を取得する
'
Public Function GetNextSheet() As String

    If History Is Nothing Then Exit Function
    If CurrentPos >= History.Count Then Exit Function

    CurrentPos = CurrentPos + 1

    GetNextSheet = History(CurrentPos)

End Function
