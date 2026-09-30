Attribute VB_Name = "modFunctionDependency"
Option Explicit

'==================================================
' 関数依存関係
'==================================================

' @JPName
' 関数依存関係作成
'
' @Category
' FunctionDependency
'
' @Input
' なし
'
' @Output
' 関数依存関係シート
'
' @Summary
' 関数トレースシートを集計し
' 関数依存関係シートを作成する
'
' @Remarks
' 呼出元・呼出先単位で集計し
' 呼出回数および付帯情報を出力する
'
Public Sub FunctionDependency()

    Dim wsTrace As Worksheet
    Dim wsDep As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object

    Dim DicCallerJP As Object
    Dim DicCalleeJP As Object

    Dim DicFromModule As Object
    Dim DicToModule As Object

    Dim DicCategory As Object
    Dim DicSummary As Object

    Dim Key As String
    Dim V As Variant

    Set Dic = CreateObject("Scripting.Dictionary")

    Set DicCallerJP = CreateObject("Scripting.Dictionary")
    Set DicCalleeJP = CreateObject("Scripting.Dictionary")

    Set DicFromModule = CreateObject("Scripting.Dictionary")
    Set DicToModule = CreateObject("Scripting.Dictionary")

    Set DicCategory = CreateObject("Scripting.Dictionary")
    Set DicSummary = CreateObject("Scripting.Dictionary")

    Set wsTrace = Worksheets("関数トレース")

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("関数依存関係").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDep = Worksheets.Add
    wsDep.Name = "関数依存関係"

    '=========================
    ' ヘッダ
    '=========================
    wsDep.Range("A1") = "呼出元"
    wsDep.Range("B1") = "呼出元JPName"

    wsDep.Range("C1") = "呼出先"
    wsDep.Range("D1") = "呼出先JPName"

    wsDep.Range("E1") = "回数"

    wsDep.Range("F1") = "呼出元モジュール"
    wsDep.Range("G1") = "呼出先モジュール"

    wsDep.Range("H1") = "呼出先Category"
    wsDep.Range("I1") = "呼出先Summary"

    LastRow = wsTrace.Cells( _
                    wsTrace.Rows.Count, _
                    "A").End(xlUp).row

    '=========================
    ' 集計
    '=========================
    For r = 2 To LastRow

        Key = _
            wsTrace.Cells(r, 1).Value & "|" & _
            wsTrace.Cells(r, 3).Value

        If Dic.Exists(Key) Then

            Dic(Key) = Dic(Key) + 1

        Else

            Dic.Add Key, 1

            DicCallerJP.Add _
                Key, wsTrace.Cells(r, 2).Value

            DicCalleeJP.Add _
                Key, wsTrace.Cells(r, 4).Value

            DicFromModule.Add _
                Key, wsTrace.Cells(r, 5).Value

            DicToModule.Add _
                Key, wsTrace.Cells(r, 6).Value

            DicCategory.Add _
                Key, wsTrace.Cells(r, 7).Value

            DicSummary.Add _
                Key, wsTrace.Cells(r, 8).Value

        End If

    Next r

    '=========================
    ' 出力
    '=========================
    r = 2

    For Each V In Dic.Keys

        wsDep.Cells(r, 1).Value = Split(V, "|")(0)
        wsDep.Cells(r, 2).Value = DicCallerJP(V)

        wsDep.Cells(r, 3).Value = Split(V, "|")(1)
        wsDep.Cells(r, 4).Value = DicCalleeJP(V)

        wsDep.Cells(r, 5).Value = Dic(V)

        wsDep.Cells(r, 6).Value = DicFromModule(V)
        wsDep.Cells(r, 7).Value = DicToModule(V)

        wsDep.Cells(r, 8).Value = DicCategory(V)
        wsDep.Cells(r, 9).Value = DicSummary(V)

        r = r + 1

    Next V

    wsDep.Rows(1).Font.Bold = True
    wsDep.Rows(1).AutoFilter
    wsDep.Columns.AutoFit

    MsgBox "関数依存関係作成完了"

End Sub


