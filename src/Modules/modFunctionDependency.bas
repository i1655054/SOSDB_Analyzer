Attribute VB_Name = "modFunctionDependency"
Option Explicit

' 依存関係集計
Public Sub CreateFunctionDependency()

    Dim wsTrace As Worksheet
    Dim wsDep As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object
    Dim DicModule As Object
    
    Dim Key As String
    Dim V As Variant

    Set Dic = CreateObject("Scripting.Dictionary")
    Set DicModule = CreateObject("Scripting.Dictionary")

    Set wsTrace = Worksheets("関数トレース")

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("関数依存関係").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDep = Worksheets.Add
    wsDep.Name = "関数依存関係"

    wsDep.Range("A1") = "呼出元"
    wsDep.Range("B1") = "呼出先"
    wsDep.Range("C1") = "回数"
    wsDep.Range("D1") = "モジュール"

    LastRow = wsTrace.Cells(wsTrace.Rows.count, "A").End(xlUp).row

    '=========================
    ' 集計
    '=========================
    For r = 2 To LastRow

        Key = wsTrace.Cells(r, 1).Value & "|" & _
              wsTrace.Cells(r, 2).Value

        If Dic.Exists(Key) Then

            Dic(Key) = Dic(Key) + 1

        Else

            Dic.Add Key, 1
            
            ' 呼出先関数のモジュール
            DicModule.Add _
                Key, _
                wsTrace.Cells(r, 3).Value

        End If

    Next r
   
    '=========================
    ' 出力
    '=========================
    r = 2


    For Each V In Dic.Keys

        wsDep.Cells(r, 1) = Split(V, "|")(0)
        wsDep.Cells(r, 2) = Split(V, "|")(1)
        wsDep.Cells(r, 3) = Dic(V)
        wsDep.Cells(r, 4).Value = DicModule(V)

        r = r + 1

    Next V
    
    wsDep.Columns.AutoFit

    MsgBox "関数依存関係作成完了"

End Sub


