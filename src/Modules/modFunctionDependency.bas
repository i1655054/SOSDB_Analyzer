Attribute VB_Name = "modFunctionDependency"
Option Explicit

'===================================================================
'@JPName     関数依存関係解析
'@Category   FunctionDependency
'@Summary    関数トレース結果から関数依存関係を集計し依存関係シートを作成する
'===================================================================

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
    
    ' 関数トレースシート
    Dim wsTrace As Worksheet
    
    ' 関数依存関係シート
    Dim wsDep As Worksheet

    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 呼出元|呼出先キー
    Dim Key As String
    
    ' Dictionary走査用
    Dim V As Variant

    ' 呼出回数管理
    Dim Dic As Object

    ' 呼出元JPName管理
    Dim DicCallerJP As Object
    
    ' 呼出先JPName管理
    Dim DicCalleeJP As Object

    ' 呼出元モジュール管理
    Dim DicFromModule As Object
    
    ' 呼出先モジュール管理
    Dim DicToModule As Object

    ' 呼出先カテゴリ管理
    Dim DicCategory As Object
    
    ' 呼出先概要管理
    Dim DicSummary As Object

    Set Dic = CreateObject("Scripting.Dictionary")

    Set DicCallerJP = CreateObject("Scripting.Dictionary")
    Set DicCalleeJP = CreateObject("Scripting.Dictionary")

    Set DicFromModule = CreateObject("Scripting.Dictionary")
    Set DicToModule = CreateObject("Scripting.Dictionary")

    Set DicCategory = CreateObject("Scripting.Dictionary")
    Set DicSummary = CreateObject("Scripting.Dictionary")

    If gContext Is Nothing Then Exit Sub
    
    '# FunctionTraceシート取得
    Set wsTrace = gContext.ManagementBook.Worksheets(SHEET_FUNCTION_TRACE)

    Application.DisplayAlerts = False
    
    On Error Resume Next
    
    '# FunctionDependencyシート再作成
    gContext.ManagementBook.Worksheets(SHEET_FUNCTION_DEPENDENCY).Delete
    
    On Error GoTo 0

    Application.DisplayAlerts = True
    
    '# FunctionDependencyシート作成
    Set wsDep = _
        gContext.ManagementBook.Worksheets.Add( _
            After:=gContext.ManagementBook.Worksheets( _
                gContext.ManagementBook.Worksheets.Count))

    wsDep.Name = SHEET_FUNCTION_DEPENDENCY

    '=========================
    ' ヘッダ
    '=========================
    '# FunctionDependencyヘッダ出力
    wsDep.Range("A1") = "呼出元"
    wsDep.Range("B1") = "呼出元JPName"

    wsDep.Range("C1") = "呼出先"
    wsDep.Range("D1") = "呼出先JPName"

    wsDep.Range("E1") = "回数"

    wsDep.Range("F1") = "呼出元モジュール"
    wsDep.Range("G1") = "呼出先モジュール"

    wsDep.Range("H1") = "呼出先Category"
    wsDep.Range("I1") = "呼出先Summary"

    '# FunctionTrace最終行取得
    LastRow = wsTrace.Cells( _
                    wsTrace.Rows.Count, _
                    "A").End(xlUp).row

    '=========================
    ' 集計
    '=========================
    '# 呼出関係集計
    For r = 2 To LastRow

        '# 呼出元・呼出先キー生成
        Key = _
            wsTrace.Cells(r, 1).Value & "|" & _
            wsTrace.Cells(r, 3).Value

        '# 呼出回数加算
        If Dic.Exists(Key) Then

            Dic(Key) = Dic(Key) + 1

        '# 呼出関係初回登録
        Else

            Dic.Add Key, 1

            '# 呼出元JPName保存
            DicCallerJP.Add _
                Key, wsTrace.Cells(r, 2).Value

            '# 呼出先JPName保存
            DicCalleeJP.Add _
                Key, wsTrace.Cells(r, 4).Value

            '# 呼出元モジュール保存
            DicFromModule.Add _
                Key, wsTrace.Cells(r, 5).Value

            '# 呼出先モジュール保存
            DicToModule.Add _
                Key, wsTrace.Cells(r, 6).Value

            '# 呼出先Category保存
            DicCategory.Add _
                Key, wsTrace.Cells(r, 7).Value

            '# 呼出先Summary保存
            DicSummary.Add _
                Key, wsTrace.Cells(r, 8).Value

        End If

    Next r

    '=========================
    ' 出力
    '=========================
    r = 2

    '# 依存関係一覧出力
    For Each V In Dic.Keys

        '# 呼出元情報出力
        wsDep.Cells(r, 1).Value = Split(V, "|")(0)
        wsDep.Cells(r, 2).Value = DicCallerJP(V)

        '# 呼出先情報出力
        wsDep.Cells(r, 3).Value = Split(V, "|")(1)
        wsDep.Cells(r, 4).Value = DicCalleeJP(V)

        '# 呼出回数出力
        wsDep.Cells(r, 5).Value = Dic(V)
        wsDep.Cells(r, 6).Value = DicFromModule(V)
        wsDep.Cells(r, 7).Value = DicToModule(V)

        '# 呼出先付帯情報出力
        wsDep.Cells(r, 8).Value = DicCategory(V)
        wsDep.Cells(r, 9).Value = DicSummary(V)

        r = r + 1

    Next V

    '# FunctionDependencyシート体裁設定
    wsDep.Rows(1).Font.Bold = True
    wsDep.Rows(1).AutoFilter
    wsDep.Columns.AutoFit

End Sub


