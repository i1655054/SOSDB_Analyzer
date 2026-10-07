Attribute VB_Name = "modFunctionTrac"
Option Explicit

'==================================================
' 関数トレース
'==================================================

' @JPName
' 関数トレース作成
'
' @Category
' FunctionTrace
'
' @Input
' なし
'
' @Output
' 関数トレースシート
'
' @Summary
' 対象ブックの関数呼出関係を解析し
' 関数トレースシートを作成する
'
' @Remarks
' 呼出元、呼出先、モジュール、
' Category、Summaryを出力する
'
Public Sub FunctionTrace()
       
    Dim wsOut As Worksheet
    
    Dim VBComp As Object
    Dim FuncList As Object
    
    Dim CodeText As String

    Dim RowOut As Long

    If gContext Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    Set wsOut = gContext.ManagementBook.Worksheets(SHEET_FUNCTION_TRACE)
    
    On Error GoTo 0

    If wsOut Is Nothing Then

        Set wsOut = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                    gContext.ManagementBook.Worksheets.Count))
        
        wsOut.Name = SHEET_FUNCTION_TRACE

    Else

        wsOut.Cells.Clear

    End If

    wsOut.Range("A1") = "呼出元"
    wsOut.Range("B1") = "呼出元JPName"

    wsOut.Range("C1") = "呼出先"
    wsOut.Range("D1") = "呼出先JPName"

    wsOut.Range("E1") = "呼出元モジュール"
    wsOut.Range("F1") = "呼出先モジュール"

    wsOut.Range("G1") = "呼出先Category"
    wsOut.Range("H1") = "呼出先Summary"

    wsOut.Range("I1") = "行番号"
    wsOut.Range("J1") = "元コード"

    Set FuncList = CreateObject("Scripting.Dictionary")

    '======================
    ' 関数一覧取得
    '======================

    For Each VBComp In gContext.TargetBook.VBProject.VBComponents

        If VBComp.Type = 1 Then

            If VBComp.CodeModule.CountOfLines > 0 Then

                CodeText = VBComp.CodeModule.Lines( _
                            1, _
                            VBComp.CodeModule.CountOfLines)

                Call GetFunctions(CodeText, VBComp.Name, FuncList)

            End If

        End If

    Next

    '======================
    ' 呼出関係取得
    '======================

    RowOut = 2

    For Each VBComp In gContext.TargetBook.VBProject.VBComponents

        If VBComp.Type = 1 Then
            
            If VBComp.CodeModule.CountOfLines > 0 Then
        
                CodeText = VBComp.CodeModule.Lines( _
                            1, _
                            VBComp.CodeModule.CountOfLines)

                RowOut = TraceModule( _
                            CodeText, _
                            VBComp.Name, _
                            FuncList, _
                            wsOut, _
                            RowOut)
            End If
        
        End If

    Next VBComp

    Call SetupTraceValidation(wsOut)
    
    wsOut.Columns.AutoFit

End Sub

'==================================================
' 関数一覧取得
'==================================================

' @JPName
' 関数一覧取得
'
' @Category
' FunctionTrace
'
' @Input
' CodeText(String)
' ModuleName(String)
' FuncList(Dictionary)
'
' @Output
' FuncList
'
' @Summary
' ソースコードから関数一覧を取得し
' 関数情報を収集する
'
Private Sub GetFunctions( _
            ByVal CodeText As String, _
            ByVal ModuleName As String, _
            ByRef FuncList As Object)

    Dim RegEx As Object
    Dim Matches As Object
    Dim M As Object
    Dim FuncName As String
    Dim Proc As clsProcInfo

    Set RegEx = CreateObject("VBScript.RegExp")

    RegEx.Global = True

    'RegEx.Pattern = _
    '    "(Public|Private|Friend)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)"
    
    RegEx.Pattern = _
        "(Public|Private|Friend)?\s*" & _
        "(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+" & _
        "([A-Za-z0-9_]+)"

    Set Matches = RegEx.Execute(CodeText)

    For Each M In Matches

        FuncName = M.SubMatches(2)

        If InStr(1, M.Value, _
                 "Declare", _
                 vbTextCompare) > 0 Then

            GoTo NextMatch

        End If
        
        Select Case UCase(FuncName)

            Case "PUBLIC", _
                 "PRIVATE", _
                 "FRIEND", _
                 "SUB", _
                 "FUNCTION", _
                 "IF", _
                 "END", _
                 "THEN", _
                 "ELSE", _
                 "DIM", _
                 "CALL"
                 
                ' 登録しない
                
            Case "ISSUE"
            
                ' GitHub Issue 等の誤検出防止

            Case Else

                If Not FuncList.Exists(FuncName) Then

                    Set Proc = New clsProcInfo

                    Proc.ProcName = FuncName
                    Proc.ModuleName = ModuleName

                    Proc.JPName = _
                        GetTagValueFromCode( _
                            CodeText, _
                            FuncName, _
                            "@JPName")

                    Proc.Category = _
                        GetTagValueFromCode( _
                            CodeText, _
                            FuncName, _
                            "@Category")

                    Proc.Summary = _
                        GetTagValueFromCode( _
                            CodeText, _
                            FuncName, _
                            "@Summary")

                    FuncList.Add FuncName, Proc

                End If

        End Select
    
NextMatch:

    Next M

End Sub

'==================================================
' トレース解析
'==================================================
' @JPName
' 関数呼出解析
'
' @Category
' FunctionTrace
'
' @Input
' CodeText(String)
' ModuleName(String)
' FuncList(Dictionary)
' wsOut(Worksheet)
' RowOut(Long)
'
' @Output
' 次出力行番号
'
' @Summary
' モジュール内の関数呼出関係を解析し
' 関数トレースシートへ出力する
'
Private Function TraceModule( _
            ByVal CodeText As String, _
            ByVal ModuleName As String, _
            ByRef FuncList As Object, _
            ByRef wsOut As Worksheet, _
            ByVal RowOut As Long) As Long

    Dim Lines() As String

    Dim CurrentProc As String

    Dim i As Long

    Dim Key As Variant

    Lines = Split(CodeText, vbCrLf)

    For i = LBound(Lines) To UBound(Lines)

        CurrentProc = GetProcedureName( _
                        Lines(i), _
                        CurrentProc)

        If CurrentProc <> "" Then

            For Each Key In FuncList.Keys

                If IsFunctionCall( _
                        Lines(i), _
                        CStr(Key)) Then

                    If UCase(Key) <> _
                       UCase(CurrentProc) Then
                       
                        If FuncList.Exists(CurrentProc) Then
                            FuncList(CurrentProc).AddCall Key
                        End If

                        If FuncList.Exists(Key) Then
                            FuncList(Key).AddCalledBy CurrentProc
                        End If
                        
                        wsOut.Cells(RowOut, 1) = CurrentProc

                        If FuncList.Exists(CurrentProc) Then
                            wsOut.Cells(RowOut, 2) = _
                                FuncList(CurrentProc).JPName
                        Else
                            wsOut.Cells(RowOut, 2) = ""
                        End If

                        wsOut.Cells(RowOut, 3) = Key

                        wsOut.Cells(RowOut, 4) = _
                            FuncList(Key).JPName

                        ' 呼出元モジュール
                        wsOut.Cells(RowOut, 5) = ModuleName

                        ' 呼出先モジュール
                        wsOut.Cells(RowOut, 6) = _
                            FuncList(Key).ModuleName

                        wsOut.Cells(RowOut, 7) = _
                            FuncList(Key).Category

                        wsOut.Cells(RowOut, 8) = _
                            FuncList(Key).Summary

                        wsOut.Cells(RowOut, 9) = i + 1

                        wsOut.Cells(RowOut, 10) = _
                            Trim$(Lines(i))

                        RowOut = RowOut + 1

                    End If

                End If

            Next Key

        End If

    Next i

    TraceModule = RowOut

End Function

' @JPName
' 現在関数取得
'
' @Category
' FunctionTrace
'
' @Input
' LineText(String)
' CurrentName(String)
'
' @Output
' ProcedureName(String)
'
' @Summary
' ソース行から現在解析中の
' プロシージャ名を取得する
'
Private Function GetProcedureName( _
            ByVal LineText As String, _
            ByVal CurrentName As String) As String

    If IsProcedureLine(LineText) Then
    
        GetProcedureName = _
            ExtractProcedureName(LineText)
    Else

        GetProcedureName = CurrentName

    End If

End Function

'==================================================
' 出力シート設定
'==================================================

' @JPName
' トレース確認設定
'
' @Category
' FunctionTrace
'
' @Summary
' 関数トレースシートへ
' 入力規則と条件付き書式を設定する
'
Private Sub SetupTraceValidation(ByVal ws As Worksheet)

    Dim LastRow As Long

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' 見出し
    ws.Range("K1").Value = "確認"
    ws.Range("L1").Value = "コメント"

    If LastRow < 2 Then Exit Sub

    ' 入力規則
    With ws.Range("K2:K" & LastRow).Validation
    
        On Error Resume Next
        
        .Delete
        
        'Debug.Print "Delete Err=" & Err.Number & " " & Err.Description
        
        On Error GoTo 0
        
        Err.Clear

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:="○ 正常,× 誤検出,△ 要確認,除外"

        'Debug.Print "Add Err=" & Err.Number & " " & Err.Description
        
        .IgnoreBlank = True
        .InCellDropdown = True

    End With


    ' 条件付き書式クリア
    ws.Range("A2:L" & LastRow).FormatConditions.Delete

    ' ○ 正常（薄緑）
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""○ 正常""", _
        RGB(198, 239, 206)

    ' × 誤検出（薄赤）
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""× 誤検出""", _
        RGB(255, 199, 206)

    ' △ 要確認（薄黄）
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""△ 要確認""", _
        RGB(255, 235, 156)

    ' 除外（薄灰）
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""除外""", _
        RGB(217, 217, 217)

End Sub

' @JPName
' 条件色設定
'
' @Category
' FunctionTrace
'
' @Input
' rng(Range)
' Formula(String)
' FillColor(Long)
'
' @Output
' なし
'
' @Summary
' 条件付き書式の色設定を追加する
'
Private Sub AddConditionColor( _
            ByVal rng As Range, _
            ByVal Formula As String, _
            ByVal FillColor As Long)

    Dim fc As FormatCondition

    Set fc = rng.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:=Formula)

    fc.Interior.Color = FillColor

End Sub

