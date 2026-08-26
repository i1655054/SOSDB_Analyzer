Attribute VB_Name = "modFunctionTrac"
Option Explicit

' 関数トレース生成

Public Sub CreateFunctionTrace()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetTargetBook()

    If Not CheckTargetBook(TargetBook) Then Exit Sub
    
    Dim wsOut As Worksheet

    On Error Resume Next
    Set wsOut = Worksheets("関数トレース")
    On Error GoTo 0

    If wsOut Is Nothing Then

        Set wsOut = Worksheets.Add

        wsOut.Name = "関数トレース"

    Else

        wsOut.Cells.Clear

    End If

    wsOut.Range("A1") = "呼出元"
    wsOut.Range("B1") = "呼出先"
    wsOut.Range("C1") = "モジュール"
    wsOut.Range("D1") = "行番号"
    wsOut.Range("E1") = "元コード"

    Dim VBComp As Object
    Dim CodeText As String

    Dim FuncList As Object

    Set FuncList = CreateObject("Scripting.Dictionary")

    '======================
    ' 関数一覧取得
    '======================

    For Each VBComp In TargetBook.VBProject.VBComponents

        If VBComp.Type = 1 Then

            ' Debug.Print VBComp.Name

            If VBComp.CodeModule.CountOfLines > 0 Then

                CodeText = VBComp.CodeModule.Lines( _
                            1, _
                            VBComp.CodeModule.CountOfLines)

                Call GetFunctions(CodeText, FuncList)

            End If

        End If

    Next


    Dim Key As Variant

    For Each Key In FuncList.Keys

        'Debug.Print Key

    Next Key

    '======================
    ' 呼出関係取得
    '======================

    Dim RowOut As Long

    RowOut = 2

    For Each VBComp In TargetBook.VBProject.VBComponents

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

    MsgBox "関数トレース生成完了"

End Sub


Private Sub GetFunctions( _
            ByVal CodeText As String, _
            ByRef FuncList As Object)

    Dim RegEx As Object
    Dim Matches As Object
    Dim M As Object

    Set RegEx = CreateObject("VBScript.RegExp")

    RegEx.Global = True

    RegEx.Pattern = _
        "(Public|Private|Friend)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)"

    Set Matches = RegEx.Execute(CodeText)

    For Each M In Matches

        Dim FuncName As String

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
            
                '除外

            Case Else

                If Not FuncList.Exists(FuncName) Then

                    FuncList.Add FuncName, True

                End If

        End Select
    
NextMatch:

    Next

End Sub

Private Function TraceModule( _
            ByVal CodeText As String, _
            ByVal ModuleName As String, _
            ByRef FuncList As Object, _
            ByRef wsOut As Worksheet, _
            ByVal RowOut As Long) As Long

    Dim Lines() As String

    Lines = Split(CodeText, vbCrLf)

    Dim CurrentProc As String

    Dim i As Long

    For i = LBound(Lines) To UBound(Lines)

        CurrentProc = GetProcedureName( _
                        Lines(i), _
                        CurrentProc)

        If CurrentProc <> "" Then

            Dim Key As Variant

            For Each Key In FuncList.Keys

                If IsFunctionCall( _
                        Lines(i), _
                        CStr(Key)) Then

                    If UCase(Key) <> _
                       UCase(CurrentProc) Then

                        wsOut.Cells(RowOut, 1) = CurrentProc
                        wsOut.Cells(RowOut, 2) = Key
                        wsOut.Cells(RowOut, 3) = ModuleName
                        wsOut.Cells(RowOut, 4) = i + 1
                        wsOut.Cells(RowOut, 5) = Trim$(Lines(i))

                        RowOut = RowOut + 1

                    End If

                End If

            Next Key

        End If

    Next i

    TraceModule = RowOut

End Function


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

Private Function IsFunctionCall( _
        ByVal LineText As String, _
        ByVal FuncName As String) As Boolean

    Dim RegEx As Object

    LineText = Trim$(LineText)

    ' コメント行除外
    If Left$(LineText, 1) = "'" Then Exit Function

    ' API説明行除外
    If Left$(LineText, 1) = "■" Then Exit Function

    Set RegEx = CreateObject("VBScript.RegExp")

    RegEx.IgnoreCase = True

    ' Call Function
    RegEx.Pattern = _
        "(^|\s)Call\s+" & FuncName & _
        "(\s*\(|\s|$)"

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    ' Function(...)
    RegEx.Pattern = _
        "(^|[^A-Za-z0-9_])" & FuncName & _
        "\s*\("

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    ' Function (...)
    RegEx.Pattern = _
        "(^|[^A-Za-z0-9_])" & FuncName & _
        "\s+\("

    IsFunctionCall = RegEx.Test(LineText)

End Function

Private Sub SetupTraceValidation(ByVal ws As Worksheet)

    Dim LastRow As Long

    LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    ' 見出し
    ws.Range("F1").Value = "確認"
    ws.Range("G1").Value = "コメント"

    ' 入力規則
    With ws.Range("F2:F" & LastRow).Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:="○ 正常,× 誤検出,△ 要確認,除外"

        .IgnoreBlank = True
        .InCellDropdown = True

    End With

    ' 条件付き書式クリア
    ws.Range("A2:G" & LastRow).FormatConditions.Delete

    ' ○ 正常（薄緑）
    AddConditionColor _
        ws.Range("A2:G" & LastRow), _
        "=$F2=""○ 正常""", _
        RGB(198, 239, 206)

    ' × 誤検出（薄赤）
    AddConditionColor _
        ws.Range("A2:G" & LastRow), _
        "=$F2=""× 誤検出""", _
        RGB(255, 199, 206)

    ' △ 要確認（薄黄）
    AddConditionColor _
        ws.Range("A2:G" & LastRow), _
        "=$F2=""△ 要確認""", _
        RGB(255, 235, 156)

    ' 除外（薄灰）
    AddConditionColor _
        ws.Range("A2:G" & LastRow), _
        "=$F2=""除外""", _
        RGB(217, 217, 217)

End Sub

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

