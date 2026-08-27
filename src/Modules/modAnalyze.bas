Attribute VB_Name = "modAnalyze"
Option Explicit

Public Sub ExtractProcedures_WithAnalysis()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)

    If Not CheckTargetBook(TargetBook) Then Exit Sub

    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim row As Long
    Dim ws As Worksheet
    
    '=========================================
    ' シート準備
    '=========================================
    On Error Resume Next
    Set ws = Worksheets("ProcList")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProcList"
    End If
    
    ws.Cells.Clear
    
    '=========================================
    ' ヘッダ
    '=========================================
    ws.Range("A1:F1").Value = Array( _
        "Module", _
        "Type", _
        "ProcedureName", _
        "Declaration", _
        "UsageCount", _
        "判定")
    
    row = 2
    
    '=========================================
    ' 関数抽出
    '=========================================
    For Each VBComp In TargetBook.VBProject.VBComponents
        
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            
            Dim line As String
            line = Trim(codeMod.Lines(i, 1))
            
            If IsProcedureLine(line) Then
                
                Dim procName As String
                procName = ExtractProcedureName(line)
                
                If Trim(procName) = "" Then
                    procName = "(Unknown)"
                End If
                
                ws.Cells(row, 1).Value = VBComp.Name
                ws.Cells(row, 2).Value = GetProcedureType(line)
                ws.Cells(row, 3).Value = procName
                ws.Cells(row, 4).Value = line
                
                ' 使用回数
                If procName = "(Unknown)" Then
                    ws.Cells(row, 5).Value = 0
                Else
                    ws.Cells(row, 5).Value = CountUsage(procName)
                End If
                
                ' 判定
                ws.Cells(row, 6).Value = GetUsageStatus(procName, ws.Cells(row, 5).Value)
                
                row = row + 1
                
            End If
            
        Next i
        
    Next VBComp
    
    '=========================================
    ' 色付け
    '=========================================
    Dim LastRow As Long
    Dim r As Long
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For r = 2 To LastRow
        
        Select Case ws.Cells(r, 6).Value
            
            Case "削除候補"
                ws.Rows(r).Interior.Color = RGB(255, 200, 200)
                
            Case "要確認"
                ws.Rows(r).Interior.Color = RGB(255, 255, 200)
                
            Case "イベント"
                ws.Rows(r).Interior.Color = RGB(200, 200, 255)
                
        End Select
        
    Next r
    
    ' 見た目
    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
    ws.Rows(1).Font.Bold = True
    
    MsgBox "解析完了"

End Sub

Sub JumpToProcedure()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)

    If Not CheckTargetBook(TargetBook) Then Exit Sub
    
    Dim procName As String
    procName = InputBox("ジャンプする関数名を入力")
    
    If procName = "" Then Exit Sub
    
    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    
    For Each VBComp In TargetBook.VBProject.VBComponents
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            
            If InStr(codeMod.Lines(i, 1), procName) > 0 Then
                
                codeMod.CodePane.Show
                codeMod.CodePane.SetSelection i, 1, i, 1
                
                MsgBox "ジャンプしました：" & VBComp.Name
                Exit Sub
                
            End If
            
        Next i
    Next VBComp
    
    MsgBox "見つかりません"

End Sub

Sub JumpFromSelectedCell()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)

    If Not CheckTargetBook(TargetBook) Then Exit Sub
    
    Dim procName As String
    procName = ActiveCell.Value
    
    If procName = "" Then
        MsgBox "セルに関数名がありません"
        Exit Sub
    End If
    
    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim line As String
    
    For Each VBComp In TargetBook.VBProject.VBComponents
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            
            line = Trim(codeMod.Lines(i, 1))
            
            ' ★ 定義行だけを対象にする
            If line Like "*Sub " & procName & "*" _
            Or line Like "*Function " & procName & "*" _
            Or line Like "*Property Get " & procName & "*" _
            Or line Like "*Property Let " & procName & "*" _
            Or line Like "*Property Set " & procName & "*" Then
                
                codeMod.CodePane.Show
                codeMod.CodePane.SetSelection i, 1, i, 1
                
                Exit Sub
                
            End If
            
        Next i
        
    Next VBComp
    
    MsgBox "定義が見つかりません：" & procName

End Sub

' ① 未使用だけ別シートに抽出
Sub AnalyzeAndExtractUnused()

    Dim wsSrc As Worksheet
    Dim wsUnused As Worksheet
    Dim LastRow As Long
    Dim r As Long
    Dim newRow As Long
    
    Set wsSrc = Worksheets("ProcList")
    
    '=========================================
    ' 未使用シート作成
    '=========================================
    On Error Resume Next
    Set wsUnused = Worksheets("UnusedList")
    On Error GoTo 0
    
    If wsUnused Is Nothing Then
        Set wsUnused = Worksheets.Add
        wsUnused.Name = "UnusedList"
    End If
    
    wsUnused.Cells.Clear
    
    ' ヘッダ
    wsUnused.Range("A1:F1").Value = Array("Module", "Type", "ProcedureName", "Declaration", "Line", "UsageCount")
    
    newRow = 2
    
    LastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row
    
    '=========================================
    ' 未使用抽出（UsageCount <=1）
    '=========================================
    For r = 2 To LastRow
        
        If wsSrc.Cells(r, 6).Value <= 1 Then
            
            wsUnused.Cells(newRow, 1).Resize(1, 6).Value = _
                wsSrc.Cells(r, 1).Resize(1, 6).Value
            
            newRow = newRow + 1
        
        End If
        
    Next r
    
    wsUnused.Columns.AutoFit
    wsUnused.Rows(1).AutoFilter
    
    MsgBox "未使用関数 抽出完了"

End Sub

' ② 削除候補リスト自動生成
Sub ExtractDeleteCandidates()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long
    
    Set ws = Worksheets("UnusedList")
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For r = 2 To LastRow
        
        ' より厳格な条件
        If ws.Cells(r, 6).Value = 0 Then
            ws.Cells(r, 6).Interior.Color = RGB(255, 0, 0) ' 完全未使用
        ElseIf ws.Cells(r, 6).Value = 1 Then
            ws.Cells(r, 6).Interior.Color = RGB(255, 200, 200) ' 定義のみ
        End If
        
    Next r
    
    MsgBox "削除候補 強調完了"

End Sub

' ③ 一括削除支援
Sub GenerateDeleteScript()

    Dim ws As Worksheet
    Dim wsOut As Worksheet
    Dim LastRow As Long
    Dim r As Long
    Dim OutRow As Long
    
    Set ws = Worksheets("UnusedList")
    
    '=========================================
    ' 出力シート用意
    '=========================================
    On Error Resume Next
    Set wsOut = Worksheets("DeleteScript")
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = Worksheets.Add
        wsOut.Name = "DeleteScript"
    End If
    
    wsOut.Cells.Clear
    
    ' ヘッダ
    wsOut.Range("A1").Value = "削除候補プロシージャ"
    wsOut.Range("B1").Value = "モジュール"
    wsOut.Range("C1").Value = "UsageCount"
    
    OutRow = 2
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    '=========================================
    ' 1行ずつ出力
    '=========================================
    For r = 2 To LastRow
        
        If ws.Cells(r, 6).Value <= 1 Then
            
            wsOut.Cells(OutRow, 1).Value = ws.Cells(r, 3).Value ' ProcedureName
            wsOut.Cells(OutRow, 2).Value = ws.Cells(r, 1).Value ' Module
            wsOut.Cells(OutRow, 3).Value = ws.Cells(r, 6).Value ' Usage
            
            OutRow = OutRow + 1
            
        End If
        
    Next r
    
    ' 見やすく
    wsOut.Columns.AutoFit
    wsOut.Rows(1).Font.Bold = True
    wsOut.Rows(1).AutoFilter

    MsgBox "削除候補リスト作成完了"

End Sub


'分析用
Sub ExtractProcedures_WithUsage()
' "ProcList"シートが追加される
' 1行目
' Module  Type    ProcedureName   Declaration Line    UsageCount
'         種別                                行番号  使用回数
'
' UsageCount 色   意味
' 0          赤   完全未使用
' 1          赤   定義のみ
' 2          黄   要確認（1回しか呼ばれてない）
' 3+         通常 使用中


    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)
    
    If TargetBook Is Nothing Then Exit Sub
    
    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim row As Long
    Dim ws As Worksheet
    
    ' シート準備
    On Error Resume Next
    Set ws = Worksheets("ProcList")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProcList"
    End If
    
    ws.Cells.Clear
    
    ws.Range("A1:F1").Value = Array("Module", "Type", "ProcedureName", "Declaration", "Line", "UsageCount")
    
    row = 2
    
    For Each VBComp In TargetBook.VBProject.VBComponents
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            
            Dim line As String
            line = Trim(codeMod.Lines(i, 1))
            
            If IsProcedureLine(line) Then
                ws.Cells(row, 1).Value = VBComp.Name
                ws.Cells(row, 2).Value = GetProcedureType(line)
                ws.Cells(row, 3).Value = ExtractProcedureName(line)
                ws.Cells(row, 4).Value = line
                ws.Cells(row, 5).Value = i
                
                ' ここで自動的に式を入れる
                ws.Cells(row, 6).Formula = "=CountUsage(C" & row & ")"
                
                row = row + 1
            End If
            
        Next i
        
    Next VBComp
    
    '=========================================
    ' 使用回数に応じた色付け（ここが追加ポイント）
    '=========================================
    Dim LastRow As Long
    Dim r As Long

    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For r = 2 To LastRow
    
        If ws.Cells(r, 6).Value <= 1 Then
            ws.Cells(r, 6).Interior.Color = RGB(255, 200, 200) ' 赤（未使用）
        ElseIf ws.Cells(r, 6).Value = 2 Then
            ws.Cells(r, 6).Interior.Color = RGB(255, 255, 200) ' 黄（要確認）
        End If
    
    Next r

    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
    
    MsgBox "プロシージャ解析 + 使用回数完了"

End Sub

'純粋一覧
Sub ExtractProcedures_Advanced()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)
    
    If TargetBook Is Nothing Then Exit Sub
    
    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim row As Long
    Dim ws As Worksheet
    
    '=========================================
    ' シート取得（なければ作成）
    '=========================================
    On Error Resume Next
    Set ws = Worksheets("ProcList")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = "ProcList"
    End If
    
    '=========================================
    ' 初期化
    '=========================================
    ws.Cells.Clear
    
    ws.Cells(1, 1).Value = "Module"
    ws.Cells(1, 2).Value = "Type"
    ws.Cells(1, 3).Value = "ProcedureName"
    ws.Cells(1, 4).Value = "FullDeclaration"
    ws.Cells(1, 5).Value = "LineNo"
    ws.Cells(1, 6).Value = "UsageCheck"
    
    row = 2
    
    '=========================================
    ' モジュール巡回
    '=========================================
    For Each VBComp In TargetBook.VBProject.VBComponents
        
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            
            Dim line As String
            line = Trim(codeMod.Lines(i, 1))
            
            If IsProcedureLine(line) Then
                
                ws.Cells(row, 1).Value = VBComp.Name
                ws.Cells(row, 2).Value = GetProcedureType(line)
                ws.Cells(row, 3).Value = ExtractProcedureName(line)
                ws.Cells(row, 4).Value = line
                ws.Cells(row, 5).Value = i
                ws.Cells(row, 6).Value = "" ' 後で使用判定
                
                row = row + 1
                
            End If
            
        Next i
        
    Next VBComp
    
    '=========================================
    ' 見た目調整
    '=========================================
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
    
    MsgBox "関数一覧 出力完了", vbInformation

End Sub

Function IsProcedureLine(line As String) As Boolean

    ' コメント行は除外
    If Left(Trim(line), 1) = "'" Then
        IsProcedureLine = False
        Exit Function
    End If

    ' 宣言行だけ判定（先頭一致）
    line = Trim(line)
    
    If line Like "Sub *" _
    Or line Like "Public Sub *" _
    Or line Like "Private Sub *" _
    Or line Like "Function *" _
    Or line Like "Public Function *" _
    Or line Like "Private Function *" _
    Or line Like "Property Get *" _
    Or line Like "Property Let *" _
    Or line Like "Property Set *" Then
    
        IsProcedureLine = True
    Else
        IsProcedureLine = False
    End If

End Function

Function GetProcedureType(line As String) As String

    If InStr(line, "Sub") > 0 Then
        GetProcedureType = "Sub"
    ElseIf InStr(line, "Function") > 0 Then
        GetProcedureType = "Function"
    ElseIf InStr(line, "Property") > 0 Then
        GetProcedureType = "Property"
    Else
        GetProcedureType = ""
    End If

End Function

Function ExtractProcedureName(line As String) As String

    Dim tmp As String
    
    tmp = line
    tmp = Replace(tmp, "Public ", "")
    tmp = Replace(tmp, "Private ", "")
    tmp = Replace(tmp, "Sub ", "")
    tmp = Replace(tmp, "Function ", "")
    tmp = Replace(tmp, "Property Get ", "")
    tmp = Replace(tmp, "Property Let ", "")
    tmp = Replace(tmp, "Property Set ", "")
    
    If InStr(tmp, "(") > 0 Then
        tmp = Left(tmp, InStr(tmp, "(") - 1)
    End If
    
    ExtractProcedureName = Trim(tmp)

End Function

Function CountUsage(procName As String) As Long

    Dim TargetBook As Workbook
    
    Set TargetBook = GetWorkbookByWorkbookName(BOOK_SOSDB)
    
    If TargetBook Is Nothing Then Exit Function
    
    Dim VBComp As Object
    Dim codeMod As Object
    Dim i As Long
    Dim Count As Long
    
    If Trim(procName) = "" Then
        CountUsage = 0
        Exit Function
    End If
    
    Count = 0
    
    For Each VBComp In TargetBook.VBProject.VBComponents
        Set codeMod = VBComp.CodeModule
        
        For i = 1 To codeMod.CountOfLines
            If InStr(codeMod.Lines(i, 1), procName) > 0 Then
                Count = Count + 1
            End If
        Next i
        
    Next VBComp
    
    If Count <= 1 Then
        CountUsage = 0
    Else
        CountUsage = Count - 1 ' 自分自身除外
    End If

End Function

Function IsEventProcedure(procName As String) As Boolean

    If procName Like "Workbook_*" _
    Or procName Like "Worksheet_*" _
    Or procName Like "UserForm_*" _
    Or procName Like "*_Click" _
    Or procName Like "*_Change" Then
    
        IsEventProcedure = True
    Else
        IsEventProcedure = False
    End If

End Function

Function GetUsageStatus(procName As String, usage As Long) As String

    If IsEventProcedure(procName) Then
        GetUsageStatus = "イベント"
        
    ElseIf usage = 0 Then
        GetUsageStatus = "削除候補"
        
    ElseIf usage = 1 Then
        GetUsageStatus = "要確認"
        
    Else
        GetUsageStatus = "使用中"
        
    End If

End Function

