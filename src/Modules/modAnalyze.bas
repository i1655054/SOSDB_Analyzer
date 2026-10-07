Attribute VB_Name = "modAnalyze"
Option Explicit

'==================================================
' ProcList定数
'==================================================

' 使用状況
Private Const STATUS_UNUSED As String = "削除候補"
Private Const STATUS_CONFIRM As String = "要確認"
Private Const STATUS_EVENT As String = "イベント"
Private Const STATUS_USED As String = "使用中"

' 表示色

' 削除候補
Private Const COLOR_UNUSED As Long = &HC8C8FF

' 要確認
Private Const COLOR_CONFIRM As Long = &HC8FFFF

' イベント
Private Const COLOR_EVENT As Long = &HFFC8C8

' 関数名抽出失敗時
Private Const PROC_UNKNOWN As String = "(Unknown)"
    
'==================================================
' ProcList生成
'==================================================

' @JPName
' 関数一覧作成
'
' @Category
' Analyze
'
' @Input
' なし
'
' @Output
' ProcListシート
'
' @Summary
' 対象ブックのプロシージャ情報を解析し
' ProcListシートへ出力する
'
' @Remarks
' UsageCountおよび使用状況を
' あわせて出力する
'
Public Sub ProcedureList()
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

    
    Dim VBComp As Object
    Dim CodeMod As Object
    
    Dim i As Long
    Dim r As Long
    Dim row As Long
    
    Dim LastRow As Long
    
    Dim ws As Worksheet
    
    Dim line As String
    Dim ProcName As String
    
    '------------------------------------
    ' シート準備
    '------------------------------------
    On Error Resume Next
    
    Set ws = _
        gContext.ManagementBook.Worksheets(SHEET_PROC_LIST)
    
    On Error GoTo 0
    
    If ws Is Nothing Then
    
        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                    gContext.ManagementBook.Worksheets.Count))
        
        ws.Name = SHEET_PROC_LIST
    
    End If
    
    ws.Cells.Clear
    
    '------------------------------------
    ' ヘッダ
    '------------------------------------
    ws.Range("A1:L1").Value = Array( _
        "Module", _
        "Type", _
        "ProcedureName", _
        "JPName", _
        "Category", _
        "Input", _
        "Output", _
        "Summary", _
        "Remarks", _
        "Declaration", _
        "UsageCount", _
        "判定")
    
    row = 2
    
    '------------------------------------
    ' 関数抽出
    '------------------------------------
    For Each VBComp In gContext.TargetBook.VBProject.VBComponents
        
        Set CodeMod = VBComp.CodeModule
        
        For i = 1 To CodeMod.CountOfLines
            
            line = Trim(CodeMod.Lines(i, 1))
            
            If IsProcedureLine(line) Then
                
                ProcName = ExtractProcedureName(line)
                
                If Trim(ProcName) = "" Then
                    ProcName = PROC_UNKNOWN
                End If
                
                ws.Cells(row, 1).Value = VBComp.Name
                ws.Cells(row, 2).Value = GetProcedureType(line)
                ws.Cells(row, 3).Value = ProcName

                ws.Cells(row, 4).Value = _
                    GetTagValue(CodeMod, i, "@JPName")

                ws.Cells(row, 5).Value = _
                    GetTagValue(CodeMod, i, "@Category")

                ws.Cells(row, 6).Value = _
                    GetTagValue(CodeMod, i, "@Input")

                ws.Cells(row, 7).Value = _
                    GetTagValue(CodeMod, i, "@Output")

                ws.Cells(row, 8).Value = _
                    GetTagValue(CodeMod, i, "@Summary")

                ws.Cells(row, 9).Value = _
                    GetTagValue(CodeMod, i, "@Remarks")

                ws.Cells(row, 10).Value = line

                ' 使用回数
                If ProcName = PROC_UNKNOWN Then
                    ws.Cells(row, 11).Value = 0
                Else
                    ws.Cells(row, 11).Value = CountUsage(ProcName)
                End If
                
                ' 判定
                ws.Cells(row, 12).Value = _
                    GetUsageStatus(ProcName, ws.Cells(row, 11).Value)
                
                row = row + 1
                
            End If
            
        Next i
        
    Next VBComp
    
    '------------------------------------
    ' 色付け
    '------------------------------------
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For r = 2 To LastRow
        
        Select Case ws.Cells(r, 12).Value
            
            Case STATUS_UNUSED
                ws.Rows(r).Interior.Color = COLOR_UNUSED
                
            Case STATUS_CONFIRM
                ws.Rows(r).Interior.Color = COLOR_CONFIRM
                
            Case STATUS_EVENT
                ws.Rows(r).Interior.Color = COLOR_EVENT
                
        End Select
        
    Next r
    
    ' 見た目
    ws.Columns.AutoFit
    ws.Rows(1).AutoFilter
    ws.Rows(1).Font.Bold = True
    

End Sub

'==================================================
' プロシージャ解析
'==================================================
' @JPName
' プロシージャ宣言判定
'
' @Category
' Analyze
'
' @Summary
' 指定行がプロシージャ宣言行か判定する
'
Public Function IsProcedureLine(line As String) As Boolean

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
    Or line Like "Property Set *" _
    Or line Like "Friend Sub *" _
    Or line Like "Friend Function *" Then
        IsProcedureLine = True
    Else
        IsProcedureLine = False
    End If

End Function

' @JPName
' プロシージャ種別取得
'
' @Category
' Analyze
'
' @Input
' line(String)
'
' @Output
' ProcedureType(String)
'
' @Summary
' プロシージャ宣言行から
' Sub、Function、Property種別を取得する
'
' @Remarks
' 宣言文字列の部分一致により判定する
'
Private Function GetProcedureType(line As String) As String

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

' @JPName
' 関数名抽出
'
' @Category
' Analyze
'
' @Summary
' プロシージャ宣言行から
' プロシージャ名を抽出する
'
Public Function ExtractProcedureName(line As String) As String

    Dim tmp As String
    
    tmp = line
    tmp = Replace(tmp, "Public ", "")
    tmp = Replace(tmp, "Private ", "")
    tmp = Replace(tmp, "Sub ", "")
    tmp = Replace(tmp, "Function ", "")
    tmp = Replace(tmp, "Property Get ", "")
    tmp = Replace(tmp, "Property Let ", "")
    tmp = Replace(tmp, "Property Set ", "")
    tmp = Replace(tmp, "Friend ", "")
    
    If InStr(tmp, "(") > 0 Then
        tmp = Left(tmp, InStr(tmp, "(") - 1)
    End If
    
    ExtractProcedureName = Trim(tmp)

End Function

'==================================================
' 使用状況判定
'==================================================

' @JPName
' 関数使用回数取得
'
' @Category
' Analyze
'
' @Summary
' 対象ブック内のプロシージャ呼出回数を集計する
'
' @Remarks
' プロシージャ宣言行は使用回数に含めない
' IsFunctionCallを利用して呼出判定を行う
'
Public Function CountUsage( _
            ByVal ProcName As String) _
            As Long

    Dim TargetBook As Workbook

    Set TargetBook = gContext.TargetBook

    If TargetBook Is Nothing Then Exit Function

    Dim VBComp As Object
    Dim CodeMod As Object
    Dim i As Long
    Dim Count As Long
    Dim LineText As String

    Count = 0

    For Each VBComp In TargetBook.VBProject.VBComponents
    
        Set CodeMod = VBComp.CodeModule

        For i = 1 To CodeMod.CountOfLines

            LineText = _
                Trim$(CodeMod.Lines(i, 1))
                
            If IsFunctionCall( _
                LineText, ProcName) Then
                
                Count = Count + 1

            End If

        Next i

    Next VBComp

    ' 宣言行を除外
    If Count <= 1 Then
        CountUsage = 0
    Else
        CountUsage = Count - 1
    End If

End Function

' @JPName
' イベント関数判定
'
' @Category
' Analyze
'
' @Summary
' プロシージャ名からイベント関数か判定する
'
Private Function IsEventProcedure(ProcName As String) As Boolean

    If ProcName Like "Workbook_*" _
    Or ProcName Like "Worksheet_*" _
    Or ProcName Like "UserForm_*" _
    Or ProcName Like "*_Click" _
    Or ProcName Like "*_Change" Then
    
        IsEventProcedure = True
    Else
        IsEventProcedure = False
    End If

End Function

' @JPName
' 使用状況判定
'
' @Category
' Analyze
'
' @Input
' ProcName(String)
' usage(Long)
'
' @Output
' Status(String)
'
' @Summary
' 使用回数およびイベント種別から
' 使用状況を判定する
'
' @Remarks
' イベント関数は使用回数に関係なく
' 「イベント」を返す
'
Private Function GetUsageStatus(ProcName As String, usage As Long) As String

    If IsEventProcedure(ProcName) Then
        GetUsageStatus = STATUS_EVENT
        
    ElseIf usage = 0 Then
        GetUsageStatus = STATUS_UNUSED
        
    ElseIf usage = 1 Then
        GetUsageStatus = STATUS_CONFIRM
        
    Else
        GetUsageStatus = STATUS_USED
        
    End If

End Function
