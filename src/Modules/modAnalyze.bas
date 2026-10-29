Attribute VB_Name = "modAnalyze"
Option Explicit

'==================================================
' ProcList定数
'==================================================

' 使用状況
' 未使用関数
Private Const STATUS_UNUSED As String = "削除候補"

' 使用状況確認対象
Private Const STATUS_CONFIRM As String = "要確認"

' イベントプロシージャ
Private Const STATUS_EVENT As String = "イベント"

' 使用中プロシージャ
Private Const STATUS_USED As String = "使用中"

'==================================================
' 表示色
'==================================================

' 未使用関数行色
Private Const COLOR_UNUSED As Long = &HC8C8FF

' 要確認関数行色
Private Const COLOR_CONFIRM As Long = &HC8FFFF

' イベント関数行色
Private Const COLOR_EVENT As Long = &HFFC8C8

' 関数名抽出失敗時の表示名
Private Const PROC_UNKNOWN As String = "(Unknown)"
    
'==================================================
' 公開API
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

    
    ' VBAコンポーネント
    Dim VBComp As Object
    
    ' コードモジュール
    Dim CodeMod As Object
    
    ' 行走査用Index
    Dim i As Long
    
    ' 色付け用行番号
    Dim r As Long
    
    ' 出力行番号
    Dim row As Long
    
    ' ProcList最終行
    Dim LastRow As Long
    
    ' ProcListシート
    Dim ws As Worksheet
    
    ' 処理対象行
    Dim line As String
    
    ' プロシージャ名
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
                
                Debug.Print "ProcName : " & VBComp.Name & "." & ProcName
                
                ws.Cells(row, COL_MODULE).Value = VBComp.Name
                ws.Cells(row, COL_TYPE).Value = GetProcedureType(line)
                ws.Cells(row, COL_PROC).Value = ProcName

                ws.Cells(row, COL_JPNAME).Value = _
                    GetTagValue(CodeMod, i, "@JPName")

                ws.Cells(row, COL_CATEGORY).Value = _
                    GetTagValue(CodeMod, i, "@Category")

                ws.Cells(row, COL_INPUT).Value = _
                    GetTagValue(CodeMod, i, "@Input")

                ws.Cells(row, COL_OUTPUT).Value = _
                    GetTagValue(CodeMod, i, "@Output")

                ws.Cells(row, COL_SUMMARY).Value = _
                    GetTagValue(CodeMod, i, "@Summary")

                ws.Cells(row, COL_REMARKS).Value = _
                    GetTagValue(CodeMod, i, "@Remarks")

                ws.Cells(row, COL_DECLARATION).Value = line

                ' 使用回数
                If ProcName = PROC_UNKNOWN Then
                    ws.Cells(row, COL_USAGECOUNT).Value = 0
                Else
                    ws.Cells(row, COL_USAGECOUNT).Value = CountUsage(ProcName)
                End If
                
                ' 判定
                ws.Cells(row, COL_STATUS).Value = _
                    GetUsageStatus(ProcName, ws.Cells(row, COL_USAGECOUNT).Value)
                
                row = row + 1
                
            End If
            
        Next i
        
    Next VBComp
    
    '------------------------------------
    ' 色付け
    '------------------------------------
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For r = 2 To LastRow
        
        Select Case ws.Cells(r, COL_STATUS).Value
            
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


' @JPName
' 未使用関数一覧出力
'
' @Category
' Analyze
'
' @Input
' なし
'
' @Output
' UnusedProcListシート
'
' @Summary
' ProcListから未使用関数を抽出し
' UnusedProcListシートへ出力する
'
' @Remarks
' 判定が「削除候補」の関数のみ出力する
'
Public Function ExportUnusedProcedureList() As Long

    ' ProcListシート
    Dim wsProc As Worksheet
    
    ' 出力先シート
    Dim wsOut As Worksheet

    ' ProcList最終行
    Dim LastRow As Long
    
    ' 出力行番号
    Dim OutRow As Long
    
    ' 行走査用Index
    Dim r As Long

    On Error Resume Next

    Set wsOut = _
        gContext.ManagementBook.Worksheets( _
            SHEET_UNUSED_PROC_LIST)

    On Error GoTo 0

    If wsOut Is Nothing Then

        Set wsOut = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                    gContext.ManagementBook.Worksheets.Count))

        wsOut.Name = _
            SHEET_UNUSED_PROC_LIST

    End If

    Set wsProc = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    wsOut.Cells.Clear

    '------------------------------------
    ' ヘッダ
    '------------------------------------
    wsOut.Range("A1:G1").Value = Array( _
        "Module", _
        "Type", _
        "ProcedureName", _
        "JPName", _
        "Category", _
        "UsageCount", _
        "判定")

    OutRow = 2

    LastRow = _
        wsProc.Cells( _
            wsProc.Rows.Count, COL_MODULE) _
            .End(xlUp).row

    '------------------------------------
    ' 未使用関数抽出
    '------------------------------------
    For r = 2 To LastRow

        If wsProc.Cells(r, COL_STATUS).Value = _
            STATUS_UNUSED Then

            wsOut.Cells(OutRow, COL_MODULE).Value = _
                wsProc.Cells(r, COL_MODULE).Value

            wsOut.Cells(OutRow, COL_TYPE).Value = _
                wsProc.Cells(r, COL_TYPE).Value

            wsOut.Cells(OutRow, COL_PROC).Value = _
                wsProc.Cells(r, COL_PROC).Value

            wsOut.Cells(OutRow, COL_JPNAME).Value = _
                wsProc.Cells(r, COL_JPNAME).Value

            wsOut.Cells(OutRow, COL_CATEGORY).Value = _
                wsProc.Cells(r, COL_CATEGORY).Value

            wsOut.Cells(OutRow, COL_INPUT).Value = _
                wsProc.Cells(r, COL_USAGECOUNT).Value

            wsOut.Cells(OutRow, COL_OUTPUT).Value = _
                wsProc.Cells(r, COL_STATUS).Value

            OutRow = OutRow + 1

        End If

    Next r

    ExportUnusedProcedureList = OutRow - 2
    
    '------------------------------------
    ' 書式
    '------------------------------------
    wsOut.Columns.AutoFit

    wsOut.Rows(1).Font.Bold = True

    wsOut.Rows(1).AutoFilter
    
End Function


'==================================================
' プロシージャ解析
'==================================================

' @JPName
' プロシージャ定義判定
'
' @Category
' Analyze
'
' @Input
' LineText(String)
'
' @Output
' True:プロシージャ定義行
' False:プロシージャ定義行以外
'
' @Summary
' 指定行がプロシージャの
' 定義行であるか判定する
'
' @Remarks
' Public Sub、Private Sub、
' Public Function、Private Function、
' Property Let/Get/Set などの
' プロシージャ定義を判定対象とする
'
Public Function IsProcedureLine( _
                ByVal line As String) _
                As Boolean

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
' プロシージャ名抽出
'
' @Category
' Analyze
'
' @Input
' LineText(String)
'
' @Output
' ProcedureName(String)
'
' @Summary
' プロシージャ定義行から
' プロシージャ名を抽出する
'
' @Remarks
' Sub、Function、Propertyの
' 定義行を解析対象とする
' プロシージャ名を取得できない場合は
' 空文字列を返す
'
Public Function ExtractProcedureName(line As String) As String

    ' 加工用文字列
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

'==================================================
' 使用状況判定
'==================================================

' @JPName
' 使用回数集計
'
' @Category
' Analyze
'
' @Input
' ProcName(String)
'
' @Output
' UsageCount(Long)
'
' @Summary
' 指定プロシージャの
' 呼出回数を取得する
'
' @Remarks
' 関数トレース情報または
' 関数依存関係情報を参照して
' 呼出回数を集計する
' 対象が存在しない場合は
' 0を返す
'
Public Function CountUsage( _
            ByVal ProcName As String) _
            As Long

    ' 解析対象ブック
    Dim TargetBook As Workbook

    Set TargetBook = gContext.TargetBook

    If TargetBook Is Nothing Then Exit Function

    ' VBAコンポーネント
    Dim VBComp As Object
    
    ' コードモジュール
    Dim CodeMod As Object
    
    ' 行番号
    Dim i As Long
    
    ' 呼出回数
    Dim Count As Long
    
    ' 判定対象行
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

    ' 宣言行は
    ' IsProcedureLineで除外済み
    CountUsage = Count

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
        
    ElseIf IsSpecialProcedure(ProcName) Then
    
        GetUsageStatus = STATUS_USED
       
    ElseIf usage = 0 Then
    
        GetUsageStatus = STATUS_UNUSED
        
    ElseIf usage = 1 Then
    
        GetUsageStatus = STATUS_CONFIRM
        
    Else
        GetUsageStatus = STATUS_USED
        
    End If

End Function

' @JPName
' イベントプロシージャ判定
'
' @Category
' Analyze
'
' @Input
' ProcName(String)
'
' @Output
' True : イベントプロシージャ
' False: 通常プロシージャ
'
' @Summary
' 指定プロシージャが
' イベントプロシージャか判定する
'
' @Remarks
' Workbook_Open、
' Worksheet_Change、
' CommandButton_Clickなどの
' イベントハンドラを判定対象とする
' イベントプロシージャの場合は
' Trueを返す
'
Private Function IsEventProcedure( _
                ByVal ProcName As String) _
                As Boolean

    IsEventProcedure = _
        ProcName Like "Workbook_*" _
        Or ProcName Like "Worksheet_*" _
        Or ProcName Like "UserForm_*" _
        Or ProcName Like "Class_*" _
        Or ProcName Like "*_Click" _
        Or ProcName Like "*_DblClick" _
        Or ProcName Like "*_Change" _
        Or ProcName Like "*_MouseDown" _
        Or ProcName Like "*_MouseUp" _
        Or ProcName Like "*_MouseMove" _
        Or ProcName Like "*_KeyDown" _
        Or ProcName Like "*_KeyUp" _
        Or ProcName Like "*_KeyPress" _
        Or ProcName Like "*_ItemCheck" _
        Or ProcName Like "*_ItemClick"

End Function

' @JPName
' 特殊プロシージャ判定
'
' @Category
' Analyze
'
' @Input
' ProcName(String)
'
' @Output
' True : 特殊プロシージャ
' False: 通常プロシージャ
'
' @Summary
' 指定プロシージャが
' 特殊プロシージャか判定する
'
' @Remarks
' 初期化処理や終了処理など
' 解析上で特別扱いする
' プロシージャを判定対象とする
' 特殊プロシージャの場合は
' Trueを返す
'
Private Function IsSpecialProcedure( _
                ByVal ProcName As String) _
                As Boolean

    Select Case ProcName

        Case "SelectNode"

            ' Shape.OnActionから起動
            IsSpecialProcedure = True

    End Select

End Function
