Attribute VB_Name = "modFunctionTrac"
Option Explicit

'===================================================================
'@JPName     関数トレース解析
'@Category   FunctionTrace
'@Summary    VBAソースコードの関数呼出関係を解析し
'            関数トレースシートを作成する
'===================================================================

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
' 解析対象モジュールを走査し
'
' 呼出元Procedure
' 呼出先Procedure
' 呼出元モジュール
' 呼出先モジュール
' Category
' Summary
'
' をFunctionTraceシートへ出力する
'
' 標準モジュール
' クラスモジュール
' UserForm
' シートモジュール
'
' を解析対象とする
'
' 出力後は確認支援用の
' 入力規則および条件付き書式を設定する
'
Public Sub FunctionTrace()
       
    ' 関数トレース出力シート
    Dim wsOut As Worksheet
    
    ' VBAコンポーネント
    Dim VBComp As Object
    
    ' Procedure情報一覧
    Dim FuncList As Object
    
    ' モジュールソースコード
    Dim CodeText As String

    ' 出力行番号
    Dim RowOut As Long
    
    ' 総モジュール数
    Dim TotalModuleCount As Long

    ' 現在モジュール数
    Dim CurrentModuleCount As Long

    If gContext Is Nothing Then Exit Sub
    
    On Error Resume Next
    
    '# FunctionTraceシート準備
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

    '# Procedure一覧管理初期化
    Set FuncList = CreateObject("Scripting.Dictionary")

    '======================
    '# 関数一覧取得
    '======================
    '# 解析対象モジュール数取得
    TotalModuleCount = _
        gContext.TargetBook.VBProject.VBComponents.Count

    CurrentModuleCount = 0

    '# Procedure一覧収集
    For Each VBComp In gContext.TargetBook.VBProject.VBComponents

        '# FunctionTrace対象判定
        If IsTraceTargetModule(VBComp) Then

            If VBComp.CodeModule.CountOfLines > 0 Then

                '# モジュールソース取得
                CodeText = VBComp.CodeModule.Lines( _
                            1, _
                            VBComp.CodeModule.CountOfLines)

                '# Procedure情報収集
                Call GetFunctions(CodeText, VBComp.Name, FuncList)

            End If

        End If

    Next

    '======================
    '# 呼出関係取得
    '======================

    RowOut = 2

    '# 呼出関係解析開始
    For Each VBComp In gContext.TargetBook.VBProject.VBComponents

        CurrentModuleCount = _
            CurrentModuleCount + 1
        
        '# 解析進捗表示更新
        Application.StatusBar = _
                "FunctionTrace解析中 [" & _
                CurrentModuleCount & "/" & _
                TotalModuleCount & _
                "] " & _
                VBComp.Name & _
                " Trace=" & RowOut

        DoEvents
        
        '# FunctionTrace対象判定
        If IsTraceTargetModule(VBComp) Then
            
            If VBComp.CodeModule.CountOfLines > 0 Then
        
                '# モジュールソース取得
                CodeText = VBComp.CodeModule.Lines( _
                            1, _
                            VBComp.CodeModule.CountOfLines)

                '# モジュール呼出関係解析
                RowOut = TraceModule( _
                            CodeText, _
                            VBComp.Name, _
                            FuncList, _
                            wsOut, _
                            RowOut)
            End If
        
        End If

    Next VBComp

    '# 確認支援設定
    Call SetupTraceValidation(wsOut)
    
    '# FunctionTraceシート体裁設定
    wsOut.Columns.AutoFit

End Sub


'===================================================================
' @JPName
' FunctionTrace対象判定
'
' @Category
' FunctionTrace
'
' @Input
' VBComp(Object)
'
' @Output
' Boolean
'
' @Summary
' FunctionTrace解析対象の
' VBAコンポーネントか判定する
'
' @Remarks
' ProcListと同一対象とするため
'
' 標準モジュール
' クラスモジュール
' UserForm
' シートモジュール
'
' を解析対象とする
'===================================================================
Private Function IsTraceTargetModule( _
                ByVal VBComp As Object) _
                As Boolean

    '-----------------------------
    '# コンポーネント種別判定
    '-----------------------------
    Select Case VBComp.Type

        '-------------------------
        '# 標準モジュール
        '-------------------------
        ' vbext_ct_StdModule = 1
        Case 1

            IsTraceTargetModule = True

        '-------------------------
        '# クラスモジュール
        '-------------------------
        ' vbext_ct_ClassModule = 2
        Case 2

            IsTraceTargetModule = True

        '-------------------------
        '# UserForm
        '-------------------------
        ' vbext_ct_MSForm = 3
        Case 3

            IsTraceTargetModule = True

        '-------------------------
        '# シート/ThisWorkbook
        '-------------------------
        ' vbext_ct_Document = 100
        Case 100

            IsTraceTargetModule = True

    End Select

End Function

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
' FuncList(Dictionary)
'
' @Summary
' ソースコードから関数一覧を取得し
' 関数情報を収集する
'
' @Remarks
' Procedure定義を走査し
' FuncListへ登録する
'
' JPName
' Category
' Summary
'
' のタグ情報を取得する
'
' 同名Procedureは
' 重複登録しない
'
Public Function GetFunctions( _
            ByVal CodeText As String, _
            ByVal ModuleName As String, _
            ByRef FuncList As Object)

    ' ソースコード行配列
    Dim Lines() As String
    
    ' 現在行
    Dim LineText As String

    ' 行走査用Index
    Dim i As Long
    
    ' Procedure名
    Dim FuncName As String

    ' Procedure情報
    Dim Proc As clsProcInfo

    '# ソースコード行分割
    Lines = Split(CodeText, vbCrLf)

    '# Procedure定義走査
    For i = LBound(Lines) To UBound(Lines)

        LineText = Trim$(Lines(i))

        '# Procedure定義判定
        If IsProcedureLine(LineText) Then

            '# Procedure名取得
            FuncName = _
                ExtractProcedureName(LineText)

            If FuncName <> "" Then

                '# 未登録Procedure判定
                If Not FuncList.Exists(FuncName) Then

                    '# Procedure情報生成
                    Set Proc = New clsProcInfo

                    
                    Proc.ProcName = FuncName
                    Proc.ModuleName = ModuleName

                    '# Procedureコメント情報取得
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

                    '# Procedure一覧登録
                    FuncList.Add _
                        FuncName, _
                        Proc

                End If

            End If

        End If

    Next i

End Function


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
' NextRow(Long)
'
' @Summary
' モジュール内の関数呼出関係を解析し
' 関数トレースシートへ出力する
'
' @Remarks
' モジュール単位で
' FunctionCallを解析する
'
' 呼出元Procedureおよび
' 呼出先Procedureを特定し
'
' FunctionTraceシートへ
' 呼出関係を出力する
'
' 自己呼出は出力対象外とする
'
Private Function TraceModule( _
            ByVal CodeText As String, _
            ByVal ModuleName As String, _
            ByRef FuncList As Object, _
            ByRef wsOut As Worksheet, _
            ByVal RowOut As Long) As Long

    ' ソースコード行配列
    Dim Lines() As String

    ' 現在解析中Procedure
    Dim CurrentProc As String

    ' 行走査用Index
    Dim i As Long

    ' Procedure一覧走査用
    Dim Key As Variant

    '# ソースコード行分割
    Lines = Split(CodeText, vbCrLf)

    '# モジュール解析開始
    For i = LBound(Lines) To UBound(Lines)

        '# 現在Procedure取得
        CurrentProc = GetProcedureName( _
                        Lines(i), _
                        CurrentProc)

        If CurrentProc <> "" Then

            '# 呼出先Procedure探索
            For Each Key In FuncList.Keys

                '# Procedure呼出判定
                If IsFunctionCall( _
                        Lines(i), _
                        CStr(Key)) Then

                    '# 自己呼出除外
                    If UCase(Key) <> _
                       UCase(CurrentProc) Then
                       
                        '# 呼出関係登録
                        If FuncList.Exists(CurrentProc) Then
                            FuncList(CurrentProc).AddCall Key
                        End If

                        '# 被呼出関係登録
                        If FuncList.Exists(Key) Then
                            FuncList(Key).AddCalledBy CurrentProc
                        End If
                        
                        '# FunctionTrace出力
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

                        '# 次出力行へ移動
                        RowOut = RowOut + 1

                    End If

                End If

            Next Key

        End If

    Next i

    '# 次出力行返却
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
' @Remarks
' プロシージャ定義行の場合は
' 定義されたプロシージャ名を返す
' それ以外の行は現在のプロシージャ名を返す
'
Private Function GetProcedureName( _
            ByVal LineText As String, _
            ByVal CurrentName As String) As String

    '# Procedure定義判定
    If IsProcedureLine(LineText) Then
    
        '# 現在Procedure更新
        GetProcedureName = _
            ExtractProcedureName(LineText)
    Else

        '# 現在Procedure継続
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
' @Input
' ws(Worksheet)
'
' @Output
' なし
'
' @Summary
' 関数トレースシートへ
' 入力規則と条件付き書式を設定する
'
' @Remarks
' K列へ確認結果入力規則を設定する
'
' ○ 正常
' × 誤検出
' △ 要確認
' 除外
'
' の選択を可能とし
' 条件付き書式による色分けを行う
'
Private Sub SetupTraceValidation(ByVal ws As Worksheet)

    ' 最終行
    Dim LastRow As Long

    '# FunctionTrace最終行取得
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' 見出し
    '# 確認列作成
    ws.Range("K1").Value = "確認"
    ws.Range("L1").Value = "コメント"

    If LastRow < 2 Then Exit Sub

    ' 入力規則
    '# 確認入力規則設定
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


    '# 条件付き書式初期化
    ws.Range("A2:L" & LastRow).FormatConditions.Delete

    ' ○ 正常（薄緑）
    '# 正常判定色設定
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""○ 正常""", _
        RGB(198, 239, 206)

    ' × 誤検出（薄赤）
    '# 誤検出色設定
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""× 誤検出""", _
        RGB(255, 199, 206)

    ' △ 要確認（薄黄）
    '# 要確認色設定
    AddConditionColor _
        ws.Range("A2:L" & LastRow), _
        "=$K2=""△ 要確認""", _
        RGB(255, 235, 156)

    ' 除外（薄灰）
    '# 除外色設定
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
' 指定範囲へ条件付き書式を追加し
' 条件成立時の背景色を設定する
'
' @Remarks
' 条件付き書式を追加し
' 条件成立時の背景色を設定する
'
' 既存条件付き書式は
' 削除せず追加する
'
Private Sub AddConditionColor( _
            ByVal rng As Range, _
            ByVal Formula As String, _
            ByVal FillColor As Long)

    ' 条件付き書式
    Dim fc As FormatCondition

    '# 条件付き書式追加
    Set fc = rng.FormatConditions.Add( _
                Type:=xlExpression, _
                Formula1:=Formula)

    '# 背景色設定
    fc.Interior.Color = FillColor

End Sub

