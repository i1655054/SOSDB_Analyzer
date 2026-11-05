Attribute VB_Name = "modAnalyzerCommon"
Option Explicit

'==================================================
' 共通解析ユーティリティ
'
' 解析共通処理
' コメント解析処理
' ProcList参照処理
' を提供する
'==================================================

' API宣言
#If VBA7 Then
    ' ユーザー32.dll（クリップボード操作）
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    ' カーネル32.dll（メモリ操作）
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    ' VBA6 用の宣言（必要なら追加）
#End If

' GlobalAllocフラグ
' GMEM_MOVEABLE + GMEM_ZEROINIT
Const GHND = &H42

' Unicodeテキスト形式
Const CF_UNICODETEXT = 13

' @JPName
' 解析コンテキスト生成
'
' @Category
' Analyze
'
' @Input
' なし
'
' @Output
' clsAnalyzeContext
'
' @Summary
' 現在選択中ツールの
' TargetBook、ManagementBook、
' ToolInfoを保持した
' 解析コンテキストを生成する
'
' @Remarks
' gContextへ設定する前の
' インスタンス生成で利用する
'
Public Function CreateAnalyzeContext() _
            As clsAnalyzeContext

    ' 解析コンテキスト
    Dim Context As clsAnalyzeContext

    '# 解析コンテキスト生成
    Set Context = New clsAnalyzeContext

    '# 対象ブック取得
    Set Context.TargetBook = GetCurrentTargetBook()

    '# 管理ブック取得
    Set Context.ManagementBook = GetCurrentManagementWorkbook()

    '# 選択ツール情報取得
    Set Context.Tool = GetCurrentToolInfo()

    '# 解析コンテキスト返却
    Set CreateAnalyzeContext = Context

End Function

' @JPName
' 全解析実行
'
' @Category
' Analyze
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ProcList、関数トレース、
' 関数依存関係の解析を実行する
'
' @Remarks
' gContextの妥当性確認後に
' 各解析処理を順次実行する
'
Public Sub AnalyzeAll()

    ' 未使用関数件数
    'Dim UnusedCount As Long

    ' 呼出元なし関数件数
    Dim NoCallerCount As Long

    '# 解析コンテキスト確認
    If gContext Is Nothing Then
    
        MsgBox "解析コンテキスト未生成"

        Exit Sub

    End If

    '# 解析コンテキスト妥当性確認
    If Not gContext.IsValid Then

        MsgBox "解析コンテキスト生成失敗"

        Exit Sub

    End If

    '# 対象ブック確認
    If Not CheckTargetBook( _
            gContext.TargetBook, _
            gContext.Tool.WorkbookName) Then
        Exit Sub
    End If

    '# 管理ブック確認
    If Not CheckManagementWorkbook( _
            gContext.ManagementBook, _
            gContext.Tool.ManagementFile) Then
        Exit Sub
    End If
    
    On Error GoTo ExitProc
    
    Application.StatusBar = "Step 1/4 ProcedureList生成中"
    
    '# ProcList生成
    ProcedureList
    
    Application.StatusBar = "Step 2/4 NoCallerProcList生成中"
    
    '# 呼出元なしProcedure抽出
    NoCallerCount = _
        ExportNoCallerProcedureList()
    
    Application.StatusBar = "Step 3/4 FunctionTrace生成中"
    
    '# FunctionTrace生成
    FunctionTrace

    Application.StatusBar = "Step 4/4 FunctionDependency生成中"
    
    '# FunctionDependency生成
    FunctionDependency

    Application.StatusBar = False
    
    '# 解析結果通知
    MsgBox _
        "解析完了" & vbCrLf & _
        "呼出元なし : " & _
        NoCallerCount & "件", _
        vbInformation

    Exit Sub
    
ExitProc:

    Application.StatusBar = False
    
    '# エラー内容通知
    MsgBox _
        "解析中にエラーが発生しました。" & vbCrLf & _
        Err.Description, _
        vbInformation

End Sub

'==================================================
' 共通ユーティリティ
'==================================================

' @JPName
' 対象ブック確認
'
' @Category
' Common
'
' @Input
' TargetBook(Workbook)
' BookName(String)
'
' @Output
' True:有効
' False:無効
'
' @Summary
' 対象ブックが開かれているか確認する
'
' @Remarks
' 未オープン時はメッセージを表示する
'
Public Function CheckTargetBook( _
                    ByVal TargetBook As Workbook, _
                    Optional ByVal BookName As String = "") _
                    As Boolean

    '# 対象ブック存在確認
    If TargetBook Is Nothing Then

        If Trim$(BookName) = "" Then

            '# 対象ブック未オープン通知
            MsgBox _
                "対象ブックが開かれていません。" & vbCrLf & _
                "対象ブック名が指定されていません。", _
                vbExclamation

        Else

            MsgBox _
                "対象ブックが開かれていません。" & vbCrLf & _
                "対象: " & BookName, _
                vbExclamation

        End If

        Exit Function

    End If

    '# 対象ブック有効
    CheckTargetBook = True

End Function

'===================================================================
'@JPName     クリップボードコピー
'@Category   UI支援
'@Input      Text(String)
'@Output     なし
'@Summary    指定文字列をクリップボードへコピーする
'@Remarks    Unicode形式でコピーする
'===================================================================
Public Sub CopyTextToClipboard(Text As String)

    ' グローバルメモリハンドル
    Dim hGlobalMemory As LongPtr
    
    ' メモリポインタ
    Dim lpGlobalMemory As LongPtr
    
    ' コピーサイズ(Byte)
    Dim Size As LongPtr

    '# コピーサイズ算出
    Size = (Len(Text) + 1) * 2 ' Unicodeは2バイト

    '# クリップボードへコピー
    If OpenClipboard(0&) Then
        EmptyClipboard
        hGlobalMemory = GlobalAlloc(GHND, Size)
        lpGlobalMemory = GlobalLock(hGlobalMemory)
        CopyMemory ByVal lpGlobalMemory, ByVal StrPtr(Text), Size
        GlobalUnlock hGlobalMemory
        SetClipboardData CF_UNICODETEXT, hGlobalMemory
        CloseClipboard
    End If
End Sub

'==================================================
' コメント解析
'==================================================

' @JPName
' タグ値取得
'
' @Category
' Analyze
'
' @Input
' CodeMod(CodeModule)
' StartLine(Long)
' TagName(String)
'
' @Output
' TagValue(String)
'
' @Summary
' プロシージャ定義直前のコメントから
' 指定タグの値を取得する
'
' @Remarks
' 旧形式コメントおよび
' 複数行コメント形式の両方に対応する
'
Public Function GetTagValue( _
        CodeMod As Object, _
        ByVal StartLine As Long, _
        ByVal TagName As String) As String

    ' コメント行走査用
    Dim i As Long
    
    ' タグ内容走査用
    Dim j As Long
    
    ' 現在行文字列
    Dim txt As String
    
    ' 取得結果
    Dim Result As String

    '# コメントブロック走査
    For i = StartLine - 1 To 1 Step -1

        '# コメント行抽出
        txt = Trim$(CodeMod.Lines(i, 1))

        ' コメントブロック終了
        If Left$(txt, 1) <> "'" Then Exit For

        '# コメント記号除去
        txt = Mid$(txt, 2)
        txt = Trim$(txt)

        '==========================
        ' 旧形式
        ' '@JPName 〇〇
        '==========================
        '# 旧形式タグ検索
        If InStr(1, txt, TagName, vbTextCompare) = 1 Then

            '# 単一行タグ取得
            Result = Trim$(Replace( _
                        txt, _
                        TagName, _
                        "", _
                        , , _
                        vbTextCompare))

            If Result <> "" Then

                GetTagValue = Result
                Exit Function

            End If

            '==========================
            ' 新形式
            ' @JPName
            ' ○○○
            '==========================
            '# 新形式複数行タグ取得
            For j = i + 1 To StartLine - 1

                txt = Trim$(CodeMod.Lines(j, 1))

                If Left$(txt, 1) <> "'" Then Exit For

                txt = Mid$(txt, 2)

                txt = Trim$(txt)

                ' 空行
                If txt = "" Then GoTo ContinueLoop

                ' 次タグ
                If Left$(txt, 1) = "@" Then Exit For

                If Result <> "" Then

                    Result = _
                        Result & vbCrLf & txt

                Else

                    Result = txt

                End If
ContinueLoop:
            Next j

            '# タグ値返却
            GetTagValue = Result
            
            'Debug.Print "Result=[" & Result & "]"

            Exit Function

        End If

    Next i

End Function

' @JPName
' タグ値取得
'
' @Category
' Analyze
'
' @Input
' CodeText(String)
' ProcName(String)
' TagName(String)
'
' @Output
' TagValue(String)
'
' @Summary
' ソースコード文字列から
' 指定タグの値を取得する
'
' @Remarks
' 指定Procedure宣言部直前の
' コメントタグを検索する
'
Public Function GetTagValueFromCode( _
        ByVal CodeText As String, _
        ByVal ProcName As String, _
        ByVal TagName As String) As String

    ' ソースコード行配列
    Dim Lines() As String

    ' 行走査用
    Dim i As Long
    
    ' コメント走査用
    Dim j As Long

    '# ソースコード行分割
    Lines = Split(CodeText, vbCrLf)

    '# Procedure定義検索
    For i = LBound(Lines) To UBound(Lines)

        If IsProcedureLine(Trim$(Lines(i))) Then

            '# 対象Procedure一致確認
            If StrComp( _
                    ExtractProcedureName( _
                        Trim$(Lines(i))), _
                    ProcName, _
                    vbTextCompare) = 0 Then

                '# コメントタグ検索
                For j = i - 1 To 0 Step -1

                    If Left$(Trim$(Lines(j)), 1) <> "'" Then
                        Exit For
                    End If

                    If StrComp( _
                            Trim$(Lines(j)), _
                            "'" & TagName, _
                            vbTextCompare) = 0 Then

                        If j + 1 <= UBound(Lines) Then

                            '# タグ値取得
                            GetTagValueFromCode = _
                                Trim$(Replace( _
                                    Trim$(Lines(j + 1)), _
                                    "'", ""))

                        End If

                        Exit Function

                    End If

                Next j

            End If
        
        End If

    Next i

End Function


'==================================================
' ProcList参照
'==================================================

' @JPName
' 関数情報取得
'
' @Category
' Analyze
'
' @Input
' ProcName(String)
'
' @Output
' clsProcInfo
'
' @Summary
' ProcListシートから
' 指定プロシージャの情報を取得する
'
' @Remarks
' 情報が存在しない場合はNothingを返す
'
Public Function GetProcInfo( _
                ByVal ProcName As String) _
                As clsProcInfo
    
    ' ProcListシート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' Procedure情報
    Dim Proc As clsProcInfo

    If gContext Is Nothing Then Exit Function
    
    '# ProcListシート取得
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PROC_LIST)

    LastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    '# ProcList走査
    For r = 2 To LastRow

        '# Procedure一致確認
        If StrComp( _
                Trim$(ws.Cells(r, COL_PROC).Value), _
                ProcName, _
                vbTextCompare) = 0 Then

            '# Procedure情報生成
            Set Proc = New clsProcInfo
            
            Proc.ModuleName = ws.Cells(r, COL_MODULE).Value

            Proc.ProcName = ProcName

            Proc.JPName = ws.Cells(r, COL_JPNAME).Value
            Proc.Category = ws.Cells(r, COL_CATEGORY).Value

            Proc.InputText = ws.Cells(r, COL_INPUT).Value
            Proc.OutputText = ws.Cells(r, COL_OUTPUT).Value

            Proc.Summary = ws.Cells(r, COL_SUMMARY).Value
            Proc.Remarks = ws.Cells(r, COL_REMARKS).Value

            '# Procedure情報返却
            Set GetProcInfo = Proc

            Exit Function

        End If

    Next r

End Function

' @JPName
' 関数名抽出
'
' @Category
' Analyze
'
' @Input
' DisplayText(String)
'
' @Output
' ProcName(String)
'
' @Summary
' 表示文字列から
' プロシージャ名部分を取得する
'
' @Remarks
' 「ProcName : JPName」形式の場合は
' 区切り文字「 : 」より前の
' プロシージャ名を返す
' 区切り文字が存在しない場合は
' 入力文字列をそのまま返す
'
Public Function GetProcNameFromDisplay( _
                    ByVal DisplayText As String) _
                    As String

    '# 表示名からProcedure名抽出
    If InStr(DisplayText, " : ") > 0 Then

        GetProcNameFromDisplay = _
            Trim$(Split(DisplayText, " : ")(0))

    Else

        '# 表示文字列をそのまま使用
        GetProcNameFromDisplay = _
            Trim$(DisplayText)

    End If

End Function

'==================================================
' 関数解析共通
'==================================================

' @JPName
' 関数呼出判定
'
' @Category
' Analyze
'
' @Input
' LineText(String)
' FuncName(String)
'
' @Output
' True:呼出あり
' False:呼出なし
'
' @Summary
' 指定行に対象プロシージャの
' 呼出構文が存在するか判定する
'
' @Remarks
' Call文
' Function呼出
' Sub単独呼出
' に対応する
'
Public Function IsFunctionCall( _
        ByVal LineText As String, _
        ByVal FuncName As String) As Boolean

    ' 正規表現オブジェクト
    Dim RegEx As Object

    '# Procedure定義行除外
    If IsProcedureLine(LineText) Then
        Exit Function
    End If

    LineText = Trim$(LineText)

    '# コメント行除外
    If Left$(LineText, 1) = "'" Then Exit Function

    '# API説明行除外
    If Left$(LineText, 1) = "■" Then Exit Function

    Set RegEx = CreateObject("VBScript.RegExp")

    RegEx.IgnoreCase = True

    '------------------------------------
    ' 関数戻り値代入除外
    ' GetCurrentToolID = xxx
    '------------------------------------
    '# 関数戻り値代入除外
    RegEx.Pattern = _
        "^\s*" & _
        FuncName & _
        "\s*="

    If RegEx.Test(LineText) Then

        Exit Function

    End If
    
    '# Call構文判定
    ' Call Function
    RegEx.Pattern = _
        "(^|\s)Call\s+" & _
        FuncName & _
        "(\s*\(|\s|$)"

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    '# Function呼出判定
    ' Function(...)
    RegEx.Pattern = _
        "(^|[\s=\(,])" & _
        FuncName & _
        "\s*\("
    If RegEx.Test(LineText) Then
    
        IsFunctionCall = True
        Exit Function
    
    End If
    
    '------------------------------------
    ' 行継続文字対応
    ' CreatePathChart_V3 _
    '------------------------------------
    '# 行継続呼出判定
    If Right$(Trim$(LineText), 1) = "_" Then
    
        Dim WorkText As String
        
        WorkText = Trim$(LineText)

        If Left$(Trim$(LineText), Len(FuncName)) = FuncName Then

            If Len(WorkText) = Len(FuncName) + 1 Then
                IsFunctionCall = True
                Exit Function

            End If
        
            Select Case Mid$( _
                            WorkText, _
                            Len(FuncName) + 1, _
                            1)

                Case " ", "(", vbTab

                    IsFunctionCall = True
                    Exit Function

            End Select
            
        End If

    End If
    
    '# Sub呼出判定
    ' Sub Arg1, Arg2
    RegEx.Pattern = _
        "^\s*" & _
        FuncName & _
        "\s+.+$"
    
    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If
    
    '# オブジェクトメソッド呼出判定
    ' Object.Method Arg1
    RegEx.Pattern = _
        "\." & FuncName & _
        "(\s+|\s*\(|$)"

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    '# モジュール修飾呼出判定
    ' Module.Procedure Arg1, Arg2
    RegEx.Pattern = _
        "(^|\s)(\w+\.)+" & _
        FuncName & _
        "\s+.+$"

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    '# 単独呼出判定
    ' Call Module.Procedure(...)
    RegEx.Pattern = _
        "(^|\s)Call\s+(\w+\.)+" & _
        FuncName & _
        "\s*\("

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    ' Module.Procedure(...)
    RegEx.Pattern = _
        "(^|\s)(\w+\.)+" & _
        FuncName & _
        "\s*\("

    If RegEx.Test(LineText) Then

        IsFunctionCall = True
        Exit Function

    End If

    ' Function 単独呼出
    RegEx.Pattern = _
        "^\s*" & FuncName & "\s*$"

    
    IsFunctionCall = RegEx.Test(LineText)

End Function

' @JPName
' Procedure名完全一致判定
'
' @Category
' Analyze
'
' @Input
' Text(String)
' ProcName(String)
'
' @Output
' Boolean
'
' @Summary
' Procedure名が完全一致するか判定する
'
' @Remarks
' Procedure名の後続文字が
' 英数字またはアンダースコアの場合は
' 不一致とする
'
'------------------------------------------------
' 現在未使用
'
' Procedure名完全一致判定用
'
' IsFunctionCall強化のために追加したが
' 現在は未適用
'------------------------------------------------
Private Function IsExactProcedureName( _
                ByVal Text As String, _
                ByVal ProcName As String) _
                As Boolean

    '# Procedure名先頭一致確認
    If Left$(Text, Len(ProcName)) <> ProcName Then
        Exit Function
    End If

    '# 完全一致判定
    If Len(Text) = Len(ProcName) Then

        IsExactProcedureName = True
        Exit Function

    End If

    '# 後続文字判定
    Select Case Mid$( _
                    Text, _
                    Len(ProcName) + 1, _
                    1)

        Case " ", "(", vbTab

            IsExactProcedureName = True

    End Select

End Function

