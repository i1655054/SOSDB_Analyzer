Attribute VB_Name = "modHCPTargetManager"
Option Explicit

'===================================================================
' @JPName
' HCP_TARGETS管理
'
' @Category
' HCP
'
' @Summary
' HCP_TARGETSシートの作成、
' 更新および一括生成機能を提供する
'
' @Remarks
' ProcListから対象Procedureを抽出し
' HCP生成対象管理を行う
'
' HCP_TARGETS作成
' HCP_TARGETS更新
' HCP一括生成
'
' を管理する
'===================================================================

Private Const COL_PROCLIST_MODULE As Long = 1
Private Const COL_PROCLIST_TYPE As Long = 2
Private Const COL_PROCLIST_NAME As Long = 3

'--------------------------------------------------
' 作成・更新
'--------------------------------------------------

' @JPName
' HCP_TARGETS作成
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' HCP_TARGETSシート
'
' @Summary
' ProcListから
' HCP_TARGETSを再作成する
'
' @Remarks
' HCP_TARGETSを初期化後
' ProcList全件を登録する
'
' 処理完了後に
' INDEXを再生成する
'
Public Sub CreateHCPTargets()

    Dim wsProc As Worksheet
    Dim wsTarget As Worksheet

    Dim LastRow As Long
    Dim RowOut As Long
    Dim r As Long
    
    Dim CreateCount As Long

    Dim ProcModule As String
    Dim ProcName As String

    Trace "HCP", _
        "CreateHCPTargets"

    '-----------------------------
    '# シート取得
    '-----------------------------
    '# ProcList取得
    Set wsProc = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    '-----------------------------
    '# 初期化
    '-----------------------------
    '# HCP_TARGETS取得
    On Error Resume Next

    Set wsTarget = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    On Error GoTo 0

    '# HCP_TARGETS未存在なら作成
    If wsTarget Is Nothing Then

        Set wsTarget = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                    gContext.ManagementBook.Worksheets.Count))

        wsTarget.Name = _
            SHEET_HCP_TARGETS

    End If

    '# HCP_TARGETS初期化
    Call InitializeHCPTargetSheet(wsTarget)

    '# ProcList最終行取得
    LastRow = _
        wsProc.Cells( _
            wsProc.Rows.Count, _
            COL_PROCLIST_MODULE).End(xlUp).row

    RowOut = 2

    '-----------------------------
    '# ProcList走査
    '-----------------------------
    '# ProcList → HCP_TARGETS
    For r = 2 To LastRow

        '# Module名取得
        ProcModule = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_MODULE).Value)

        '# Procedure名取得
        ProcName = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_NAME).Value)

        '# Procedure名未設定は除外
        If ProcModule <> "" _
        And ProcName <> "" Then

            wsTarget.Cells(RowOut, 2).Value = _
                ProcModule

            wsTarget.Cells(RowOut, 3).Value = _
                ProcName

            RowOut = RowOut + 1
            
            CreateCount = CreateCount + 1

        End If

    Next r

    '# 体裁設定
    wsTarget.Columns("A:E").AutoFit

    Trace "HCP", _
        "CreateHCPTargets Count=" & _
        CreateCount
  
    '-----------------------------
    '# INDEX再生成
    '-----------------------------
    '# INDEX更新
    RebuildAllIndexes
 
    '-----------------------------
    '# 完了通知
    '-----------------------------
    MsgBox _
        "HCP_TARGETSを作成しました。" & vbCrLf & _
        "対象件数 : " & CreateCount, _
        vbInformation

End Sub

'===================================================================
' @JPName
' HCP_TARGETS更新
'
' @Category
' HCP
'
' @Input
' Silent(Boolean)
'
' @Output
' HCP_TARGETS
'
' @Summary
' ProcListとHCP_TARGETSを同期する
'
' @Remarks
' ProcListに存在しないProcedureは削除し
' ProcListに存在するProcedureは追加する
'
' HCP_TARGETSを
' ProcListと同一状態へ同期する
'===================================================================
Public Sub UpdateHCPTargets( _
                Optional ByVal Silent As Boolean = False)

    Dim wsProc As Worksheet
    Dim wsTarget As Worksheet

    Dim Dict As Object
    Dim ProcDict As Object

    Dim LastProcRow As Long
    Dim LastTargetRow As Long

    Dim RowOut As Long
    Dim r As Long

    Dim ProcModule As String
    Dim ProcName As String

    Dim Key As String

    Dim AddCount As Long
    Dim DeleteCount As Long

    Trace "HCP", _
        "UpdateHCPTargets"

    '-----------------------------
    '# Dictionary初期化
    '-----------------------------
    Set Dict = _
        CreateObject("Scripting.Dictionary")

    Set ProcDict = _
        CreateObject("Scripting.Dictionary")

    '-----------------------------
    '# シート取得
    '-----------------------------
    '# ProcList取得
    Set wsProc = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    '# HCP_TARGETS取得
    Set wsTarget = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    '-----------------------------
    '# ProcList最終行取得
    '-----------------------------
    LastProcRow = _
        wsProc.Cells( _
            wsProc.Rows.Count, _
            COL_PROCLIST_MODULE).End(xlUp).row

    '-----------------------------
    '# ProcList読込
    '-----------------------------
    For r = 2 To LastProcRow

        ProcModule = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_MODULE).Value)

        ProcName = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_NAME).Value)

        If ProcModule <> "" _
        And ProcName <> "" Then

            Key = _
                ProcModule & "|" & _
                ProcName

            ProcDict(Key) = True

        End If

    Next r

    '-----------------------------
    '# HCP_TARGETS最終行取得
    '-----------------------------
    LastTargetRow = _
        wsTarget.Cells( _
            wsTarget.Rows.Count, _
            2).End(xlUp).row

    '-----------------------------
    '# 不要行削除
    '-----------------------------
    For r = LastTargetRow To 2 Step -1

        Key = _
            Trim$(wsTarget.Cells(r, 2).Value) & _
            "|" & _
            Trim$(wsTarget.Cells(r, 3).Value)

        If Not ProcDict.Exists(Key) Then

            Trace "HCP", _
                "Delete " & _
                Key

            wsTarget.Rows(r).Delete

            DeleteCount = _
                DeleteCount + 1

        End If

    Next r

    '-----------------------------
    '# HCP_TARGETS再読込
    '-----------------------------
    LastTargetRow = _
        wsTarget.Cells( _
            wsTarget.Rows.Count, _
            2).End(xlUp).row

    For r = 2 To LastTargetRow

        Key = _
            Trim$(wsTarget.Cells(r, 2).Value) & _
            "|" & _
            Trim$(wsTarget.Cells(r, 3).Value)

        Dict(Key) = True

    Next r

    RowOut = _
        LastTargetRow + 1

    '-----------------------------
    '# ProcListとの差分追加
    '-----------------------------
    For r = 2 To LastProcRow

        ProcModule = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_MODULE).Value)

        ProcName = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_NAME).Value)

        If ProcModule <> "" _
        And ProcName <> "" Then

            Key = _
                ProcModule & "|" & _
                ProcName

            If Not Dict.Exists(Key) Then

                Trace "HCP", _
                    "Add " & _
                    ProcModule & "." & _
                    ProcName

                wsTarget.Cells(RowOut, 2).Value = _
                    ProcModule

                wsTarget.Cells(RowOut, 3).Value = _
                    ProcName

                RowOut = _
                    RowOut + 1

                AddCount = _
                    AddCount + 1

            End If

        End If

    Next r

    '-----------------------------
    '# 表示更新
    '-----------------------------
    wsTarget.Columns("A:G").AutoFit

    Trace "HCP", _
        "UpdateHCPTargets Add=" & _
        AddCount & _
        " Delete=" & _
        DeleteCount

    '-----------------------------
    '# 完了通知
    '-----------------------------
    If Not Silent Then

        MsgBox _
            "HCP_TARGETS更新完了" & vbCrLf & _
            "追加件数 : " & _
            AddCount & vbCrLf & _
            "削除件数 : " & _
            DeleteCount & vbCrLf & _
            "現在件数 : " & _
            RowOut - 2, _
            vbInformation

    End If

End Sub


'===================================================================
' @JPName
' HCP_TARGETS初期化
'
' @Category
' HCP
'
' @Input
' ws(Worksheet)
'
' @Output
' HCP_TARGETSシート
'
' @Summary
' HCP_TARGETSシートを
' 初期状態へ設定する
'
' @Remarks
' ヘッダ生成、
' 入力規則設定、
' 条件付き書式設定、
' オートフィルタ設定を行う
'===================================================================
Private Sub InitializeHCPTargetSheet( _
                ByVal ws As Worksheet)

    '# HCP_TARGETS初期化
    ws.Cells.Clear

    '# ヘッダ出力
    ws.Range("A1:E1").Value = Array( _
        "Exec", _
        "ModuleName", _
        "ProcName", _
        "Status", _
        "LastUpdate")

    '# Exec入力規則
    With ws.Range("A2:A10000").Validation

        .Delete

        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Formula1:=",Y"

        .IgnoreBlank = True

        .InCellDropdown = True

    End With

    '# Exec=Yを緑表示
    With ws.Range("A2:A10000")

        .FormatConditions.Delete

        .FormatConditions.Add _
            Type:=xlCellValue, _
            Operator:=xlEqual, _
            Formula1:="=""Y"""

        .FormatConditions(1).Interior.Color = _
            RGB(198, 239, 206)

    End With

    '# 体裁設定
    ws.Rows(1).Font.Bold = True

    ws.Range("A1:E1").AutoFilter

End Sub

'--------------------------------------------------
' 一括生成
'--------------------------------------------------

'===================================================================
' @JPName
' HCP一括生成
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' HCPファイル群
'
' @Summary
' HCP_TARGETSシートの生成対象を走査し
' HCPチャートを一括生成する
'
' @Remarks
' HCP_TARGETSシートの
' Exec=Yかつ
' Status<>"Done"
'
' のProcedureを対象とする
'
' 生成結果は
'
' Status
' 実行日時
' ErrorNo
' ErrorMessage
'
' としてHCP_TARGETSへ記録する
'
' エラー発生時も
' 処理を継続し
' 全対象の生成を実施する
'
' HCP_TARGETS更新は
' 事前に実施すること
'===================================================================
Public Sub GenerateHCPBatch()

    ' HCP_TARGETSシート
    Dim ws As Worksheet

    ' 最終行
    Dim LastRow As Long

    ' 行番号
    Dim r As Long

    ' モジュール名
    Dim ModuleName As String

    ' Procedure名
    Dim ProcName As String

    ' 生成成功件数
    Dim CreateCount As Long

    ' エラー件数
    Dim ErrorCount As Long

    ' 対象件数
    Dim TargetCount As Long

    ' 処理件数
    Dim CurrentCount As Long

    Trace "HCP", _
        "GenerateHCPBatch Start"

    '-----------------------------
    '# HCP_TARGETS取得
    '-----------------------------
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            2).End(xlUp).row

    '-----------------------------
    '# 対象件数取得
    '-----------------------------
    For r = 2 To LastRow

        If UCase$(Trim$( _
            ws.Cells(r, 1).Value)) = "Y" Then

            If Trim$( _
                ws.Cells(r, 4).Value) <> "Done" Then

                TargetCount = _
                    TargetCount + 1

            End If

        End If

    Next r

    '-----------------------------
    '# 進捗表示初期化
    '-----------------------------
    frmHCPGenerator.lblProgress.Caption = _
        "0 / " & TargetCount

    frmHCPGenerator.lblCurrentProc.Caption = _
        "開始準備中"

    DoEvents

    '-----------------------------
    '# HCP生成開始
    '-----------------------------
    For r = 2 To LastRow

        '-----------------------------
        '# 生成対象判定
        '-----------------------------
        If UCase$(Trim$( _
            ws.Cells(r, 1).Value)) = "Y" Then

            If Trim$( _
                ws.Cells(r, 4).Value) <> "Done" Then

                '-----------------------------
                '# 対象取得
                '-----------------------------
                ModuleName = _
                    Trim$(ws.Cells(r, 2).Value)

                ProcName = _
                    Trim$(ws.Cells(r, 3).Value)

                '-----------------------------
                '# 実行中表示
                '-----------------------------
                frmHCPGenerator.lblCurrentProc.Caption = _
                    "実行中 : " & _
                    ModuleName & "." & _
                    ProcName

                DoEvents

                Trace "HCP", _
                    "Generate " & _
                    ModuleName & "." & _
                    ProcName

                On Error GoTo ProcError

                '-----------------------------
                '# HCP生成
                '-----------------------------
                CreateHCP _
                    ModuleName, _
                    ProcName

                '-----------------------------
                '# 正常終了記録
                '-----------------------------
                ws.Cells(r, 4).Value = _
                    "Done"

                ws.Cells(r, 5).Value = _
                    Format$( _
                        Now, _
                        "yyyy/mm/dd hh:nn:ss")

                ' ErrorNo
                ws.Cells(r, 6).ClearContents

                ' ErrorMessage
                ws.Cells(r, 7).ClearContents

                CreateCount = _
                    CreateCount + 1

                CurrentCount = _
                    CurrentCount + 1

                frmHCPGenerator.lblProgress.Caption = _
                    CurrentCount & _
                    " / " & _
                    TargetCount

                frmHCPGenerator.lblCurrentProc.Caption = _
                    "[OK] " & _
                    ModuleName & "." & _
                    ProcName

                Application.StatusBar = _
                    "HCP生成完了 : " & _
                    ModuleName & "." & _
                    ProcName & _
                    " [OK]" & _
                    "  エラー:" & _
                    ErrorCount & "件"
    

                DoEvents

                GoTo ContinueLoop

ProcError:

                '-----------------------------
                '# HCP生成失敗
                '-----------------------------
                ws.Cells(r, 4).Value = _
                    "Error"

                ws.Cells(r, 5).Value = _
                    Format$( _
                        Now, _
                        "yyyy/mm/dd hh:nn:ss")

                ' ErrorNo
                ws.Cells(r, 6).Value = _
                    Err.Number

                ' ErrorMessage
                ws.Cells(r, 7).Value = _
                    Err.Description

                ws.Cells(r, 8).Value = _
                    ModuleName & "." & ProcName

                ws.Cells(r, 9).Value = Err.Source

                ErrorCount = _
                    ErrorCount + 1

                CurrentCount = _
                    CurrentCount + 1

                frmHCPGenerator.lblProgress.Caption = _
                    CurrentCount & _
                    " / " & _
                    TargetCount

                frmHCPGenerator.lblCurrentProc.Caption = _
                    "[ERROR] " & _
                    ModuleName & "." & _
                    ProcName

                Trace "ERROR", _
                    ModuleName & "." & _
                    ProcName & _
                    " : " & _
                    Err.Description

                Application.StatusBar = _
                    "HCP生成失敗 : " & _
                    ModuleName & "." & _
                    ProcName & _
                    " [NG]" & _
                    "  エラー:" & _
                    ErrorCount & "件"
    
                Err.Clear

                Resume ContinueLoop

            End If

        End If

ContinueLoop:

        On Error GoTo 0

    Next r

    '-----------------------------
    '# 完了表示
    '-----------------------------
    frmHCPGenerator.lblProgress.Caption = _
        CurrentCount & _
        " / " & _
        TargetCount

    If ErrorCount = 0 Then

        frmHCPGenerator.lblCurrentProc.Caption = _
            "完了"

    Else

        frmHCPGenerator.lblCurrentProc.Caption = _
            "エラーあり : " & _
            ErrorCount & "件"

    End If

    '-----------------------------
    '# INDEX再生成
    '-----------------------------
    Trace "HCP", _
        "RebuildAllIndexes"

    RebuildAllIndexes

    Trace "HCP", _
        "GenerateHCPBatch End Success=" & _
        CreateCount & _
        " Error=" & _
        ErrorCount

    '-----------------------------
    '# 完了通知
    '-----------------------------
    
    Application.StatusBar = _
        "HCP一括生成完了 " & _
        "成功=" & CreateCount & _
        "件  エラー=" & ErrorCount & "件"
    
    MsgBox _
        "HCP一括生成完了" & vbCrLf & _
        "成功 : " & _
        CreateCount & vbCrLf & _
        "エラー : " & _
        ErrorCount, _
        vbInformation

End Sub

' @JPName
' HCP一括生成確認メッセージ作成
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' Message(String)
'
' @Summary
' HCP一括生成対象一覧の
' 確認メッセージを生成する
'
' @Remarks
' Exec=Y かつ
' Status<>"Done"
'
' のProcedureを対象とする
'
' 表示件数は
' 先頭20件までとする
'
Public Function BuildHCPBatchMessage() _
                As String

    Dim ws As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim TargetCount As Long

    Dim Message As String

    '-----------------------------
    '# HCP_TARGETS取得
    '-----------------------------
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            2).End(xlUp).row

    '-----------------------------
    '# 対象一覧作成
    '-----------------------------
    For r = 2 To LastRow

        If UCase$(Trim$( _
            ws.Cells(r, 1).Value)) = "Y" Then

            If Trim$( _
                ws.Cells(r, 4).Value) <> "Done" Then

                TargetCount = _
                    TargetCount + 1

                If TargetCount <= 20 Then

                    Message = _
                        Message & _
                        ws.Cells(r, 2).Value & _
                        "." & _
                        ws.Cells(r, 3).Value & _
                        vbCrLf

                End If

            End If

        End If

    Next r

    '-----------------------------
    '# 確認メッセージ生成
    '-----------------------------
    BuildHCPBatchMessage = _
        "HCP一括生成を実行します。" & _
        vbCrLf & vbCrLf & _
        "対象件数 : " & _
        TargetCount & _
        vbCrLf & vbCrLf & _
        Message

    If TargetCount > 20 Then

        BuildHCPBatchMessage = _
            BuildHCPBatchMessage & _
            vbCrLf & _
            "※先頭20件のみ表示"

    End If

End Function

'===================================================================
' @JPName
' HCP_TARGETS整合性チェック
'
' @Category
' Debug
'
' @Input
' なし
'
' @Output
' 検証結果
'
' @Summary
' ProcListとHCP_TARGETSの
' 整合性を検証する
'
' @Remarks
'
' ProcListに存在しない
' Procedureが
' HCP_TARGETSへ残っていないか確認する
'
' トラブル調査用
'===================================================================
Public Sub CheckHCPTargetIntegrity()

    Dim wsProc As Worksheet
    Dim wsTarget As Worksheet

    Dim Dict As Object

    Dim r As Long
    Dim LastRow As Long

    Dim Key As String

    ' NG件数
    Dim NGCount As Long

    ' 結果一覧
    Dim Result As String

    Set Dict = _
        CreateObject("Scripting.Dictionary")

    Set wsProc = _
        Worksheets(SHEET_PROC_LIST)

    Set wsTarget = _
        Worksheets(SHEET_HCP_TARGETS)

    LastRow = _
        wsProc.Cells( _
            wsProc.Rows.Count, _
            COL_PROCLIST_MODULE).End(xlUp).row

    For r = 2 To LastRow

        Key = _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_MODULE).Value) & _
            "|" & _
            Trim$(wsProc.Cells( _
                r, _
                COL_PROCLIST_NAME).Value)

        Dict(Key) = True

    Next r

    Result = _
        "CheckHCPTargetIntegrity Result" & vbCrLf & _
        String(40, "-") & vbCrLf

    Result = Result & _
        "ProcList Count = " & _
        Dict.Count & vbCrLf

    LastRow = _
        wsTarget.Cells( _
            wsTarget.Rows.Count, _
            2).End(xlUp).row

    Result = Result & _
        "HCP_TARGETS Count = " & _
        LastRow - 1 & vbCrLf

    For r = LastRow To 2 Step -1

        Key = _
            Trim$(wsTarget.Cells(r, 2).Value) & _
            "|" & _
            Trim$(wsTarget.Cells(r, 3).Value)

        If Not Dict.Exists(Key) Then

            NGCount = _
                NGCount + 1

            Result = Result & _
                "NG : " & _
                Key & vbCrLf

        End If

    Next r

    Result = Result & vbCrLf & _
        "NG Count = " & _
        NGCount

    Debug.Print Result

    If NGCount = 0 Then

        MsgBox _
            "HCP_TARGETS整合性OK" & vbCrLf & _
            "ProcList=" & Dict.Count, _
            vbInformation

    Else

        MsgBox _
            "HCP_TARGETS不整合あり" & vbCrLf & _
            "NG=" & NGCount, _
            vbExclamation

    End If

End Sub


