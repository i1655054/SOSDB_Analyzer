Attribute VB_Name = "modHCPTargetManager"
Option Explicit

'===================================================================
'@JPName      HCP_TARGETS管理
'===================================================================

Private Const COL_PROCLIST_MODULE As Long = 1
Private Const COL_PROCLIST_TYPE As Long = 2
Private Const COL_PROCLIST_NAME As Long = 3

'--------------------------------------------------
' 作成・更新
'--------------------------------------------------

'
' @JPName
' HCP_TARGETS作成
'
' @Summary
' ProcListから
' HCP_TARGETSを再作成する
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

    '# ProcList取得
    Set wsProc = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

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
  
    '# INDEX更新
    RebuildAllIndexes
 
    MsgBox _
        "HCP_TARGETSを作成しました。" & vbCrLf & _
        "対象件数 : " & CreateCount, _
        vbInformation

End Sub

'
' @JPName
' HCP_TARGETS更新
'
' @Summary
' ProcListとの差分を
' HCP_TARGETSへ追加する
'
Public Sub UpdateHCPTargets( _
                Optional ByVal Silent As Boolean = False)
    
    Dim wsProc As Worksheet
    Dim wsTarget As Worksheet

    Dim Dict As Object

    Dim LastProcRow As Long
    Dim LastTargetRow As Long

    Dim RowOut As Long
    Dim r As Long

    Dim ProcModule As String
    Dim ProcName As String

    Dim Key As String

    Dim AddCount As Long

    Trace "HCP", _
        "UpdateHCPTargets"

    Set Dict = _
        CreateObject("Scripting.Dictionary")

    '# ProcList取得
    Set wsProc = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    '# HCP_TARGETS取得
    Set wsTarget = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    '# HCP_TARGETS最終行取得
    LastTargetRow = _
        wsTarget.Cells( _
            wsTarget.Rows.Count, _
            2).End(xlUp).row

    '# 既存登録情報取得
    For r = 2 To LastTargetRow

        Key = _
            Trim$(wsTarget.Cells(r, 2).Value) & _
            "|" & _
            Trim$(wsTarget.Cells(r, 3).Value)

        Dict(Key) = True

    Next r

    '# ProcList最終行取得
    LastProcRow = _
        wsProc.Cells( _
            wsProc.Rows.Count, _
            COL_PROCLIST_MODULE).End(xlUp).row

    RowOut = LastTargetRow + 1

    '# 差分追加
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
                ProcModule & "|" & ProcName

            If Not Dict.Exists(Key) Then

                Trace "HCP", _
                    "Add " & _
                    ProcModule & "." & ProcName
                
                wsTarget.Cells(RowOut, 2).Value = _
                    ProcModule

                wsTarget.Cells(RowOut, 3).Value = _
                    ProcName

                RowOut = RowOut + 1

                AddCount = _
                    AddCount + 1

            End If

        End If

    Next r

    '# 体裁設定
    wsTarget.Columns("A:E").AutoFit

    Trace "HCP", _
        "UpdateHCPTargets AddCount=" & _
        AddCount

    If Not Silent Then
    
        MsgBox _
            "HCP_TARGETS更新完了" & vbCrLf & _
            "追加件数 : " & AddCount, _
            vbInformation
    
    End If
    
End Sub

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

'
' @JPName
' HCP一括生成
'
' @Summary
' HCP_TARGETSをもとに
' HCPを一括生成する
'
' @Remarks
' Exec=Y かつ
' Status<>"Done" のものを
' 順次処理する
'
Public Sub GenerateHCPBatch()

    Dim ws As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim ModuleName As String
    Dim ProcName As String

    Dim CreateCount As Long
    Dim ErrorCount As Long

    Dim TargetCount As Long
    Dim CurrentCount As Long

    Trace "HCP", _
        "GenerateHCPBatch Start"

    '# HCP_TARGETS更新
    UpdateHCPTargets True

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

    '# 初期表示
    frmHCPGenerator.lblProgress.Caption = _
        "0 / " & TargetCount

    frmHCPGenerator.lblCurrentProc.Caption = _
        "開始準備中"

    DoEvents

    '-----------------------------
    '# HCP生成
    '-----------------------------
    For r = 2 To LastRow

        If UCase$(Trim$( _
            ws.Cells(r, 1).Value)) = "Y" Then

            If Trim$( _
                ws.Cells(r, 4).Value) <> "Done" Then

                ModuleName = _
                    Trim$(ws.Cells(r, 2).Value)

                ProcName = _
                    Trim$(ws.Cells(r, 3).Value)

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

                '# HCP生成
                CreateHCP _
                    ModuleName, _
                    ProcName

                ws.Cells(r, 4).Value = _
                    "Done"

                ws.Cells(r, 5).Value = _
                    Format$( _
                        Now, _
                        "yyyy/mm/dd hh:nn:ss")

                CreateCount = _
                    CreateCount + 1

                CurrentCount = _
                    CurrentCount + 1

                frmHCPGenerator.lblProgress.Caption = _
                    CurrentCount & _
                    " / " & _
                    TargetCount

                DoEvents

                GoTo ContinueLoop

ProcError:

                ws.Cells(r, 4).Value = _
                    "Error"

                ws.Cells(r, 5).Value = _
                    Format$( _
                        Now, _
                        "yyyy/mm/dd hh:nn:ss")

                ErrorCount = _
                    ErrorCount + 1

                frmHCPGenerator.lblCurrentProc.Caption = _
                    "[ERROR] " & _
                    ModuleName & "." & _
                    ProcName

                Trace "ERROR", _
                    ModuleName & "." & _
                    ProcName & _
                    " : " & _
                    Err.Description

                Err.Clear

            End If

        End If

ContinueLoop:

        On Error GoTo 0

    Next r

    '# 完了表示
    frmHCPGenerator.lblProgress.Caption = _
        CurrentCount & _
        " / " & _
        TargetCount

    frmHCPGenerator.lblCurrentProc.Caption = _
        "完了"

    '# INDEX再生成
        Trace "HCP", _
            "RebuildAllIndexes"

    RebuildAllIndexes
    
    Trace "HCP", _
        "GenerateHCPBatch End Success=" & _
        CreateCount & _
        " Error=" & _
        ErrorCount

    MsgBox _
        "HCP一括生成完了" & vbCrLf & _
        "成功 : " & CreateCount & vbCrLf & _
        "エラー : " & ErrorCount, _
        vbInformation

End Sub

'
' @JPName
' HCP一括生成確認メッセージ作成
'
' @Summary
' HCP_TARGETSから
' 一括生成対象一覧を取得し
' 確認メッセージを作成する
'
'
Public Function BuildHCPBatchMessage() _
                As String

    Dim ws As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim TargetCount As Long

    Dim Message As String

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            2).End(xlUp).row

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

