VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPathChart 
   Caption         =   "frmPathChart"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmPathChart.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPathChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' @JPName
' PathChart画面初期化
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChart画面の初期表示処理を行う
'
' @Remarks
' 設定読込および開始関数候補を初期化する
'
Private Sub UserForm_Initialize()

    ' 開始関数ComboBoxのIndex
    Dim i As Long
    
    ' 保存済み開始関数
    Dim DefaultProc As String

    ' 開始関数存在フラグ
    Dim Found As Boolean
    
    '# PathChart設定読込
    LoadPathChartConfig
    
    '# 呼出条件設定反映
    chkExternal.Value = gExternal
    
    '# 再帰表示設定反映
    chkRecursive.Value = gRecursive

    '# 最大深度初期設定
    txtMaxDepth.Text = GetDefaultMaxDepth()
    
    '# 開始関数候補読込
    LoadStartProcedure cmbStartProc

    '# 既定開始関数取得
    DefaultProc = GetDefaultStartProcedure()

    '# 保存済み開始関数選択
    For i = 0 To cmbStartProc.ListCount - 1
    
        If cmbStartProc.List(i) = DefaultProc Then

            cmbStartProc.ListIndex = i
            
            Found = True

            Exit For

        End If

    Next i
    
    '# 開始関数既定設定
    If Not Found Then

        If cmbStartProc.ListCount > 0 Then

            cmbStartProc.ListIndex = 0
        
        End If
        
    End If

    If cmbStartProc.ListIndex < 0 _
    And cmbStartProc.ListCount > 0 Then

        cmbStartProc.ListIndex = 0

    End If

    '# 初期画面状態設定
    lblStatus.Caption = "状態： 待機中"
    
    lblElapsed.Caption = "解析時間： 0.00 秒"

End Sub

' @JPName
' 開始関数変更
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択された開始関数の情報を表示する
'
' @Remarks
' Procedure情報を取得し
' JPName、Category、Summaryを画面へ表示する
'
Private Sub cmbStartProc_Change()

    ' Procedure情報
    Dim Proc As clsProcInfo
    
    ' 選択されたProcedure名
    Dim ProcName As String
    
    '# 選択Procedure取得
    ProcName = Trim$(cmbStartProc.Value)
    
    If ProcName = "" Then Exit Sub

    '# Procedure情報取得
    Set Proc = GetProcInfo(ProcName)
    
    '# Procedure情報表示
    If Not Proc Is Nothing Then
      
        If Proc.JPName <> "" Then
            lblSelectProc.Caption = Proc.JPName
        Else
            lblSelectProc.Caption = ProcName
        End If
        
        lblCategory.Caption = Proc.Category
        txtSummary.Text = Proc.Summary

    '# Procedure情報クリア
    Else

        lblSelectProc.Caption = ""
        lblCategory.Caption = ""
        txtSummary.Text = ""

    End If

End Sub

' @JPName
' 画面終了制御
'
' @Category
' PathChart
'
' @Input
' Cancel(Integer)
' CloseMode(Integer)
'
' @Output
' なし
'
' @Summary
' フォーム終了時の制御を行う
'
' @Remarks
' ×ボタン押下時は
' 関数管理画面を再表示する
'
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    '# ×ボタン終了判定
    If CloseMode = vbFormControlMenu Then

        'Cancel = True
        '# 関数管理画面表示
        frmFunctionManager.Show vbModeless

    End If

End Sub

' @JPName
' PathChart解析実行
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 指定開始関数からPathChart解析を実行する
'
' @Remarks
' 解析前に開始関数および
' 最大深度設定を保存する
' 実行結果をログと画面へ反映する
'
Private Sub cmdAnalyze_Click()

    ' 開始時刻
    Dim StartTime As Double
    
    ' 開始Procedure
    Dim StartProc As String
    
    ' 最大探索深度
    Dim MaxDepth As Long

    On Error GoTo EH

    '# 開始Procedure取得
    StartProc = Trim$(cmbStartProc.Text)

    ' 開始Procedure未選択チェック
    '# 開始Procedure確認
    If StartProc = "" Then

        MsgBox _
            "開始Procedureを選択してください。", _
            vbExclamation

        Exit Sub

    End If

    '# 最大深度取得
    MaxDepth = CLng(txtMaxDepth.Text)
    
    '# PathChart条件設定
    gRecursive = chkRecursive.Value

    '# 外部呼出設定
    gExternal = chkExternal.Value

    '# 最大深度確認
    If MaxDepth <= 0 Then

        MsgBox _
            "最大深度は1以上を指定してください。", _
            vbExclamation

        Exit Sub

    End If

    '=======================================
    '# 解析状態初期化
    '=======================================
    StartTime = Timer

    '# ログ初期化
    lstLog.Clear

    lblStatus.Caption = "状態： 解析中"

    '# 統計表示初期化
    lblNodeCount.Caption = "Node数： 0"
    lblEdgeCount.Caption = "Edge数： 0"
    lblDepth.Caption = "最大深度： 0"

    frmPathChart.AddLog String(40, "=")

    '# 解析開始ログ出力
    AddLog "解析開始"
    AddLog "開始Procedure ： " & StartProc
    AddLog "最大深度 ： " & MaxDepth

    DoEvents

    '=======================================
    '# 最大深度保存
    '=======================================
    SaveDefaultMaxDepth MaxDepth

    '=======================================
    '# 開始Procedure保存
    '=======================================
    SaveDefaultStartProcedure StartProc

    '=======================================
    '# PathChart解析実行
    '=======================================
    ExecutePathChart _
        StartProc, _
        MaxDepth

    '=======================================
    '結果表示
    '# 実行時間表示
    '=======================================
    lblElapsed.Caption = _
        "解析時間: " & _
        Format(Timer - StartTime, "0.00 秒")

    '# 解析完了表示
    lblStatus.Caption = "状態： 完了"

    '# 解析完了ログ出力
    AddLog "解析完了"

    Exit Sub
    
'# 異常終了処理
EH:

    lblStatus.Caption = _
        "状態： 異常終了"

    '# エラーログ出力
    AddLog _
        "ERROR ： " & Err.Description

    '# エラー通知
    MsgBox Err.Description, vbCritical

End Sub

' @JPName
' ログクリア
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ログおよび統計情報を初期化する
'
' @Remarks
' ログ一覧および統計情報表示を
' 初期状態へ戻す
'
Private Sub cmdClear_Click()

    '# ログ初期化
    lstLog.Clear

    '# 状態表示初期化
    lblStatus.Caption = "状態： 待機中"
    lblElapsed.Caption = "解析時間： 0.00 秒"

    '# 統計表示初期化
    lblNodeCount.Caption = "Node数： 0"
    lblEdgeCount.Caption = "Edge数： 0"
    lblDepth.Caption = "最大深度： 0"

End Sub

' @JPName
' オプション表示
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChartオプション画面を表示する
'
' @Remarks
' PathChart設定変更用の
' オプション画面を表示する
'
Private Sub cmdOption_Click()

    '# オプション画面表示
    frmPathChartOption.Show vbModal

End Sub

' @JPName
' ナビゲータ表示
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChartナビゲータを表示する
'
' @Remarks
' 検索候補を再読込して表示する
'
Private Sub cmdNavigator_Click()

    '# 検索候補初期化
    frmPathChartNavigator.cmbSearch.Clear
    
    '# 検索候補読込
    LoadSearchNodeList _
        frmPathChartNavigator.cmbSearch
    
    '# PathChartナビゲータ表示
    frmPathChartNavigator.Show _
        vbModeless
    
    '# PathChart画面非表示
    Me.Hide

End Sub

' @JPName
' ログ追加
'
' @Category
' PathChart
'
' @Input
' Msg(String)
'
' @Output
' なし
'
' @Summary
' ログ一覧へメッセージを追加する
'
' @Remarks
' ログ追加後に末尾行を選択し
' 最新ログが表示されるようにする
'
Public Sub AddLog(ByVal Msg As String)

    '# ログ追加
    lstLog.AddItem _
        Format(Now, "hh:nn:ss") & _
        "  " & Msg

    '# 最新ログ選択
    If lstLog.ListCount > 0 Then

        lstLog.ListIndex = lstLog.ListCount - 1

    End If

    '# 画面更新
    DoEvents

End Sub
