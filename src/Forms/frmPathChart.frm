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

    LoadPathChartConfig
    
    chkExternal.Value = gExternal
    
    chkRecursive.Value = gRecursive

    txtMaxDepth.Text = GetDefaultMaxDepth()
    
    LoadStartProcedure cmbStartProc

    Dim i As Long
    Dim DefaultProc As String

    DefaultProc = GetDefaultStartProcedure()

    Dim Found As Boolean
    
    For i = 0 To cmbStartProc.ListCount - 1
    
        If cmbStartProc.List(i) = DefaultProc Then

            cmbStartProc.ListIndex = i
            
            Found = True

            Exit For

        End If

    Next i
    
    If Not Found Then

        If cmbStartProc.ListCount > 0 Then

            cmbStartProc.ListIndex = 0
        
        End If
        
    End If

    If cmbStartProc.ListIndex < 0 _
    And cmbStartProc.ListCount > 0 Then

        cmbStartProc.ListIndex = 0

    End If

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
' JPName、Category、Summaryを更新する
'
Private Sub cmbStartProc_Change()

    Dim Proc As clsProcInfo
    Dim ProcName As String
    
    ProcName = Trim$(cmbStartProc.Value)
    
    If ProcName = "" Then Exit Sub

    Set Proc = GetProcInfo(ProcName)
    
    If Not Proc Is Nothing Then
      
        If Proc.JPName <> "" Then
            lblSelectProc.Caption = Proc.JPName
        Else
            lblSelectProc.Caption = ProcName
        End If
        
        lblCategory.Caption = Proc.Category
        txtSummary.Text = Proc.Summary

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

    If CloseMode = vbFormControlMenu Then

        'Cancel = True
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
' 最大深度保存および
' 開始関数保存を行う
'
Private Sub cmdAnalyze_Click()

    Dim StartTime As Double
    Dim StartProc As String
    Dim MaxDepth As Long

    On Error GoTo EH

    StartProc = Trim$(cmbStartProc.Text)

    ' 開始Procedure未選択チェック
    If StartProc = "" Then

        MsgBox _
            "開始Procedureを選択してください。", _
            vbExclamation

        Exit Sub

    End If

    MaxDepth = CLng(txtMaxDepth.Text)
    
    gRecursive = chkRecursive.Value

    gExternal = chkExternal.Value

    ' 最大深度チェック
    If MaxDepth <= 0 Then

        MsgBox _
            "最大深度は1以上を指定してください。", _
            vbExclamation

        Exit Sub

    End If

    '=======================================
    '初期化
    '=======================================
    StartTime = Timer

    lstLog.Clear

    lblStatus.Caption = "状態： 解析中"

    lblNodeCount.Caption = "Node数： 0"
    lblEdgeCount.Caption = "Edge数： 0"
    lblDepth.Caption = "最大深度： 0"

    frmPathChart.AddLog String(40, "=")

    AddLog "解析開始"
    AddLog "開始Procedure ： " & StartProc
    AddLog "最大深度 ： " & MaxDepth

    DoEvents

    '=======================================
    '最大深度保存実行
    '=======================================
    SaveDefaultMaxDepth MaxDepth

    '=======================================
    '開Procedure設定実行
    '=======================================
    SaveDefaultStartProcedure StartProc

    '=======================================
    '解析実行
    '=======================================
    ExecutePathChart _
        StartProc, _
        MaxDepth

    '=======================================
    '結果表示
    '=======================================
    lblElapsed.Caption = _
        "解析時間: " & _
        Format(Timer - StartTime, "0.00 秒")

    lblStatus.Caption = "状態： 完了"

    AddLog "解析完了"

    Exit Sub

EH:

    lblStatus.Caption = _
        "状態： 異常終了"

    AddLog _
        "ERROR ： " & Err.Description

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
Private Sub cmdClear_Click()

    lstLog.Clear

    lblStatus.Caption = "状態： 待機中"
    lblElapsed.Caption = "解析時間： 0.00 秒"

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
Private Sub cmdOption_Click()

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

    frmPathChartNavigator.cmbSearch.Clear
    
    LoadSearchNodeList _
        frmPathChartNavigator.cmbSearch
    
    frmPathChartNavigator.Show _
        vbModeless
    
    Me.Hide

End Sub

' @JPName
' PathChart既定値初期化
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
' PathChartの既定設定を初期化する
'
' @Remarks
' 初回起動時の設定値を設定する
'
Public Sub InitPathChartOption()

    gProperty = False
    gPrivate = False
    gPublic = True

    gStdModule = True
    gUserForm = True
    gClassModule = True

    gIgnoreAPI = True
    gIgnoreExcel = False
    gIgnoreSelf = False
    gSameModuleOnly = False

    gColorTheme = DEFAULT_COLOR_THEME

    gMaxNode = DEFAULT_MAX_NODE
    gMaxEdge = DEFAULT_MAX_EDGE
    
    InitSearchHistory

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
' 追加時に最新行へスクロールする
'
Public Sub AddLog(ByVal Msg As String)

    lstLog.AddItem _
        Format(Now, "hh:nn:ss") & _
        "  " & Msg

    If lstLog.ListCount > 0 Then

        lstLog.ListIndex = lstLog.ListCount - 1

    End If

    DoEvents

End Sub
