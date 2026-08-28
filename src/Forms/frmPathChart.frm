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

Private Sub UserForm_Initialize()

    cmbStartProc.Clear

    cmbStartProc.AddItem "UpdateItemsInHardCoded"

    cmbStartProc.ListIndex = 0

End Sub

' 解析実行
Private Sub cmdAnalyze_Click()

    Dim StartTime As Double

    On Error GoTo EH

    '=======================================
    '入力チェック
    '=======================================
    If Trim$(cmbStartProc.Text) = "" Then

        MsgBox "開始Procedureを選択してください。", _
               vbExclamation

        cmbStartProc.SetFocus

        Exit Sub

    End If

    If Val(txtMaxDepth.Text) <= 0 Then

        MsgBox "最大深度を入力してください。", _
               vbExclamation

        txtMaxDepth.SetFocus

        Exit Sub

    End If

    '=======================================
    '初期化
    '=======================================
    StartTime = Timer

    lstLog.Clear

    lblStatus.Caption = "状態: 解析中"
    lblNodeCount.Caption = "Node数: 0"
    lblEdgeCount.Caption = "Edge数: 0"
    lblDepth.Caption = "最大深度: 0"

    lstLog.AddItem "解析開始"
    lstLog.AddItem "開始Procedure : " _
                 & cmbStartProc.Text

    lstLog.AddItem "最大深度 : " _
                 & txtMaxDepth.Text

    DoEvents

    '=======================================
    '解析実行
    '=======================================
    ExecutePathChart _
        cmbStartProc.Text, _
        CLng(txtMaxDepth.Text)

    '=======================================
    '結果表示
    '=======================================
    lblElapsed.Caption = _
        "解析時間: " & _
        Format(Timer - StartTime, "0.00 秒")

    lblStatus.Caption = "状態: 完了"

    lstLog.AddItem "解析完了"

    Exit Sub

EH:

    lblStatus.Caption = "状態: 異常終了"

    lstLog.AddItem _
        "ERROR : " & Err.Description

    MsgBox Err.Description, vbCritical

End Sub

' クリア
Private Sub cmdClear_Click()

    lstLog.Clear

    lblStatus.Caption = "状態: 待機中"
    lblElapsed.Caption = "解析時間: 0.00 秒"

    lblNodeCount.Caption = "Node数: 0"
    lblEdgeCount.Caption = "Edge数: 0"
    lblDepth.Caption = "最大深度: 0"

End Sub

Private Sub cmdOption_Click()

    frmPathChartOption.Show vbModal

End Sub

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

    gVertical = True

    gColorTheme = "標準"

    gMaxNode = 1000
    gMaxEdge = 3000

End Sub

Public Sub AddLog(ByVal Msg As String)

    lstLog.AddItem _
        Format(Now, "hh:nn:ss") & _
        "  " & Msg

    If lstLog.ListCount > 0 Then

        lstLog.ListIndex = lstLog.ListCount - 1

    End If

    DoEvents

End Sub

