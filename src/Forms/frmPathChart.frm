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

    lblStatus.Caption = "状態: 待機中"
    
    lblElapsed.Caption = "解析時間: 0.00 秒"

End Sub

' 解析実行
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

    lblStatus.Caption = "状態: 解析中"

    lblNodeCount.Caption = "Node数: 0"
    lblEdgeCount.Caption = "Edge数: 0"
    lblDepth.Caption = "最大深度: 0"

    frmPathChart.AddLog String(40, "=")

    AddLog "解析開始"
    AddLog "開始Procedure : " & StartProc
    AddLog "最大深度 : " & MaxDepth

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

    lblStatus.Caption = "状態: 完了"

    AddLog "解析完了"

    Exit Sub

EH:

    lblStatus.Caption = _
        "状態: 異常終了"

    AddLog _
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

