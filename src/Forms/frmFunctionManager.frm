VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFunctionManager 
   Caption         =   "関数管理"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3435
   OleObjectBlob   =   "frmFunctionManager.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFunctionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' @JPName
' 画面初期化
'
' @Category
' FunctionManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 現在選択中ツール情報を表示する
'
' @Remarks
' ToolNameおよびWorkbook名を
' 画面へ表示する
'
Private Sub UserForm_Initialize()

    Me.Caption = "関数管理"

    ' 現在選択中ツール情報
    Dim Tool As clsToolInfo

    Set Tool = GetCurrentToolInfo()

    If Tool Is Nothing Then

        MsgBox "ツール情報取得失敗", _
                vbExclamation

        Exit Sub

    End If

    lblTool.Caption = _
        "対象 : " & Tool.ToolName & vbCrLf & _
        "Book : " & Tool.WorkbookName

End Sub

' @JPName
' 解析実行
'
' @Category
' FunctionManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 関数解析処理を実行する
'
' @Remarks
' 関数一覧解析実行後に
' ProcedureRepositoryを初期化する
'
Private Sub cmdAnalyze_Click()

    If Not ValidateTargetBook("関数解析") Then
        Exit Sub
    End If

    AnalyzeAll

    ClearProcedureRepository

End Sub

' @JPName
' HCP生成起動
'
' @Category
' FunctionManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' HCP生成画面を表示する
'
' @Remarks
' ProcedureRepositoryを初期化後
' HCP生成フォームを表示する
'
Private Sub cmdHCP_Click()

    If Not ValidateTargetBook("HCP生成") Then
        Exit Sub
    End If

    InitializeProcedureRepository

    frmHCPGenerator.Show vbModeless

End Sub

' @JPName
' PathChart起動
'
' @Category
' FunctionManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChart画面を表示する
'
' @Remarks
' ProcedureRepositoryを初期化後
' PathChartフォームを表示する
'
Private Sub cmdPathChart_Click()

    If Not ValidateTargetBook("PathChart") Then
        Exit Sub
    End If

    InitializeProcedureRepository

    frmPathChart.Show vbModeless

End Sub

' @JPName
' 画面終了
'
' @Category
' FunctionManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 関数管理画面を閉じる
'
' @Remarks
' 関数管理画面を閉じた後
' メインメニューを再表示する
'
Private Sub cmdClose_Click()

    Unload Me
    
    frmMainMenu.Show vbModeless
    'frmTarget.Show vbModeless

End Sub
