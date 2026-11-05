VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTargetTool 
   Caption         =   "関数管理対象選択"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTargetTool.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTargetTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @JPName
' ツール選択画面初期化
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ツール選択画面を初期化する
'
' @Remarks
' ToolManagerシートから
' ツール一覧を読み込む
'
Private Sub UserForm_Initialize()

    '# フォーム初期化
    Me.Caption = "関数管理対象選択"
    
    '# ツール一覧表示設定
    lstTool.ColumnCount = 2
    lstTool.ColumnWidths = "180 pt;0 pt"

    '# ツール一覧読込
    LoadToolList

End Sub

' @JPName
' ツール一覧読込
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ToolManagerシートから
' ツール一覧を読み込む
'
' @Remarks
' ツール名を表示し
' ToolIDを内部保持する
'
Private Sub LoadToolList()

    ' ToolManagerシート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    '# ToolManagerシート取得
    Set ws = _
        ThisWorkbook.Worksheets(SHEET_TOOL_MANAGER)

    '# Tool一覧最終行取得
    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    '# ツール一覧初期化
    lstTool.Clear

    '# ツール一覧生成
    For r = 2 To LastRow

        '# ツール名追加
        lstTool.AddItem _
            ws.Cells(r, "B").Value

        '# ToolID保持
        lstTool.List( _
            lstTool.ListCount - 1, 1) = _
            ws.Cells(r, "A").Value

    Next r

End Sub


' @JPName
' ツール選択確定
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択したツールを解析対象として設定する
'
' @Remarks
' 解析コンテキスト生成後に
' Repositoryを初期化し
' 関数管理画面を表示する
'
Private Sub cmdOK_Click()

    '# 対象ツール選択確認
    If lstTool.ListIndex < 0 Then

        MsgBox _
            "対象ツールを選択してください。"

        Exit Sub

    End If

    ' 選択ツールIDを取得し
    ' 解析コンテキストを生成する
    '# 選択ツール取得
    gTargetToolID = _
        CLng(lstTool.List( _
            lstTool.ListIndex, _
            1))

    '# 解析コンテキスト生成
    Set gContext = _
        CreateAnalyzeContext()
    
    '# 対象ブック確認
    If Not ValidateTargetBook("対象ツール選択") Then
        Exit Sub
    End If
    
    '# ProcedureRepository初期化解除
    ClearProcedureRepository
    
    '# ツール選択画面終了
    Unload Me

    ' 管理ファイル終了時に
    ' FunctionManagerが終了しないよう
    ' 解析ツールをアクティブ化する
    '# AnalyzerBookアクティブ化
    ActivateAnalyzerBook
    
    '# 関数管理画面表示
    frmFunctionManager.Show vbModeless

End Sub

' @JPName
' ツール選択終了
'
' @Category
' UI
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ツール選択画面を閉じる
'
' @Remarks
' ツール選択画面を閉じた後
' メインメニューを再表示する
'
Private Sub cmdClose_Click()

    '# ツール選択画面終了
    Unload Me
    
    '# メインメニュー再表示
    frmMainMenu.Show vbModeless
    
End Sub

