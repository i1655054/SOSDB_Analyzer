VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPathChartNavigator 
   Caption         =   "PathChart Navigator"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmPathChartNavigator.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPathChartNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @JPName
' ナビゲータ初期化
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
' PathChartナビゲータの初期設定を行う
'
' @Remarks
' 検索履歴および入力補完設定を初期化する
'
Private Sub UserForm_Initialize()

    '# 検索履歴読込
    LoadSearchHistory cmbSearch
    
    '# 入力補完設定
    cmbSearch.MatchEntry = fmMatchEntryComplete
    
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
' ナビゲータ表示位置を設定する
'
' @Remarks
' Excel画面右上へ配置し
' ナビゲータ表示内容を初期化する
'
Private Sub UserForm_Activate()

    '# ナビゲータ表示位置設定
    Me.Left = _
        Application.Left + _
        Application.Width - _
        Me.Width - 30

    Me.Top = _
        Application.Top + 80
    
    '# ノード情報初期化
    Call ClearNavigatorInfo
    
End Sub

' @JPName
' ナビゲータ終了制御
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
' ナビゲータ終了時の制御を行う
'
' @Remarks
' ×ボタン押下時は
' PathChart画面を再表示する
'
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    '# ×ボタン終了判定
    If CloseMode = vbFormControlMenu Then

        'Cancel = True
        
        '# PathChart画面再表示
        frmPathChart.Show vbModeless

    End If

End Sub

' @JPName
' ノード検索
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
' 指定文字列でノード検索を実行する
'
' @Remarks
' 検索履歴へ追加後に検索する
'
Private Sub cmdFind_Click()

    '# 検索文字列取得
    gSearchText = GetProcNameFromDisplay( _
                    Trim$(cmbSearch.Text))
    
    '# 完全一致検索設定
    gExactMatch = chkExactMatch.Value
    
    '# 検索履歴登録
    AddSearchHistory gSearchText

    '# ノード検索実行
    FindNode

End Sub

' @JPName
' 次候補検索
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
' 次の検索候補を検索する
'
' @Remarks
' 前回検索位置を基準に
' 次の一致ノードを検索する
'
Private Sub cmdNext_Click()

    '# 完全一致検索設定
    gExactMatch = chkExactMatch.Value
    
    '# 次候補検索実行
    FindNextNode

End Sub

' @JPName
' 親ノード移動
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
' 親ノードへ移動する
'
' @Remarks
' 現在選択ノードの親ノードへ移動する
'
Private Sub cmdParent_Click()

    '# 親ノード移動
    MoveToParentNode

End Sub

' @JPName
' 子ノード移動
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
' 子ノードへ移動する
'
' @Remarks
' 現在選択ノードの子ノードへ移動する
'
Private Sub cmdChild_Click()

    '# 子ノード移動
    MoveToChildNode

End Sub

' @JPName
' 兄弟ノード移動
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
' 兄弟ノードへ移動する
'
' @Remarks
' 現在選択ノードの兄弟ノードへ移動する
'
Private Sub cmdSibling_Click()

    '# 兄弟ノード移動
    MoveToSiblingNode

End Sub

' @JPName
' 強調解除
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
' ノード強調表示を解除する
'
' @Remarks
' ノードおよび接続線の
' 強調表示を解除する
'
Private Sub cmdClear_Click()

    '# 強調表示解除
    ClearHighlight

End Sub

' @JPName
' ナビゲータ情報初期化
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
' ナビゲータ表示情報を初期化する
'
' @Remarks
' ナビゲータ上の
' ノード情報表示をクリアする
'
Public Sub ClearNavigatorInfo()

    '# 現在ノード表示初期化
    txtCurrentNode.Text = "-"

    '# Procedure情報表示初期化
    lblJPName.Caption = ""

    lblCategory.Caption = ""

    lblModule.Caption = ""

    '# 概要表示初期化
    txtSummary.Text = ""

    '# 状態表示初期化
    txtStatus.Text = ""

End Sub

' @JPName
' 現在情報コピー
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
' 現在ノード情報をクリップボードへコピーする
'
' @Remarks
' 現在ノードの表示情報を
' クリップボードへコピーする
'
Private Sub cmdCopyCurrentInfo_Click()

    ' コピー対象文字列
    Dim S As String

    '# コピー文字列生成
    S = _
        "Procedure : " & txtCurrentNode.Text & vbCrLf & _
        "JPName    : " & lblJPName.Caption & vbCrLf & _
        "Category  : " & lblCategory.Caption & vbCrLf & _
        "Module    : " & lblModule.Caption & vbCrLf & _
        "Summary   : " & txtSummary.Text

    '# 現在ノード情報コピー
    Call CopyTextToClipboard(S)

    '# コピー完了表示
    Application.StatusBar = "現在情報をコピーしました"
    
End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' TextValue(String)
' Title(String)
'
' @Output
' なし
'
' @Summary
' 指定文字列をクリップボードへコピーする
'
' @Remarks
' ステータスバーへコピー結果を表示する
'
Private Sub CopyControlValue( _
                ByVal TextValue As String, _
                Optional ByVal Title As String = "")

    '# コピー文字列確認
    If Trim$(TextValue) = "" Then Exit Sub

    '# クリップボードコピー
    Call CopyTextToClipboard(TextValue)

    '# コピー結果表示
    If Title <> "" Then

        Application.StatusBar = _
            "コピーしました : " & _
            Title

    Else

        Application.StatusBar = _
            "コピーしました : " & _
            TextValue

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub txtCurrentNode_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# 現在ノードコピー
        CopyControlValue _
            txtCurrentNode.Text, _
            "現在ノード"

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub txtStatus_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# 状態コピー
        CopyControlValue _
            txtStatus.Text, _
            "状態"

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub lblJPName_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# 日本語名コピー
        CopyControlValue _
            lblJPName.Caption, _
            "日本語名"

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub lblCategory_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# カテゴリコピー
        CopyControlValue _
            lblCategory.Caption, _
            "カテゴリ"

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub lblModule_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# モジュール名コピー
        CopyControlValue _
            lblModule.Caption, _
            "モジュール"

    End If

End Sub

' @JPName
' 値コピー
'
' @Category
' PathChart
'
' @Input
' Button(Integer)
' Shift(Integer)
' X(Single)
' Y(Single)
'
' @Output
' なし
'
' @Summary
' 右クリック時に表示値をコピーする
'
' @Remarks
' 右クリック時に表示内容を
' クリップボードへコピーする
'
Private Sub txtSummary_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal x As Single, _
        ByVal y As Single)

    '# 右クリック判定
    If Button = 2 Then

        '# 概要コピー
        CopyControlValue _
            txtSummary.Text, _
            "概要"

    End If

End Sub

