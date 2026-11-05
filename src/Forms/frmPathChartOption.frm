VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPathChartOption 
   Caption         =   "PathChart詳細設定"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   OleObjectBlob   =   "frmPathChartOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPathChartOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' @JPName
' オプション初期化
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
' PathChartオプション画面を初期化する
'
' @Remarks
' 現在設定値を画面へ反映する
'
Private Sub UserForm_Initialize()

    '# 制限値設定読込
    If gMaxNode = 0 Then

        txtMaxNode.Text = "1000"
        txtMaxEdge.Text = "3000"
        
    Else
        txtMaxNode.Text = gMaxNode
        txtMaxEdge.Text = gMaxEdge

    End If

    '# 解析対象設定読込
    chkProperty.Value = gProperty
    chkPrivate.Value = gPrivate
    chkPublic.Value = gPublic

    '# モジュール種別設定読込
    chkStdModule.Value = gStdModule
    chkUserForm.Value = gUserForm
    chkClassModule.Value = gClassModule

    '# 除外条件設定読込
    chkIgnoreAPI.Value = gIgnoreAPI
    chkIgnoreExcel.Value = gIgnoreExcel
    chkIgnoreSelf.Value = gIgnoreSelf
    chkSameModuleOnly.Value = gSameModuleOnly

    '# PathChart表示設定読込
    chkSkipCommon.Value = gSkipCommon
    chkShowCount.Value = gShowCount

    '# 出力設定読込
    chkPowerPoint.Value = gPowerPoint
    
    '# PowerPoint向き設定読込
    Select Case gPptOrientation

        Case PPT_ORIENT_PORTRAIT

            optPptPortrait.Value = True

        Case PPT_ORIENT_LANDSCAPE

            optPptLandscape.Value = True

        Case Else

            optPptAuto.Value = True

    End Select
    
    chkPngOutput.Value = gPngOutput

    '# モジュール表示設定読込
    chkShowModule.Value = gShowModule
    chkColorModule.Value = gColorModule

    '# テーマ候補設定
    cmbColorTheme.Clear

    cmbColorTheme.AddItem "標準"
    cmbColorTheme.AddItem "明るい"
    cmbColorTheme.AddItem "ダーク"

    '# テーマ設定反映
    If gColorTheme = "" Then

        cmbColorTheme.ListIndex = 0

    Else

        cmbColorTheme.Value = gColorTheme

    End If

    '# 出力方向設定状態反映
    Call chkPowerPoint_Click
    
End Sub

' @JPName
' オプション取消
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
' オプション画面を閉じる
'
' @Remarks
' 設定変更を反映せずに
' オプション画面を閉じる
'
Private Sub cmdCancel_Click()

    '# オプション画面終了
    Unload Me

End Sub

' @JPName
' オプション保存
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
' オプション設定を保存する
'
' @Remarks
' 画面設定をグローバル設定へ反映し
' PathChart設定を保存する
'
Private Sub cmdOK_Click()

    ' 設定値保存
    
    '# 解析対象設定保存
    gProperty = chkProperty.Value
    gPrivate = chkPrivate.Value
    gPublic = chkPublic.Value

    '# モジュール種別設定保存
    gStdModule = chkStdModule.Value
    gUserForm = chkUserForm.Value
    gClassModule = chkClassModule.Value

    '# 除外条件設定保存
    gIgnoreAPI = chkIgnoreAPI.Value
    gIgnoreExcel = chkIgnoreExcel.Value
    gIgnoreSelf = chkIgnoreSelf.Value
    gSameModuleOnly = chkSameModuleOnly.Value

    '# PathChart表示設定保存
    gSkipCommon = chkSkipCommon.Value
    gShowCount = chkShowCount.Value

    '# 出力設定保存
    gPowerPoint = chkPowerPoint.Value
    
    '# PowerPoint向き設定保存
    SetOptionValue
    
    gPngOutput = chkPngOutput.Value

    '# モジュール表示設定保存
    gShowModule = chkShowModule.Value
    gColorModule = chkColorModule.Value

    '# テーマ設定保存
    gColorTheme = cmbColorTheme.Value

    '# 制限値設定保存
    gMaxNode = Val(txtMaxNode.Text)
    gMaxEdge = Val(txtMaxEdge.Text)

    '# PathChart設定保存
    SavePathChartConfig
    
    '# オプション画面終了
    Me.Hide

End Sub

' @JPName
' 既定値設定
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
' 画面の全オプションを初期状態へ戻す
'
' @Remarks
' PathChartオプションを
' 既定値へ戻す
'
Private Sub cmdDefault_Click()

    '# PathChart表示設定初期化
    chkSkipCommon.Value = False
    chkShowCount.Value = False

    '# 出力設定初期化
    chkPowerPoint.Value = False
    chkPngOutput.Value = False

    '# モジュール表示設定初期化
    chkShowModule.Value = False
    chkColorModule.Value = False

    '# テーマ設定初期化
    cmbColorTheme.Value = DEFAULT_COLOR_THEME

    '# 制限値初期化
    txtMaxNode.Text = DEFAULT_MAX_NODE
    txtMaxEdge.Text = DEFAULT_MAX_EDGE

End Sub

' @JPName
' 出力オプション設定
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
' PowerPoint出力方向設定を反映する
'
' @Remarks
' 選択された出力方向を
' グローバル設定へ反映する
'
Private Sub SetOptionValue()

    '# PowerPoint向き設定
    If optPptAuto.Value Then

        '# 自動向き設定
        gPptOrientation = PPT_ORIENT_AUTO

    ElseIf optPptPortrait.Value Then

        '# 縦向き設定
        gPptOrientation = PPT_ORIENT_PORTRAIT

    ElseIf optPptLandscape.Value Then

        '# 横向き設定
        gPptOrientation = PPT_ORIENT_LANDSCAPE

    End If

End Sub

' @JPName
' PowerPoint出力切替
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
' PowerPoint出力設定に応じて
' 方向選択を有効化する
'
' @Remarks
' PowerPoint出力ON時のみ
' 出力方向選択を有効にする
'
Private Sub chkPowerPoint_Click()

    '# PowerPoint向き選択制御
    optPptAuto.Enabled = _
        chkPowerPoint.Value

    '# 縦向き選択制御
    optPptPortrait.Enabled = _
        chkPowerPoint.Value

    '# 横向き選択制御
    optPptLandscape.Enabled = _
        chkPowerPoint.Value

End Sub

