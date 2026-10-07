Attribute VB_Name = "modPathChartConfig"
Option Explicit

'==================================================
' PowerPoint出力定数
'==================================================

Public Const PPT_ORIENT_AUTO As String = "AUTO"
Public Const PPT_ORIENT_PORTRAIT As String = "PORTRAIT"
Public Const PPT_ORIENT_LANDSCAPE As String = "LANDSCAPE"

'==================================================
' PathChartConfigセル定義
'==================================================

Private Const CFG_SKIP_COMMON      As String = "B1"
Private Const CFG_SHOW_COUNT       As String = "B2"

Private Const CFG_POWERPOINT       As String = "B3"
Private Const CFG_PPT_ORIENTATION  As String = "B4"

Private Const CFG_PNG_OUTPUT       As String = "B5"

Private Const CFG_SHOW_MODULE      As String = "B6"
Private Const CFG_COLOR_MODULE     As String = "B7"

Private Const CFG_PROPERTY         As String = "B8"
Private Const CFG_PRIVATE          As String = "B9"
Private Const CFG_PUBLIC           As String = "B10"

Private Const CFG_STD_MODULE       As String = "B11"
Private Const CFG_USER_FORM        As String = "B12"
Private Const CFG_CLASS_MODULE     As String = "B13"

Private Const CFG_IGNORE_API       As String = "B14"
Private Const CFG_IGNORE_EXCEL     As String = "B15"
Private Const CFG_IGNORE_SELF      As String = "B16"
Private Const CFG_SAME_MODULE_ONLY As String = "B17"

Private Const CFG_COLOR_THEME      As String = "B18"

Private Const CFG_MAX_NODE         As String = "B19"
Private Const CFG_MAX_EDGE         As String = "B20"
Private Const CFG_MAX_DEPTH        As String = "B21"
Private Const CFG_START_PROC       As String = "B22"
Private Const CFG_EXTERNAL         As String = "B23"
Private Const CFG_RECURSIVE        As String = "B24"

Private Const CFG_SHOW_JPNAME      As String = "B25"
Private Const CFG_SHOW_CATEGORY    As String = "B26"

'==================================================
' PathChart既定値
'==================================================

Public Const DEFAULT_MAX_NODE As Long = 1000

Public Const DEFAULT_MAX_EDGE As Long = 3000

Public Const DEFAULT_MAX_DEPTH As Long = 5

Public Const DEFAULT_COLOR_THEME As String = "標準"

'==================================================
' 実行時設定
'==================================================

'--------------------------------------------------
' 解析対象
'--------------------------------------------------
Public gProperty As Boolean
Public gPrivate As Boolean
Public gPublic As Boolean

Public gStdModule As Boolean
Public gUserForm As Boolean
Public gClassModule As Boolean

'--------------------------------------------------
' 除外条件
'--------------------------------------------------
Public gIgnoreAPI As Boolean
Public gIgnoreExcel As Boolean
Public gIgnoreSelf As Boolean
Public gSameModuleOnly As Boolean

'--------------------------------------------------
'出力設定
'--------------------------------------------------
Public gColorTheme As String

'--------------------------------------------------
'制限
'--------------------------------------------------
Public gMaxNode As Long
Public gMaxEdge As Long

'共通関数除外
Public gSkipCommon As Boolean
'呼出回数表示
Public gShowCount As Boolean

'PowerPoint出力
Public gPowerPoint As Boolean
'PowerPoint出力オプション
Public gPptOrientation As String

'PNG出力
Public gPngOutput As Boolean

'モジュール表示
Public gShowModule As Boolean
'モジュール色分け
Public gColorModule As Boolean

'検索文字列
Public gSearchText As String

'現在ノード名
Public gCurrentNodeText As String

'検索履歴文字列
Public gSearchHistory As Collection

'検索完全一致表示
Public gExactMatch As Boolean

'再帰呼出表示
' OFF は、A→Aの自己再帰を除外
' ON  は、A→Aを表示
Public gRecursive As Boolean

'外部モジュール呼出表示
' OFF は、同一モジュール内のみ modA → modA を表示
' ON  は、modA → modBも表示
Public gExternal As Boolean

Public gShowJPName As Boolean
Public gShowCategory As Boolean


'==================================================
' 設定保存・復元
'==================================================

' @JPName
' PathChart設定保存
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
' PathChart設定を
' PathChartConfigシートへ保存する
'
' @Remarks
' 出力設定、解析条件、
' 表示設定を保存する
'
Public Sub SavePathChartConfig()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    ws.Range(CFG_SKIP_COMMON).Value = gSkipCommon
    ws.Range(CFG_SHOW_COUNT).Value = gShowCount

    ws.Range(CFG_POWERPOINT).Value = gPowerPoint
    ws.Range(CFG_PPT_ORIENTATION).Value = gPptOrientation
    
    ws.Range(CFG_PNG_OUTPUT).Value = gPngOutput

    ws.Range(CFG_SHOW_MODULE).Value = gShowModule
    ws.Range(CFG_COLOR_MODULE).Value = gColorModule

    ws.Range(CFG_PROPERTY).Value = gProperty
    ws.Range(CFG_PRIVATE).Value = gPrivate
    ws.Range(CFG_PUBLIC).Value = gPublic

    ws.Range(CFG_STD_MODULE).Value = gStdModule
    ws.Range(CFG_USER_FORM).Value = gUserForm
    ws.Range(CFG_CLASS_MODULE).Value = gClassModule

    ws.Range(CFG_IGNORE_API).Value = gIgnoreAPI
    ws.Range(CFG_IGNORE_EXCEL).Value = gIgnoreExcel
    ws.Range(CFG_IGNORE_SELF).Value = gIgnoreSelf
    ws.Range(CFG_SAME_MODULE_ONLY).Value = gSameModuleOnly

    ws.Range(CFG_COLOR_THEME).Value = gColorTheme

    ws.Range(CFG_MAX_NODE).Value = gMaxNode
    ws.Range(CFG_MAX_EDGE).Value = gMaxEdge

    ws.Range(CFG_EXTERNAL).Value = gExternal
    ws.Range(CFG_RECURSIVE).Value = gRecursive

    ws.Range(CFG_SHOW_JPNAME).Value = gShowJPName
    ws.Range(CFG_SHOW_CATEGORY).Value = gShowCategory

End Sub

' @JPName
' PathChart設定読込
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
' PathChartConfigシートから
' PathChart設定を読み込む
'
' @Remarks
' 未設定項目には既定値を設定する
'
Public Sub LoadPathChartConfig()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    gSkipCommon = ws.Range(CFG_SKIP_COMMON).Value
    gShowCount = ws.Range(CFG_SHOW_COUNT).Value

    gPowerPoint = ws.Range(CFG_POWERPOINT).Value
    gPptOrientation = Trim$(ws.Range(CFG_PPT_ORIENTATION).Value)
    
    If gPptOrientation = "" Then
    
        gPptOrientation = PPT_ORIENT_AUTO
        
    End If
    

    gPngOutput = ws.Range(CFG_PNG_OUTPUT).Value

    gShowModule = ws.Range(CFG_SHOW_MODULE).Value
    gColorModule = ws.Range(CFG_COLOR_MODULE).Value

    gProperty = ws.Range(CFG_PROPERTY).Value
    gPrivate = ws.Range(CFG_PRIVATE).Value
    gPublic = ws.Range(CFG_PUBLIC).Value

    gStdModule = ws.Range(CFG_STD_MODULE).Value
    gUserForm = ws.Range(CFG_USER_FORM).Value
    gClassModule = ws.Range(CFG_CLASS_MODULE).Value

    gIgnoreAPI = ws.Range(CFG_IGNORE_API).Value
    gIgnoreExcel = ws.Range(CFG_IGNORE_EXCEL).Value
    gIgnoreSelf = ws.Range(CFG_IGNORE_SELF).Value
    gSameModuleOnly = ws.Range(CFG_SAME_MODULE_ONLY).Value

    gColorTheme = ws.Range(CFG_COLOR_THEME).Value

    gMaxNode = Val(ws.Range(CFG_MAX_NODE).Value)
    gMaxEdge = Val(ws.Range(CFG_MAX_EDGE).Value)

    gExternal = ws.Range(CFG_EXTERNAL).Value

    If Trim$(ws.Range(CFG_EXTERNAL).Value & "") = "" Then

        gExternal = True

    End If

    gRecursive = ws.Range(CFG_RECURSIVE).Value

    If Trim$(ws.Range(CFG_RECURSIVE).Value & "") = "" Then

        gRecursive = True

    End If

    gShowJPName = ws.Range(CFG_SHOW_JPNAME).Value
    gShowCategory = ws.Range(CFG_SHOW_CATEGORY).Value

    ' 初回起動対策
    ' Load時に未設定なら既定値。
    If gMaxNode = 0 Then gMaxNode = DEFAULT_MAX_NODE
    If gMaxEdge = 0 Then gMaxEdge = DEFAULT_MAX_EDGE

    If gColorTheme = "" Then

        gColorTheme = DEFAULT_COLOR_THEME

    End If

End Sub

'==================================================
' PathChart既定値
'==================================================
' @JPName
' 既定深度保存
'
' @Category
' PathChart
'
' @Input
' MaxDepth(Long)
'
' @Output
' なし
'
' @Summary
' PathChart作成時の既定探索深度を保存する
'
Public Sub SaveDefaultMaxDepth( _
                    ByVal MaxDepth As Long)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    ws.Range(CFG_MAX_DEPTH).Value = MaxDepth

End Sub

' @JPName
' 既定深度取得
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' MaxDepth(Long)
'
' @Summary
' PathChart作成時の既定探索深度を取得する
'
' @Remarks
' 未設定時は5を返す
'
Public Function GetDefaultMaxDepth() As Long

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    GetDefaultMaxDepth = Val(ws.Range(CFG_MAX_DEPTH).Value)

    If GetDefaultMaxDepth <= 0 Then

        GetDefaultMaxDepth = DEFAULT_MAX_DEPTH

    End If

End Function

' @JPName
' 開始関数保存
'
' @Category
' PathChart
'
' @Input
' StartProc(String)
'
' @Output
' なし
'
' @Summary
' PathChart作成時の開始関数を保存する
'
Public Sub SaveDefaultStartProcedure( _
                    ByVal StartProc As String)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    ws.Range(CFG_START_PROC).Value = StartProc

End Sub

' @JPName
' 開始関数取得
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' ProcedureName(String)
'
' @Summary
' PathChart作成時の開始関数を取得する
'
Public Function GetDefaultStartProcedure() As String

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_PATH_CHART_CONFIG)

    GetDefaultStartProcedure = _
        Trim$(ws.Range(CFG_START_PROC).Value)

End Function
