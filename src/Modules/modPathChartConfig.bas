Attribute VB_Name = "modPathChartConfig"
Option Explicit

'解析対象
Public gProperty As Boolean
Public gPrivate As Boolean
Public gPublic As Boolean

Public gStdModule As Boolean
Public gUserForm As Boolean
Public gClassModule As Boolean

'除外条件
Public gIgnoreAPI As Boolean
Public gIgnoreExcel As Boolean
Public gIgnoreSelf As Boolean
Public gSameModuleOnly As Boolean

'出力設定
Public gVertical As Boolean
Public gColorTheme As String

'制限
Public gMaxNode As Long
Public gMaxEdge As Long

'共通関数除外
Public gSkipCommon As Boolean
'呼出回数表示
Public gShowCount As Boolean

'PowerPoint出力
Public gPowerPoint As Boolean
'PNG出力
Public gPngOutput As Boolean

'モジュール表示
Public gShowModule As Boolean
'モジュール色分け
Public gColorModule As Boolean

Private Const CFG_SKIP_COMMON      As String = "B1"
Private Const CFG_SHOW_COUNT       As String = "B2"

Private Const CFG_POWERPOINT       As String = "B3"
Private Const CFG_PNG_OUTPUT       As String = "B4"

Private Const CFG_SHOW_MODULE      As String = "B5"
Private Const CFG_COLOR_MODULE     As String = "B6"

Private Const CFG_PROPERTY         As String = "B7"
Private Const CFG_PRIVATE          As String = "B8"
Private Const CFG_PUBLIC           As String = "B9"

Private Const CFG_STD_MODULE       As String = "B10"
Private Const CFG_USER_FORM        As String = "B11"
Private Const CFG_CLASS_MODULE     As String = "B12"

Private Const CFG_IGNORE_API       As String = "B13"
Private Const CFG_IGNORE_EXCEL     As String = "B14"
Private Const CFG_IGNORE_SELF      As String = "B15"
Private Const CFG_SAME_MODULE_ONLY As String = "B16"

Private Const CFG_VERTICAL         As String = "B17"
Private Const CFG_COLOR_THEME      As String = "B18"

Private Const CFG_MAX_NODE         As String = "B19"
Private Const CFG_MAX_EDGE         As String = "B20"
Private Const CFG_MAX_DEPTH        As String = "B21"
Private Const CFG_START_PROC       As String = "B22"

Private Const DEFAULT_MAX_DEPTH As Long = 5

' 設定保存機能
Public Sub SavePathChartConfig()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets( _
            "PathChartConfig")

    ws.Range(CFG_SKIP_COMMON).Value = gSkipCommon
    ws.Range(CFG_SHOW_COUNT).Value = gShowCount

    ws.Range(CFG_POWERPOINT).Value = gPowerPoint
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

    ws.Range(CFG_VERTICAL).Value = gVertical
    ws.Range(CFG_COLOR_THEME).Value = gColorTheme

    ws.Range(CFG_MAX_NODE).Value = gMaxNode
    ws.Range(CFG_MAX_EDGE).Value = gMaxEdge

End Sub

' 設定読込機能
Public Sub LoadPathChartConfig()

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets( _
            "PathChartConfig")

    gSkipCommon = ws.Range(CFG_SKIP_COMMON).Value
    gShowCount = ws.Range(CFG_SHOW_COUNT).Value

    gPowerPoint = ws.Range(CFG_POWERPOINT).Value
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

    gVertical = ws.Range(CFG_VERTICAL).Value
    gColorTheme = ws.Range(CFG_COLOR_THEME).Value

    gMaxNode = Val(ws.Range(CFG_MAX_NODE).Value)
    gMaxEdge = Val(ws.Range(CFG_MAX_EDGE).Value)

    ' 初回起動対策
    ' Load時に未設定なら既定値。
    If gMaxNode = 0 Then gMaxNode = 1000
    If gMaxEdge = 0 Then gMaxEdge = 3000

    If gColorTheme = "" Then

        gColorTheme = "標準"

    End If

End Sub

' 最大深度保存関数
Public Sub SaveDefaultMaxDepth( _
                    ByVal MaxDepth As Long)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("PathChartConfig")

    ws.Range(CFG_MAX_DEPTH).Value = _
        MaxDepth

End Sub

' 最大深度取得関数
Public Function GetDefaultMaxDepth() As Long

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("PathChartConfig")

    GetDefaultMaxDepth = _
        Val(ws.Range(CFG_MAX_DEPTH).Value)

    If GetDefaultMaxDepth <= 0 Then

        GetDefaultMaxDepth = 5

    End If

End Function

' 開始Procedure保存関数
Public Sub SaveDefaultStartProcedure( _
                    ByVal StartProc As String)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("PathChartConfig")

    ws.Range(CFG_START_PROC).Value = _
        StartProc

End Sub

' 開始Procedures復元関数
Public Function GetDefaultStartProcedure() As String

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("PathChartConfig")

    GetDefaultStartProcedure = _
        Trim$(ws.Range(CFG_START_PROC).Value)

End Function


