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
