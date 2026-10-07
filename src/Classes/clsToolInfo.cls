VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToolInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================
' ツール情報
'==================================================
'
' @JPName
' ツール情報
'
' @Category
' Model
'
' @Summary
' ToolManagerで管理する
' ツール定義情報を保持する
'
' @Remarks
' ToolManagerシート1行分の情報を表現する
'

' ツールID
Public ToolID As Long

' ツール名
Public ToolName As String

' 対象Workbook名
Public WorkbookName As String

' GitHubリポジトリ名
Public GitRepository As String

' エクスポート対象フラグ
Public ExportEnabled As Boolean

' 説明
Public Description As String

' 管理ファイル名
Public ManagementFile As String

