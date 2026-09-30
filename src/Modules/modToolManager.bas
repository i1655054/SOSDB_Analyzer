Attribute VB_Name = "modToolManager"
Option Explicit

Public Const BOOK_SOSDB As String = "●SOSDB問診票登録.xlsm"
Public Const BOOK_ANALYZER As String = "SOSDB解析ツール.xlsm"
    
' ToolManager管理

Public Type ToolInfo

    ToolID As Long
    ToolName As String
    WorkbookName As String
    GitRepository As String
    ExportEnabled As Boolean
    Description As String

End Type

'==================================================
' Tool情報取得
'==================================================

' @JPName
' ツール情報取得
'
' @Category
' ToolManager
'
' @Input
' ToolID(Long)
'
' @Output
' ToolInfo
'
' @Summary
' ToolManagerシートから
' 指定ツールの情報を取得する
'
' @Remarks
' ToolIDをキーとして検索する
'
Public Function GetToolInfo( _
                ByVal ToolID As Long) _
                As ToolInfo

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("ToolManager")

    LastRow = ws.Cells( _
                ws.Rows.Count, "A") _
                .End(xlUp).row

    For r = 2 To LastRow

        If ws.Cells(r, "A").Value = ToolID Then

            With GetToolInfo

                .ToolID = ToolID

                .ToolName = TrimEx(ws.Cells(r, "B").Value)
                .WorkbookName = TrimEx(ws.Cells(r, "C").Value)
                .GitRepository = TrimEx(ws.Cells(r, "D").Value)
                .ExportEnabled = ws.Cells(r, "E").Value
                .Description = TrimEx(ws.Cells(r, "F").Value)

            End With

            Exit Function

        End If

    Next r

End Function

' @JPName
' 文字列整形
'
' @Category
' ToolManager
'
' @Input
' Value(Variant)
'
' @Output
' String
'
' @Summary
' 前後空白および全角空白を除去する
'
Private Function TrimEx(ByVal Value As Variant) As String

    TrimEx = Trim$(Replace(CStr(Value), "　", ""))

End Function

'==================================================
' Workbook取得
'==================================================

' @JPName
' 対象ブック取得
'
' @Category
' ToolManager
'
' @Input
' ToolID(Long)
'
' @Output
' Workbook
'
' @Summary
' ToolIDに対応する
' Workbookを取得する
'
' @Remarks
' Workbook未オープン時はNothingを返す
'
Public Function GetWorkbookByToolID( _
                ByVal ToolID As Long) _
                As Workbook

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    On Error Resume Next

    Set GetWorkbookByToolID = _
        Workbooks(Tool.WorkbookName)

    On Error GoTo 0

End Function

' @JPName
' Workbook取得
'
' @Category
' ToolManager
'
' @Input
' WorkbookName(String)
'
' @Output
' Workbook
'
' @Summary
' Workbook名から
' Workbookオブジェクトを取得する
'
' @Remarks
' Workbook未オープン時はNothingを返す
'
Public Function GetWorkbookByWorkbookName( _
                    ByVal WorkbookName As String) _
                    As Workbook

    On Error Resume Next

    Set GetWorkbookByWorkbookName = _
        Workbooks(WorkbookName)

    On Error GoTo 0

End Function

'==================================================
' Repository取得
'==================================================

' @JPName
' リポジトリ名取得
'
' @Category
' ToolManager
'
' @Input
' ToolID(Long)
'
' @Output
' RepositoryName(String)
'
' @Summary
' ToolIDに対応する
' GitHubリポジトリ名を取得する
'
Public Function GetRepositoryByToolID( _
                ByVal ToolID As Long) _
                As String

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    GetRepositoryByToolID = _
        Tool.GitRepository

End Function

