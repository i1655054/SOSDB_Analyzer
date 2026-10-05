Attribute VB_Name = "modToolManager"
Option Explicit

Public Const BOOK_SOSDB As String = "●SOSDB問診票登録.xlsm"
Public Const BOOK_ANALYZER As String = "SOSDB解析ツール.xlsm"

Public gTargetToolID As Long

' ToolManager管理

Public Type ToolInfo

    ToolID As Long
    ToolName As String
    WorkbookName As String
    GitRepository As String
    ExportEnabled As Boolean
    Description As String
    ManagementFile As String

End Type

' @JPName
' 現在ツール情報取得
'
' @Category
' ToolManager
'
' @Summary
' 現在選択中ツールの情報を取得する
'
Public Function GetCurrentToolInfo() As ToolInfo

    GetCurrentToolInfo = _
        GetToolInfo(gTargetToolID)

End Function

' @JPName
' 現在対象ブック取得
'
' @Category
' ToolManager
'
' @Summary
' 現在選択中ツールのWorkbookを取得する
'
Public Function GetCurrentTargetBook() _
        As Workbook

    Set GetCurrentTargetBook = _
        GetWorkbookByToolID(gTargetToolID)

End Function

' @JPName
' 管理ファイル取得
'
' @Category
' ToolManager
'
' @Summary
' ToolIDに対応する管理ファイル名を取得する
'
Public Function GetManagementFileByToolID( _
                    ByVal ToolID As Long) _
                    As String

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    GetManagementFileByToolID = _
        Tool.ManagementFile

End Function

' @JPName
' 現在管理ファイル取得
'
' @Category
' ToolManager
'
' @Summary
' 現在選択中ツールの管理ファイル名を取得する
'
Public Function GetCurrentManagementFile() _
                    As String

    GetCurrentManagementFile = _
        GetManagementFileByToolID( _
            GetCurrentToolInfo().ToolID)

End Function

Public Function GetCurrentToolID() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets("ToolManager")

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 2 To LastRow

        If ws.Cells(r, "H").Value = True Then

            GetCurrentToolID = ws.Cells(r, "A").Value

            Exit Function

        End If

    Next r

End Function

' 管理ブック取得
Public Function GetManagementWorkbookByToolID( _
                    ByVal ToolID As Long) _
                    As Workbook

    Dim Tool As ToolInfo
    Dim FullPath As String

    Tool = GetToolInfo(ToolID)

    On Error Resume Next

    Set GetManagementWorkbookByToolID = _
        Workbooks(Tool.ManagementFile)

    On Error GoTo 0

    If GetManagementWorkbookByToolID Is Nothing Then

        FullPath = _
            GetManagementFilePathByToolID(ToolID)

        If Dir(FullPath) = "" Then

            Set GetManagementWorkbookByToolID = _
                CreateManagementWorkbook(FullPath, Tool.ToolName)
            
            ThisWorkbook.Activate
            
            Exit Function

        Else
        
            Set GetManagementWorkbookByToolID = _
                Workbooks.Open(FullPath)
        
            ThisWorkbook.Activate
            
        End If
        
    End If

End Function

' 現在管理ブック取得
Public Function GetCurrentManagementWorkbook() _
                As Workbook

    Set GetCurrentManagementWorkbook = _
        GetManagementWorkbookByToolID( _
            gTargetToolID)

End Function

' 管理ブック存在確認
Public Function CheckManagementWorkbook( _
                ByVal TargetBook As Workbook, _
                ByVal FileName As String) _
                As Boolean

    If TargetBook Is Nothing Then

        MsgBox _
            "管理ファイルが開かれていません。" _
            & vbCrLf & _
            FileName, _
            vbExclamation

        Exit Function

    End If

    CheckManagementWorkbook = True

End Function

' 管理ファイルフルパス取得
Public Function GetManagementFilePathByToolID( _
                    ByVal ToolID As Long) _
                    As String

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    GetManagementFilePathByToolID = _
        GetOutputFolder() & _
        Tool.ManagementFile

End Function

' 現在管理ファイルフルパス取得
Public Function GetCurrentManagementFilePath() _
                    As String

    GetCurrentManagementFilePath = _
        GetManagementFilePathByToolID( _
            gTargetToolID)

End Function

Public Function CreateManagementWorkbook( _
                    ByVal FilePath As String, _
                    ByVal ToolName As String) _
                    As Workbook
    
    On Error GoTo ErrHandler
    
    Dim wb As Workbook

    Set wb = Workbooks.Add

    wb.Worksheets(1).Name = "ProcList"

    wb.Worksheets.Add.Name = "関数トレース"

    wb.Worksheets.Add.Name = "関数依存関係"

    wb.Worksheets.Add.Name = "PathChart"

    wb.Worksheets.Add.Name = "PathChartTree"

    Application.DisplayAlerts = False

    wb.SaveAs FilePath

    Application.DisplayAlerts = True

    MsgBox _
        "管理ファイルを作成しました。" _
        & vbCrLf & vbCrLf _
        & "対象ツール : " & ToolName _
        & vbCrLf _
        & "管理ファイル : " & FilePath, _
        vbInformation

    Set CreateManagementWorkbook = wb
    
    Debug.Print "SaveOK"

    Exit Function
    
ErrHandler:

    MsgBox _
        "管理ファイル作成失敗" & vbCrLf & _
        Err.Number & " : " & _
    Err.Description

End Function

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
                .ManagementFile = TrimEx(ws.Cells(r, "G").Value)

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

