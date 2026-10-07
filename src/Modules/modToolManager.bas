Attribute VB_Name = "modToolManager"
Option Explicit

'==================================================
' Book名
'==================================================
Public Const BOOK_SOSDB As String = "●SOSDB問診票登録.xlsm"
Public Const BOOK_ANALYZER As String = "SOSDB解析ツール.xlsm"

'==================================================
' シート名
'==================================================

' ToolManagerシート名
Public Const SHEET_TOOL_MANAGER As String = "ToolManager"

' PathChartConfigシート名
Public Const SHEET_PATH_CHART_CONFIG As String = "PathChartConfig"

' ProcListシート名
Public Const SHEET_PROC_LIST As String = "ProcList"

' 関数トレースシート名
Public Const SHEET_FUNCTION_TRACE As String = "関数トレース"

' 関数依存関係シート名
Public Const SHEET_FUNCTION_DEPENDENCY As String = "関数依存関係"

' PathChartシート名
Public Const SHEET_PATH_CHART As String = "PathChart"

' PathChartTreeシート名
Public Const SHEET_PATH_CHART_TREE As String = "PathChartTree"
               
'==================================================
' グローバル状態
'==================================================

' 現在選択中ツール
Public gTargetToolID As Long

' 現在解析コンテキスト
Public gContext As clsAnalyzeContext

' @JPName
' 現在ツール情報取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' ToolInfo(clsToolInfo)
'
' @Summary
' 現在選択中ツールの情報を取得する
'
Public Function GetCurrentToolInfo() As clsToolInfo

    Set GetCurrentToolInfo = _
        GetToolInfo(gTargetToolID)

End Function

' @JPName
' 現在対象ブック取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' Workbook
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
' @Input
' ToolID(Long)
'
' @Output
' FileName(String)
'
' @Summary
' ToolIDに対応する管理ファイル名を取得する
'
Public Function GetManagementFileByToolID( _
                    ByVal ToolID As Long) _
                    As String

    Dim Tool As clsToolInfo

    Set Tool = GetToolInfo(ToolID)

    GetManagementFileByToolID = _
        Tool.ManagementFile

End Function

' @JPName
' 現在管理ファイル取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' FileName(String)
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

' @JPName
' 現在ツールID取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' ToolID(Long)
'
' @Summary
' ToolManagerシートから
' 現在選択中のToolIDを取得する
'
Public Function GetCurrentToolID() As Long

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_TOOL_MANAGER)

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    For r = 2 To LastRow

        If ws.Cells(r, "H").Value = True Then

            GetCurrentToolID = ws.Cells(r, "A").Value

            Exit Function

        End If

    Next r

End Function

' @JPName
' 管理ブック取得
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
' ToolIDに対応する管理ブックを取得する
'
' @Remarks
' 管理ブックが未作成の場合は
' 新規作成して返却する
'
Public Function GetManagementWorkbookByToolID( _
                    ByVal ToolID As Long) _
                    As Workbook

    Dim Tool As clsToolInfo
    Dim FullPath As String

    Set Tool = GetToolInfo(ToolID)

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

' @JPName
' 現在管理ブック取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' Workbook
'
' @Summary
' 現在選択中ツールの管理ブックを取得する
'
Public Function GetCurrentManagementWorkbook() _
                As Workbook

    Set GetCurrentManagementWorkbook = _
        GetManagementWorkbookByToolID( _
            gTargetToolID)

End Function

' @JPName
' 管理ブック確認
'
' @Category
' ToolManager
'
' @Input
' TargetBook(Workbook)
' FileName(String)
'
' @Output
' Boolean
'
' @Summary
' 管理ブックが有効か確認する
'
' @Remarks
' 未オープン時はメッセージを表示する
'
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

' @JPName
' 管理ファイルパス取得
'
' @Category
' ToolManager
'
' @Input
' ToolID(Long)
'
' @Output
' FilePath(String)
'
' @Summary
' ToolIDに対応する
' 管理ファイルのフルパスを取得する
'
Public Function GetManagementFilePathByToolID( _
                    ByVal ToolID As Long) _
                    As String
    
    Dim Tool As clsToolInfo

    Set Tool = GetToolInfo(ToolID)
    
    GetManagementFilePathByToolID = _
        GetOutputFolder() & _
        Tool.ManagementFile

End Function

' @JPName
' 現在管理ファイルパス取得
'
' @Category
' ToolManager
'
' @Input
' なし
'
' @Output
' FilePath(String)
'
' @Summary
' 現在選択中ツールの
' 管理ファイルフルパスを取得する
'
Public Function GetCurrentManagementFilePath() _
                    As String

    GetCurrentManagementFilePath = _
        GetManagementFilePathByToolID( _
            gTargetToolID)

End Function

' @JPName
' 管理ブック作成
'
' @Category
' ToolManager
'
' @Input
' FilePath(String)
' ToolName(String)
'
' @Output
' Workbook
'
' @Summary
' 管理ブックを新規作成する
'
' @Remarks
' ProcList
' 関数トレース
' 関数依存関係
' PathChart
' PathChartTree
' シートを作成して保存する
'
Public Function CreateManagementWorkbook( _
                    ByVal FilePath As String, _
                    ByVal ToolName As String) _
                    As Workbook
    
    On Error GoTo ErrHandler
    
    Dim wb As Workbook

    Set wb = Workbooks.Add

    wb.Worksheets(1).Name = SHEET_PROC_LIST

    wb.Worksheets.Add.Name = SHEET_FUNCTION_TRACE

    wb.Worksheets.Add.Name = SHEET_FUNCTION_DEPENDENCY

    wb.Worksheets.Add.Name = SHEET_PATH_CHART

    wb.Worksheets.Add.Name = SHEET_PATH_CHART_TREE

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
    
    'Debug.Print "SaveOK"

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
                As clsToolInfo

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim Tool As clsToolInfo

    Set ws = ThisWorkbook.Worksheets(SHEET_TOOL_MANAGER)

    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    For r = 2 To LastRow

        If ws.Cells(r, "A").Value = ToolID Then

            Set Tool = New clsToolInfo

            Tool.ToolID = ToolID

            Tool.ToolName = _
                TrimEx(ws.Cells(r, "B").Value)

            Tool.WorkbookName = _
                TrimEx(ws.Cells(r, "C").Value)

            Tool.GitRepository = _
                TrimEx(ws.Cells(r, "D").Value)

            Tool.ExportEnabled = _
                CBool(ws.Cells(r, "E").Value)

            Tool.Description = _
                TrimEx(ws.Cells(r, "F").Value)

            Tool.ManagementFile = _
                TrimEx(ws.Cells(r, "G").Value)

            Set GetToolInfo = Tool

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
' TrimmedText(String)
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

    Dim Tool As clsToolInfo

    Set Tool = GetToolInfo(ToolID)

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

    Dim Tool As clsToolInfo

    Set Tool = GetToolInfo(ToolID)

    GetRepositoryByToolID = _
        Tool.GitRepository

End Function

