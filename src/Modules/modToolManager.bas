Attribute VB_Name = "modToolManager"
Option Explicit

Public Const BOOK_SOSDB As String = "ÅúSOSDBñ‚êfï[ìoò^.xlsm"
Public Const BOOK_ANALYZER As String = "SOSDBâêÕÉcÅ[Éã.xlsm"
    
' ToolManagerä«óù

Public Type ToolInfo

    ToolID As Long
    ToolName As String
    WorkbookName As String
    GitRepository As String
    ExportEnabled As Boolean
    Description As String

End Type

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
                .ToolName = ws.Cells(r, "B").Value
                .WorkbookName = ws.Cells(r, "C").Value
                .GitRepository = ws.Cells(r, "D").Value
                .ExportEnabled = ws.Cells(r, "E").Value
                .Description = ws.Cells(r, "F").Value

            End With

            Exit Function

        End If

    Next r

End Function


Public Function GetToolList()

End Function

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

Public Function GetRepositoryByToolID( _
                ByVal ToolID As Long) _
                As String

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    GetRepositoryByToolID = _
        Tool.GitRepository

End Function

Public Sub ShowToolList()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = Worksheets("ToolManager")

    LastRow = ws.Cells( _
        ws.Rows.Count, "A") _
        .End(xlUp).row

    For r = 2 To LastRow

        Debug.Print _
            ws.Cells(r, "A").Value, _
            ws.Cells(r, "B").Value

    Next r

End Sub

Public Function GetWorkbookByWorkbookName( _
                    ByVal WorkbookName As String) _
                    As Workbook

    On Error Resume Next

    Set GetWorkbookByWorkbookName = _
        Workbooks(WorkbookName)

    On Error GoTo 0

End Function

