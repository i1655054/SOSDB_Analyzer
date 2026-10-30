Attribute VB_Name = "modHCPTextExport"
Option Explicit

'==================================================
' HCPテキスト出力
'==================================================

' @JPName
' HCPテキスト出力
'
' @Category
' HCP
'
' @Input
' ModuleName(String)
' ProcName(String)
'
' @Output
' HCPテキストファイル
'
' @Summary
' HCPChartシート内容を
' テキストファイルへ出力する
'
' @Remarks
' A列(HCP)およびB列(Business)を
' TAB区切りで出力する
'
Public Sub ExportHCPTextFileFromSheet( _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    Dim ws As Worksheet

    Dim FileNo As Integer

    Dim FilePath As String

    Dim LastRow As Long

    Dim r As Long

    Dim HCPText As String

    Dim BusinessText As String

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_CHART)

    FilePath = _
        GetOutputFolder() & _
        ProcName & "." & _
        ModuleName & ".hcp.txt"

    FileNo = FreeFile

    Open FilePath For Output As #FileNo

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    For r = 1 To LastRow

        HCPText = _
            CStr(ws.Cells(r, 1).Value)

        BusinessText = _
            Trim$(ws.Cells(r, 2).Value)

        If Trim$(HCPText) = "" _
        And BusinessText = "" Then

            Print #FileNo, ""

        ElseIf BusinessText <> "" Then

            Print #FileNo, _
                HCPText & vbTab & _
                BusinessText

        Else

            Print #FileNo, _
                HCPText

        End If

    Next r

    Close #FileNo

End Sub
