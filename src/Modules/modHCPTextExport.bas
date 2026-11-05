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

    '# HCPチャートシート取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_CHART)

    '# HCPテキスト出力先生成
    FilePath = _
        GetOutputFolder() & _
        ProcName & "." & _
        ModuleName & ".hcp.txt"

    '# 出力ファイルオープン
    FileNo = FreeFile

    Open FilePath For Output As #FileNo

    '# HCPチャート最終行取得
    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    '# HCPテキスト出力
    For r = 1 To LastRow

        HCPText = _
            CStr(ws.Cells(r, 1).Value)

        BusinessText = _
            Trim$(ws.Cells(r, 2).Value)

        '# 空行出力
        If Trim$(HCPText) = "" _
        And BusinessText = "" Then

            Print #FileNo, ""

        '# HCP・業務説明出力
        ElseIf BusinessText <> "" Then

            Print #FileNo, _
                HCPText & vbTab & _
                BusinessText

        '# HCPのみ出力
        Else

            Print #FileNo, _
                HCPText

        End If

    Next r

    '# 出力ファイルクローズ
    Close #FileNo

End Sub
