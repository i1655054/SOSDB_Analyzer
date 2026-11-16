Attribute VB_Name = "modHCPExporter"
Option Explicit

'===================================================================
'@JPName     hcp.txt & hcp.HTML 出力
'===================================================================

'--------------------------------------------------
' Export API
'--------------------------------------------------

'===================================================================
' @JPName
' HCPテキスト出力
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
' ProcName   : Procedure名
'
' @Output
' なし
'
' @Summary
' HCPテキストファイルを出力する
'
' @Remarks
' 出力処理は
' WriteHCPTextFileへ委譲する
'===================================================================
Public Sub ExportHCPTextFile( _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    '# 出力ファイルパス
    Dim FilePath As String

    Trace "HCP", _
        "ExportHCPTextFile " & _
        ModuleName & "." & _
        ProcName

    FilePath = _
        GetHCPOutputFolder( _
            ModuleName) & _
        ProcName & "." & _
        ModuleName & _
        ".hcp.txt"

    WriteHCPTextFile _
        FilePath

End Sub

'===================================================================
' @JPName
' HCPHTML出力
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
' ProcName   : Procedure名
'
' @Output
' なし
'
' @Summary
' HCPHTMLファイルを出力する
'
' @Remarks
' 出力処理は
' WriteHCPHtmlFileへ委譲する
'===================================================================
Public Sub ExportHCPHtmlFile( _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    '# 出力ファイルパス
    Dim FilePath As String

    Trace "HCP", _
        "ExportHCPHtmlFile " & _
        ModuleName & "." & _
        ProcName

    FilePath = _
        GetHCPOutputFolder( _
            ModuleName) & _
        ProcName & "." & _
        ModuleName & _
        ".hcp.html"

    WriteHCPHtmlFile _
        FilePath, _
        ModuleName, _
        ProcName

End Sub

'--------------------------------------------------
' Text出力
'--------------------------------------------------

'===================================================================
' @JPName
' HCPテキスト書込
'
' @Category
' HCP
'
' @Input
' FilePath : 出力ファイル
'
' @Output
' HCPテキストファイル
'
' @Summary
' HCPChart内容を
' テキスト出力する
'
' @Remarks
' HCP列およびBusiness列を
' TAB区切りで出力する
'===================================================================
Private Sub WriteHCPTextFile( _
                ByVal FilePath As String)

    '# HCPチャートシート
    Dim ws As Worksheet

    '# 出力ファイル番号
    Dim FileNo As Integer

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# HCP文字列
    Dim HCPText As String

    '# 業務説明
    Dim BusinessText As String

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_CHART)

    FileNo = FreeFile

    Open FilePath _
        For Output As #FileNo

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
                HCPText & _
                vbTab & _
                BusinessText

        Else

            Print #FileNo, _
                HCPText

        End If

    Next r

    Close #FileNo

End Sub

'==================================================
' HCPテキスト & HTML 出力
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
'削除対象
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
        GetHCPOutputFolder( _
            ModuleName) & _
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

'--------------------------------------------------
' HTML出力
'--------------------------------------------------

'===================================================================
' @JPName
' HCPHTML書込
'
' @Category
' HCP
'
' @Input
' FilePath   : 出力先
' ModuleName : モジュール名
' ProcName   : Procedure名
'
' @Output
' HCPHTMLファイル
'
' @Summary
' HCPをHTML形式で出力する
'
' @Remarks
' ブラウザ閲覧用
'===================================================================
Private Sub WriteHCPHtmlFile( _
                ByVal FilePath As String, _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    '# 出力ファイル番号
    Dim FileNo As Integer
    
    Dim ws As Worksheet
    
    Dim LastRow As Long
    
    Dim r As Long
    
    '# 行高
    Dim Line_Height As Single
    
    Line_Height = 1.2
    
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_CHART)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row
    
    FileNo = FreeFile

    Open FilePath _
        For Output As #FileNo

    'WriteHtmlHeader _
    '    FileNo, _
    '    ProcName

    WriteHtmlHeader _
        FileNo, _
        ProcName, _
        Line_Height

    Print #FileNo, _
        "<h1>" & _
        ProcName & _
        "</h1>"

    Print #FileNo, _
        "<p>Module : " & _
        ModuleName & _
        "</p>"

    Print #FileNo, _
        "<p><a href=""./_INDEX.html"">← Module Index</a></p>"
    
    Print #FileNo, "<table>"

    'Print #FileNo, _
    '    "<tr>" & _
    '    "<th>HCP</th>" & _
    '    "<th>Business</th>" & _
    '    "</tr>"
    
    ' --- 修正 追加 ここから ---
    Print #FileNo, "<colgroup><col/><col/></colgroup>"

    Print #FileNo, _
        "<thead>" & _
        "<tr>" & _
        "<th class=""resize"">HCP</th>" & _
        "<th class=""resize"">Business</th>" & _
        "</tr>" & _
        "</thead>"

    Print #FileNo, "<tbody>"
    
    ' --- 修正 追加 ここまで ---
    
    '# HCP本体出力
    
    For r = 1 To LastRow

        Print #FileNo, _
            "<tr><td>" & _
            HtmlEncode(ws.Cells(r, 1).Text) & _
            "</td><td>" & _
            HtmlEncode(ws.Cells(r, 2).Text) & _
            "</td></tr>"

    Next r

    ' --- 修正 追加 ここから ---
    Print #FileNo, "</tbody>"
    ' --- 修正 追加 ここまで ---
    
    Print #FileNo, "</table>"

    WriteHtmlFooter _
        FileNo

    Close #FileNo

End Sub

'--------------------------------------------------
' HTML補助
'--------------------------------------------------


Private Function HtmlEncode( _
                ByVal Text As String) _
                As String

    HtmlEncode = Text

    HtmlEncode = _
        Replace(HtmlEncode, "&", "&amp;")

    HtmlEncode = _
        Replace(HtmlEncode, "<", "&lt;")

    HtmlEncode = _
        Replace(HtmlEncode, ">", "&gt;")

    HtmlEncode = _
        Replace(HtmlEncode, vbCrLf, "<br>")

End Function

