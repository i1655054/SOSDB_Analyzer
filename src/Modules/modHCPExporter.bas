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
                ByVal ProcName As String, _
                ByVal HCPNodes As Collection)

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
        ProcName, _
        HCPNodes

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
'--------------------------------------------------
' HTML出力
'--------------------------------------------------
'===================================================================
' @JPName
' HCP HTML出力
'
' @Category
' HTML
'
' @Input
' FilePath(String)
' ModuleName(String)
' ProcName(String)
' HCPNodes(Collection)
'
' @Output
' HCP HTMLファイル
'
' @Summary
' HCPノード一覧を
' HTML形式で出力する
'
' @Remarks
' RenderTextを利用して
' HCPチャートをHTMLへ出力する
'
' CALLノードは
' GetCallNodeHtmlにより
' Procedureリンクへ変換する
'
' Business列が複数行の場合は
' TXT出力と同様に
' 複数<tr>へ展開する
'===================================================================
Private Sub WriteHCPHtmlFile( _
                ByVal FilePath As String, _
                ByVal ModuleName As String, _
                ByVal ProcName As String, _
                ByVal HCPNodes As Collection)

    '# 出力ファイル番号
    Dim FileNo As Integer

    '# HCPノード
    Dim Node As clsHCPNode

    '# HCP表示文字列
    Dim HCPText As String

    '# 行高
    Dim Line_Height As Single

    '# Business分割
    Dim Lines() As String

    '# 行Index
    Dim i As Long

    Line_Height = 1.2

    FileNo = FreeFile

    Open FilePath _
        For Output As #FileNo

    '# HTMLヘッダ
    WriteHtmlHeader _
        FileNo, _
        ProcName, _
        Line_Height, _
        True

    '# 見出し
    Print #FileNo, _
        "<h1>" & _
        ProcName & _
        "</h1>"

    '# モジュール名
    Print #FileNo, _
        "<p>Module : " & _
        ModuleName & _
        "</p>"

    '# Module INDEXへ戻る
    Print #FileNo, _
        "<p>" & _
        "<a href=""./_INDEX.html"">" & _
        "← Module Index" & _
        "</a>" & _
        "</p>"

    '# テーブル開始
    Print #FileNo, _
        "<table>"

    '# 列定義
    Print #FileNo, _
        "<colgroup>" & _
        "<col/>" & _
        "<col/>" & _
        "</colgroup>"

    '# ヘッダ開始
    Print #FileNo, _
        "<thead>"

    Print #FileNo, _
        "<tr>" & _
        "<th class=""resize"">HCP</th>" & _
        "<th class=""resize"">Business</th>" & _
        "</tr>"

    Print #FileNo, _
        "</thead>"

    '# 本体開始
    Print #FileNo, _
        "<tbody>"

    '# HCP開始
    Print #FileNo, _
        "<tr>" & _
        "<td>┬</td>" & _
        "<td></td>" & _
        "</tr>"

    '# STARTノード出力
    For Each Node In HCPNodes

        If Node.NodeType = NODE_START Then

            HCPText = _
                GetNodeHtmlText(Node)

            Print #FileNo, _
                "<tr>" & _
                "<td>" & _
                HCPText & _
                "</td>" & _
                "<td>" & _
                HtmlEncode( _
                    Node.BusinessText) & _
                "</td>" & _
                "</tr>"

            Exit For

        End If

    Next Node

    '# HCP開始終了
    Print #FileNo, _
        "<tr>" & _
        "<td>┴└┐</td>" & _
        "<td></td>" & _
        "</tr>"

    '--------------------------------
    '# HCP本体出力
    '--------------------------------
    For Each Node In HCPNodes

        '# STARTノードは出力済
        If Node.NodeType = NODE_START Then
            GoTo ContinueLoop
        End If

        HCPText = _
            GetNodeHtmlText(Node)

        If Trim$(HCPText) = "" Then
            GoTo ContinueLoop
        End If

        '# Businessなし
        If Trim$(Node.BusinessText) = "" Then

            Print #FileNo, _
                "<tr>" & _
                "<td>" & _
                HCPText & _
                "</td>" & _
                "<td></td>" & _
                "</tr>"

        Else

            '# Business分割
            Lines = Split( _
                Node.BusinessText, _
                vbCrLf)

            '# 1行目
            Print #FileNo, _
                "<tr>" & _
                "<td>" & _
                HCPText & _
                "</td>" & _
                "<td>" & _
                HtmlEncode( _
                    Lines(0)) & _
                "</td>" & _
                "</tr>"

            '# 2行目以降
            For i = 1 To UBound(Lines)

                Print #FileNo, _
                    "<tr>" & _
                    "<td>" & _
                    HtmlEncode( _
                        GetBusinessPrefix( _
                            Node.RenderText)) & _
                    "</td>" & _
                    "<td>" & _
                    HtmlEncode( _
                        Lines(i)) & _
                    "</td>" & _
                    "</tr>"

            Next i

        End If

ContinueLoop:

    Next Node

    '# 本体終了
    Print #FileNo, _
        "</tbody>"

    '# テーブル終了
    Print #FileNo, _
        "</table>"

    '# HTMLフッタ
    WriteHtmlFooter _
        FileNo, _
        True

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

'===================================================================
' @JPName
' ノードHTML文字列取得
'
' @Category
' HCP
'
' @Input
' Node : HCPノード
'
' @Output
' HtmlText
'
' @Summary
' ノード種別に応じた
' HTML文字列を返却する
'
' @Remarks
' HCP記号付与および
' Procedureリンク化を行う
'===================================================================
Private Function GetNodeHtmlText( _
                ByVal Node As clsHCPNode) _
                As String

    '# 描画対象ではないノードを除外
    If Trim$(Node.RenderText) = "" Then

        GetNodeHtmlText = ""

        Exit Function

    End If

    Select Case Node.NodeType

        Case NODE_CALL

            GetNodeHtmlText = _
                GetCallNodeHtml(Node)

        Case Else

            GetNodeHtmlText = _
                GetNormalNodeHtml(Node)

    End Select

End Function

'===================================================================
' @JPName
' Procedureリンク変換
'
' @Category
' HCP
'
' @Input
' HCPText : HCP文字列
'
' @Output
' HtmlText
'
' @Summary
' Procedure名をHCP HTMLリンクへ変換する
'
' @Remarks
' ProcList登録済Procedureのみ
' リンク化する
'===================================================================
Private Function BuildProcedureLink( _
                ByVal HCPText As String) _
                As String

    '# ProcList
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# Procedure名
    Dim ProcName As String

    '# HTMLパス
    Dim HtmlPath As String

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_MODULE).End(xlUp).row

    For r = 2 To LastRow

        ProcName = _
            Trim$( _
                ws.Cells( _
                    r, _
                    COL_PROC).Value)

        If ProcName <> "" Then

            If InStr( _
                    1, _
                    HCPText, _
                    ProcName, _
                    vbTextCompare) > 0 Then

                HtmlPath = _
                    GetProcedureHtmlPath( _
                        ProcName)

                If HtmlPath <> "" Then

                    HCPText = _
                        Replace( _
                            HCPText, _
                            ProcName, _
                            "<a href=""" & _
                            HtmlPath & _
                            """>" & _
                            ProcName & _
                            "</a>")

                End If

            End If

        End If

    Next r

    BuildProcedureLink = _
        HCPText

End Function

'===================================================================
' @JPName
' CALLノードHTML取得
'
' @Category
' HCP
'
' @Input
' Node : HCPノード
'
' @Output
' HtmlText
'
' @Summary
' CALLノード用HTML文字列を返却する
'
' @Remarks
' CallNameに対応する
' Procedureリンクを生成する
'===================================================================
Private Function GetCallNodeHtml( _
                ByVal Node As clsHCPNode) _
                As String

    '# HTMLリンク
    Dim LinkText As String

    '# CallName未設定
    If Trim$(Node.CallName) = "" Then

        GetCallNodeHtml = _
            HtmlEncode( _
                Node.RenderText)

        Exit Function

    End If

    '# Procedureリンク生成
    LinkText = _
        BuildProcedureLink( _
            Node.CallName)

    '# CALLノードHTML生成
    GetCallNodeHtml = _
        Replace( _
            HtmlEncode( _
                Node.RenderText), _
            Node.CallName, _
            LinkText)

End Function

'===================================================================
' @JPName
' 通常ノードHTML取得
'
' @Category
' HCP
'
' @Input
' Node : HCPノード
'
' @Output
' HtmlText
'
' @Summary
' 通常ノードのHTML文字列を返却する
'
' @Remarks
' HCP記号を付与して返却する
'===================================================================
Private Function GetNormalNodeHtml( _
                ByVal Node As clsHCPNode) _
                As String


        GetNormalNodeHtml = _
            HtmlEncode( _
                Node.RenderText)


End Function

'===================================================================
' @JPName
' Business継続行プレフィックス取得
'
' @Category
' HTML
'
' @Input
' RenderText(String)
'
' @Output
' Prefix(String)
'
' @Summary
' Business複数行出力時の
' HCP継続行プレフィックスを取得する
'
' @Remarks
' HTML出力時にBusiness列が
' 複数行となる場合に利用する
'
' 1行目
'   │◎ CallProc      業務説明
'
' 2行目以降
'   ││               内部呼出:XXX
'
' の形式で表示するため
' HCP表示行から継続行用の
' HCPプレフィックスを生成する
'
' RenderText:
'   "　　│◎ CallProc"
'
' 戻り値:
'   "　　││"
'
' TXT出力とHTML出力の
' 表示形式を一致させるために使用する
'===================================================================
Private Function GetBusinessPrefix( _
                ByVal RenderText As String) _
                As String

    Dim Pos As Long

    '# CALL記号検索
    Pos = InStr( _
                RenderText, _
                "◎")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# REFERENCE記号検索
    Pos = InStr( _
                RenderText, _
                "◇")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# NORMAL記号検索
    Pos = InStr( _
                RenderText, _
                "〇")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# DECLARE記号検索
    Pos = InStr( _
                RenderText, _
                "□")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# IF記号検索
    Pos = InStr( _
                RenderText, _
                "◆")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# FOR記号検索
    Pos = InStr( _
                RenderText, _
                "＠")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# 継続行生成不可
    GetBusinessPrefix = "│"

End Function


