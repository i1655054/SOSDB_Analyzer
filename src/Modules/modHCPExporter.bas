Attribute VB_Name = "modHCPExporter"
Option Explicit

' HCPExporterデバッグ出力有効化
#Const DEBUG_MODE = True

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
' モジュール別出力フォルダへ
' .hcp.txt を生成する
'
' 実出力処理は
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
        FilePath, _
        ProcName

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
' モジュール別出力フォルダへ
' .hcp.html を生成する
'
' 実出力処理は
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
                ByVal FilePath As String, _
                ByVal ProcName As String)

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

    '# Calls出力
    WriteCallsText _
        FileNo, _
        ProcName

    '# Called By出力
    WriteCalledByText _
        FileNo, _
        ProcName

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
' HCPNodesを走査し
' RenderTextを利用して
' HCPチャートをHTML化する
'
' CALLノードは
' Procedureリンクへ変換する
'
' Business列が複数行の場合は
' 複数<tr>へ展開する
'
' STARTノードは
' HCP先頭として出力する
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
        "<table class=""notranslate"">"

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
                GetBusinessHtml( _
                    Node, _
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

    '# Calls出力
    WriteCallsHtml _
        FileNo, _
        ProcName

    '# Called By出力
    WriteCalledByHtml _
        FileNo, _
        ProcName

    '# HTMLフッタ
    WriteHtmlFooter _
        FileNo, _
        True

    Close #FileNo

End Sub

'--------------------------------------------------
' HTML補助
'--------------------------------------------------




'===================================================================
' @JPName
' Procedureリンク生成
'
' @Category
' HCP
'
' @Input
' ProcName : Procedure名
'
' @Output
' HtmlText
'
' @Summary
' Procedure名に対応する
' HTMLリンクを生成する
'
' @Remarks
' Procedure名から
' HCP HTML相対パスを取得し
'
' <a href="...">
' </a>
'
' を生成して返却する
'
' Procedure未登録時は
' Procedure名をそのまま返却する
'===================================================================
Public Function BuildProcedureLink( _
                ByVal ProcName As String) _
                As String

    '# HTMLパス
    Dim HtmlPath As String

    '# Procedure名確認
    If Trim$(ProcName) = "" Then

        BuildProcedureLink = ""

        Exit Function

    End If

    '# Procedure HTMLパス取得
    HtmlPath = _
        GetProcedureHtmlPath( _
            ProcName)

    '# Procedure未登録
    If HtmlPath = "" Then

        BuildProcedureLink = _
            ProcName

        Exit Function

    End If

    '# Procedureリンク生成
    BuildProcedureLink = _
        "<a href=""" & _
        HtmlPath & _
        """>" & _
        ProcName & _
        "</a>"

End Function


'===================================================================
' @JPName
' Called By一覧出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出元一覧を出力する
'
' @Remarks
' FunctionDependencyを参照し
' 呼出元Procedure一覧を出力する
'
' 呼出元が存在しない場合は
' 「呼出元なし」を出力する
'===================================================================
Private Sub WriteCalledByText( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    ' 呼出元一覧
    Dim Callers As Collection

    ' 呼出元Procedure
    Dim Caller As Variant

    ' 呼出元Procedure情報
    Dim Info As clsProcInfo
    
    '# 呼出元一覧取得
    Set Callers = _
        GetCalledByList( _
            ProcName)

    Print #FileNo, ""
    Print #FileNo, String$(40, "=")
    Print #FileNo, "Called By"
    Print #FileNo, String$(40, "=")
    Print #FileNo, ""

    '# 呼出元なし
    If Callers.Count = 0 Then

        Print #FileNo, "呼出元なし"

        Exit Sub

    End If

    '# ヘッダ出力
    Print #FileNo, _
        "Procedure" & vbTab & _
        "JPName" & vbTab & _
        "Module"

    Print #FileNo, _
        String$(60, "-")

    '# 呼出元一覧出力
    For Each Caller In Callers

        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Caller))

        If Info Is Nothing Then

            Print #FileNo, _
                CStr(Caller)

        Else

            Print #FileNo, _
                Info.ProcName & vbTab & _
                Info.JPName & vbTab & _
                Info.ModuleName

        End If

    Next Caller

End Sub

'===================================================================
' @JPName
' Called By一覧HTML出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出元一覧をHTML出力する
'
' @Remarks
' FunctionDependencyを参照し
' 呼出元Procedure一覧を出力する
'
' 呼出元Procedureは
' HCPページへのリンクとして出力する
'
' 呼出元が存在しない場合は
' 「呼出元なし」を出力する
'===================================================================
Public Sub WriteCalledByHtml( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    ' 呼出元一覧
    Dim Callers As Collection

    ' 呼出元Procedure
    Dim Caller As Variant

    ' リンクURL
    Dim HtmlPath As String
    
    ' 呼出元Procedure情報
    Dim Info As clsProcInfo

    '# 呼出元一覧取得
    Set Callers = _
        GetCalledByList( _
            ProcName)

    Print #FileNo, "<hr>"

    Print #FileNo, _
        "<h2>Called By (" & _
        Callers.Count & _
        ")</h2>"

    '# 呼出元なし
    If Callers.Count = 0 Then

        Print #FileNo, _
            "<p>呼出元なし</p>"

        Exit Sub

    End If

    '# 呼出元一覧テーブル出力
    Print #FileNo, "<table>"
    
    '# テーブルヘッダ出力
    Print #FileNo, _
        "<tr>" & _
        "<th>Procedure</th>" & _
        "<th>JPName</th>" & _
        "<th>Module</th>" & _
        "</tr>"

    '# 呼出元一覧出力
    For Each Caller In Callers

        '# Procedure情報取得
        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Caller))

        '# HCP HTMLパス取得
        HtmlPath = _
            GetProcedureHtmlPath(CStr(Caller))

        If Info Is Nothing Then

            Print #FileNo, _
                "<tr>" & _
                "<td><a href=""" & _
                HtmlPath & _
                """>" & _
                CStr(Caller) & _
                "</a></td>" & _
                "<td></td>" & _
                "<td></td>" & _
                "</tr>"
      
        Else

            Print #FileNo, _
                "<tr>" & _
                "<td><a href=""" & _
                HtmlPath & _
                """>" & _
                CStr(Caller) & _
                "</a></td>" & _
                "<td>" & _
                HtmlEncode(Info.JPName) & _
                "</td>" & _
                "<td>" & _
                HtmlEncode(Info.ModuleName) & _
                "</td>" & _
                "</tr>"

        End If

    Next Caller

    Print #FileNo, "</table>"

End Sub

'===================================================================
' @JPName
' Calls一覧出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出先一覧を出力する
'
' @Remarks
' FunctionDependencyを参照し
' 呼出先Procedure一覧を出力する
'
' 呼出先が存在しない場合は
' 「呼出先なし」を出力する
'===================================================================
Private Sub WriteCallsText( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    '# 呼出先一覧
    Dim Calls As Collection

    '# 呼出先Procedure
    Dim Proc As Variant

    ' 呼出元Procedure情報
    Dim Info As clsProcInfo
    
    '# 呼出先一覧取得
    Set Calls = _
        GetCallList(ProcName)

    Print #FileNo, ""
    Print #FileNo, String$(40, "=")
    Print #FileNo, _
        "Calls (" & _
        Calls.Count & _
        ")"
    Print #FileNo, String$(40, "=")
    Print #FileNo, ""

    '# 呼出先なし
    If Calls.Count = 0 Then

        Print #FileNo, "呼出先なし"

        Exit Sub

    End If

    '# ヘッダ出力
    Print #FileNo, _
        "Procedure" & vbTab & _
        "JPName" & vbTab & _
        "Module"

    Print #FileNo, _
        String$(60, "-")

    '# 呼出先一覧出力
    For Each Proc In Calls

        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Proc))

        If Info Is Nothing Then

            Print #FileNo, _
                CStr(Proc)

        Else

            Print #FileNo, _
                Info.ProcName & vbTab & _
                Info.JPName & vbTab & _
                Info.ModuleName

        End If

    Next Proc

End Sub

'===================================================================
' @JPName
' Calls一覧HTML出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出先一覧をHTML出力する
'
' @Remarks
' 呼出先Procedureは
' HCPページへのリンクとして出力する
'
' 呼出先が存在しない場合は
' 「呼出先なし」を出力する
'===================================================================
Public Sub WriteCallsHtml( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    '# 呼出先一覧
    Dim Calls As Collection

    '# 呼出先Procedure
    Dim Proc As Variant

    '# HTMLパス
    Dim HtmlPath As String
    
    ' 呼出先Procedure情報
    Dim Info As clsProcInfo

    '# 呼出先一覧取得
    Set Calls = _
        GetCallList(ProcName)

    Print #FileNo, "<hr>"

    Print #FileNo, _
        "<h2>Calls (" & _
        Calls.Count & _
        ")</h2>"

    '# 呼出先なし
    If Calls.Count = 0 Then

        Print #FileNo, _
            "<p>呼出先なし</p>"

        Exit Sub

    End If

    '# 呼出先一覧テーブル出力
    Print #FileNo, "<table>"

    '# テーブルヘッダ出力
    Print #FileNo, _
        "<tr>" & _
        "<th>Procedure</th>" & _
        "<th>JPName</th>" & _
        "<th>Module</th>" & _
        "</tr>"
    
    '# 呼出先一覧出力
    For Each Proc In Calls

        '# Procedure情報取得
        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Proc))

        '# HCP HTMLパス取得
        HtmlPath = _
            GetProcedureHtmlPath(CStr(Proc))

        If Info Is Nothing Then

            Print #FileNo, _
                "<tr>" & _
                "<td><a href=""" & _
                HtmlPath & _
                """>" & _
                CStr(Proc) & _
                "</a></td>" & _
                "<td></td>" & _
                "<td></td>" & _
                "</tr>"

        Else
        
            Print #FileNo, _
                "<tr>" & _
                "<td><a href=""" & _
                HtmlPath & _
                """>" & _
                CStr(Proc) & _
                "</a></td>" & _
                "<td>" & _
                HtmlEncode(Info.JPName) & _
                "</td>" & _
                "<td>" & _
                HtmlEncode(Info.ModuleName) & _
                "</td>" & _
                "</tr>"

        End If

    Next Proc

    Print #FileNo, "</table>"

End Sub


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
Public Function GetNodeHtmlText( _
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

        Case NODE_COMMENT

            GetNodeHtmlText = _
                "<span class=""hcp-comment"">" & _
                HtmlEncode(Node.RenderText) & _
                "</span>"
        
        Case Else

            GetNodeHtmlText = _
                GetNormalNodeHtml(Node)

    End Select

End Function


'===================================================================
' @JPName
' HTMLエンコード
'
' @Category
' HTML
'
' @Input
' Text(String)
'
' @Output
' HtmlText(String)
'
' @Summary
' HTML表示用に文字列を
' エンコードする
'
' @Remarks
' &, <, > をHTMLエンティティへ変換し
' 改行は<br>へ変換する
'===================================================================
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
' CallNameが設定されている場合は
' Procedureリンクへ変換する
'
' CallName未設定時は
' RenderTextをそのまま出力する
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
' Business HTML取得
'
' @Category
' HTML
'
' @Input
' Node(clsHCPNode)
' BusinessText(String)
'
' @Output
' HtmlText(String)
'
' @Summary
' Business列用HTMLを生成する
'
' @Remarks
' CALLノードは
' Business列もProcedureリンク化する
'===================================================================
Private Function GetBusinessHtml( _
                ByVal Node As clsHCPNode, _
                ByVal BusinessText As String) _
                As String

    Dim HtmlPath As String

    If Node.NodeType <> NODE_CALL Then

        GetBusinessHtml = _
            HtmlEncode(BusinessText)

        Exit Function

    End If

    If Trim$(Node.CallName) = "" Then

        GetBusinessHtml = _
            HtmlEncode(BusinessText)

        Exit Function

    End If

    HtmlPath = _
        GetProcedureHtmlPath( _
            Node.CallName)

    If HtmlPath = "" Then

        GetBusinessHtml = _
            HtmlEncode(BusinessText)

        Exit Function

    End If

    GetBusinessHtml = _
        "<a href=""" & _
        HtmlPath & _
        """>" & _
        HtmlEncode(BusinessText) & _
        "</a>"

End Function


