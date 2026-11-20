Attribute VB_Name = "modHCPIndex"
Option Explicit

'===================================================================
' @JPName
' HCP INDEX生成
'
' @Category
' HCP
'
' @Summary
' モジュール単位INDEXおよび
' ツール単位INDEXの生成機能を提供する
'
' @Remarks
' HCP出力フォルダを走査し
' _INDEX.html を生成する
'
' Module INDEX
' Tool INDEX
'
' の両方を管理する
'===================================================================
'==================================================
' HCP_TARGETS列定義
'==================================================
Public Const COL_HCP_EXEC As Long = 1
Public Const COL_HCP_MODULE As Long = 2
Public Const COL_HCP_PROC As Long = 3
Public Const COL_HCP_STATUS As Long = 4
Public Const COL_HCP_UPDATE_DATE As Long = 5

'==================================================
' Status表示
'==================================================
Public Const STATUS_OK = "[OK]"
Public Const STATUS_ERROR = "[ERROR]"
Public Const STATUS_PENDING = "[PENDING]"
Public Const STATUS_SKIP = "[SKIP]"
Public Const STATUS_UNKNOWN = "[UNKNOWN]"

'--------------------------------------------------
' Index生成
'--------------------------------------------------

'===================================================================
' @JPName
' モジュールINDEX再生成
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
'
' @Output
' なし
'
' @Summary
' 指定モジュール配下のHCPファイルから
' _INDEX.html を再生成する
'
' @Remarks
' 指定モジュールフォルダ配下の
' HCPファイル一覧を取得し
' _INDEX.html を再生成する
'
' 既存INDEXは上書き作成する
'
'===================================================================
Public Sub RebuildModuleIndex( _
                ByVal ModuleName As String)

    '# Procedure一覧
    Dim Procedures As Collection

    '# INDEXファイル
    Dim IndexPath As String

    Trace "INDEX", _
        "RebuildModuleIndex " & _
        ModuleName

    Set Procedures = _
        GetModuleProcedureList( _
            ModuleName)

    IndexPath = _
        GetHCPOutputFolder( _
            ModuleName) & _
        "_INDEX.html"

    WriteHtmlIndexFile _
        IndexPath, _
        ModuleName, _
        Procedures

    Trace "INDEX", _
        "ModuleIndex Created " & _
        ModuleName & _
        " Procedures=" & _
        Procedures.Count

End Sub

'===================================================================
' @JPName
' ツールINDEX再生成
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ツール配下のモジュール一覧から
' _INDEX.html を再生成する
'
' @Remarks
' HCPルート配下の
' モジュールフォルダを走査し
'
' Procedure数
' 生成済数
' 未生成数
'
' を集計して
' Tool INDEXを再生成する
'
'===================================================================
Public Sub RebuildToolIndex()

    '# HCPルート
    Dim RootPath As String

    '# モジュールフォルダ名
    Dim FolderName As String

    '# モジュール数
    Dim ModuleCount As Long

    '# INDEXファイル
    Dim IndexPath As String

    Trace "INDEX", _
        "RebuildToolIndex"

    RootPath = _
        GetHCPToolFolder()

    '-----------------------------
    '# モジュール数取得
    '-----------------------------
    FolderName = _
        Dir( _
            RootPath & "*", _
            vbDirectory)

    Do While FolderName <> ""

        If FolderName <> "." _
        And FolderName <> ".." Then

            If (GetAttr( _
                    RootPath & FolderName) _
                    And vbDirectory) <> 0 Then

                ModuleCount = _
                    ModuleCount + 1

            End If

        End If

        FolderName = Dir()

    Loop

    '-----------------------------
    '# INDEX HTML 出力
    '-----------------------------
    IndexPath = _
        RootPath & _
        "_INDEX.html"

    WriteToolIndexHtml _
        IndexPath, _
        ModuleCount
    
    Trace "INDEX", _
        "ToolIndex Created Modules=" & _
        ModuleCount

End Sub

'===================================================================
' @JPName
' INDEX一括再生成
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 全モジュールINDEXおよび
' ツールINDEXを再生成する
'
' @Remarks
' HCP_TARGETSシートから
' モジュール一覧を取得する
'
' 各Module INDEX生成後に
' Tool INDEXを再生成する
'
'===================================================================
Public Sub RebuildAllIndexes()

    '# HCP_TARGETS
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# モジュール名
    Dim ModuleName As String

    '# モジュール管理
    Dim Modules As Object

    '# Dictionaryキー
    Dim Key As Variant

    Trace "INDEX", _
        "RebuildAllIndexes"

    Set Modules = _
        CreateObject( _
            "Scripting.Dictionary")

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_HCP_MODULE).End(xlUp).row

    '# モジュール一覧作成
    For r = 2 To LastRow

        ModuleName = _
            Trim$( _
                ws.Cells( _
                    r, _
                    COL_HCP_MODULE).Value)

        If ModuleName <> "" Then

            If Not Modules.Exists( _
                    ModuleName) Then

                Modules.Add _
                    ModuleName, _
                    True

            End If

        End If

    Next r

    '# モジュールINDEX再生成
    For Each Key In Modules.Keys

        RebuildModuleIndex _
            CStr(Key)

    Next Key

    '# ツールINDEX再生成
    RebuildToolIndex

End Sub

'--------------------------------------------------
' Index確認
'--------------------------------------------------

'===================================================================
' @JPName
' モジュールINDEX存在確認
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
'
' @Output
' Boolean
'
' @Summary
' 指定モジュールの_INDEX.html
' 存在有無を返却する
'
' @Remarks
' modHCPIndex用
'
' 現在未使用
' 将来的なINDEX整合性確認機能で
' 利用予定
'
'===================================================================
' 未使用
' 将来のINDEX整合性確認機能で再利用予定
Public Function ExistsModuleIndex( _
                ByVal ModuleName As String) _
                As Boolean

    '# INDEXファイル
    Dim IndexPath As String

    IndexPath = _
        GetHCPOutputFolder( _
            ModuleName) & _
        "_INDEX.html"

    ExistsModuleIndex = _
        (Dir(IndexPath) <> "")

End Function


'===================================================================
' @JPName
' ツールINDEX存在確認
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' Boolean
'
' @Summary
' ツールINDEXの存在有無を返却する
'
' @Remarks
' 現在未使用
'
' Tool INDEX整合性確認機能で
' 利用予定
'
'===================================================================
' 未使用
' 将来のINDEX整合性確認機能で再利用予定
Public Function ExistsToolIndex() _
                As Boolean

    '# INDEXファイル
    Dim IndexPath As String

    IndexPath = _
        GetHCPToolFolder() & _
        "_INDEX.html"

    ExistsToolIndex = _
        (Dir(IndexPath) <> "")

End Function

'--------------------------------------------------
' Index HTML生成
'--------------------------------------------------

'===================================================================
' @JPName
' Module INDEX HTML出力
'
' @Category
' HCP
'
' @Input
' IndexPath     : INDEX出力先
' HeaderText    : INDEX見出し
' ProcedureList : Procedure一覧
'
' @Output
' なし
'
' @Summary
' Module INDEXをHTML形式で出力する
'
' @Remarks
' モジュール単位INDEX生成用
'===================================================================
Public Sub WriteHtmlIndexFile( _
                ByVal IndexPath As String, _
                ByVal HeaderText As String, _
                ByVal ProcedureList As Collection)

    '# 出力ファイル番号
    Dim FileNo As Integer

    '# Procedure情報
    Dim Info As clsProcedureInfo

    FileNo = FreeFile

    Open IndexPath _
        For Output As #FileNo

    '# HTMLヘッダ
    WriteHtmlHeader _
        FileNo, _
        HeaderText

    '# Tool Indexへ戻る
    Print #FileNo, _
        "<p>" & _
        "<a href=""" & _
        GetIndexRelativePath(1) & _
        """>← Tool Index</a>" & _
        "</p>"
    
    '# 見出し
    Print #FileNo, _
        "<h1>" & _
        HeaderText & _
        "</h1>"

    '# 件数
    Print #FileNo, _
        "<p>Procedure Count : " & _
        ProcedureList.Count & _
        "</p>"

    '# テーブル開始
    Print #FileNo, _
        "<table>"

    '# ヘッダ行
    Print #FileNo, _
        "<tr>" & _
        "<th>Procedure</th>" & _
        "<th>JPName</th>" & _
        "<th>Status</th>" & _
        "<th>Updated</th>" & _
        "</tr>"

    '# 明細行
    For Each Info In ProcedureList

        Print #FileNo, _
            "<tr>" & _
            "<td><a href=""" & _
            Info.FileName & _
            ".hcp.html"">" & _
            Info.ProcedureName & _
            "</a></td>" & _
            "<td>" & _
            Info.JPName & _
            "</td>" & _
            "<td>" & _
            Info.Status & _
            "</td>" & _
            "<td>" & _
            Info.LastUpdate & _
            "</td>" & _
            "</tr>"

    Next Info

    '# テーブル終了
    Print #FileNo, _
        "</table>"

    '# HTMLフッタ
    WriteHtmlFooter _
        FileNo

    Close #FileNo

End Sub

'===================================================================
' @JPName
' Tool INDEX HTML出力
'
' @Category
' HCP
'
' @Input
' IndexPath   : 出力先
' ModuleCount : モジュール数
'
' @Output
' なし
'
' @Summary
' Tool単位の_INDEX.htmlを生成する
'
' @Remarks
' モジュール毎のINDEXへの
' リンク一覧を出力する
'===================================================================
Private Sub WriteToolIndexHtml( _
                ByVal IndexPath As String, _
                ByVal ModuleCount As Long)

    '# 出力ファイル番号
    Dim FileNo As Integer

    '# HCPルート
    Dim RootPath As String

    '# モジュールフォルダ名
    Dim FolderName As String

    '# Procedure一覧
    Dim Procedures As Collection

    '# Procedure情報
    Dim Info As clsProcedureInfo

    '# Procedure総数
    Dim ProcedureCount As Long

    '# HCP生成済数
    Dim GeneratedCount As Long

    '# 未生成数
    Dim PendingCount As Long

    '# モジュール生成済数
    Dim ModuleGenerated As Long

    FileNo = FreeFile

    Open IndexPath _
        For Output As #FileNo

    WriteHtmlHeader _
        FileNo, _
        GetCurrentToolInfo.ToolName

    Print #FileNo, _
        "<h1>" & _
        GetCurrentToolInfo.ToolName & _
        "</h1>"

    RootPath = _
        GetHCPToolFolder()

    '--------------------------------
    '# 全体集計
    '--------------------------------
    FolderName = _
        Dir( _
            RootPath & "*", _
            vbDirectory)

    Do While FolderName <> ""

        If FolderName <> "." _
        And FolderName <> ".." Then

            If (GetAttr( _
                    RootPath & FolderName) _
                    And vbDirectory) <> 0 Then

                Set Procedures = _
                    GetModuleProcedureList( _
                        FolderName)

                ProcedureCount = _
                    ProcedureCount + _
                    Procedures.Count

                For Each Info In Procedures

                    If Info.Status = _
                        STATUS_OK Then

                        GeneratedCount = _
                            GeneratedCount + 1

                    Else

                        PendingCount = _
                            PendingCount + 1

                    End If

                Next Info

            End If

        End If

        FolderName = Dir()

    Loop

    Print #FileNo, _
        "<p>Module Count : " & _
        ModuleCount & _
        "</p>"

    Print #FileNo, _
        "<p>Procedure Count : " & _
        ProcedureCount & _
        "</p>"

    Print #FileNo, _
        "<p>Generated Count : " & _
        GeneratedCount & _
        "</p>"

    Print #FileNo, _
        "<p>Pending Count : " & _
        PendingCount & _
        "</p>"

    Print #FileNo, _
        "<table>"

    Print #FileNo, _
        "<tr>" & _
        "<th>Module</th>" & _
        "<th>Procedures</th>" & _
        "<th>Generated</th>" & _
        "</tr>"

    '--------------------------------
    '# モジュール一覧
    '--------------------------------
    FolderName = _
        Dir( _
            RootPath & "*", _
            vbDirectory)

    Do While FolderName <> ""

        If FolderName <> "." _
        And FolderName <> ".." Then

            If (GetAttr( _
                    RootPath & FolderName) _
                    And vbDirectory) <> 0 Then

                Set Procedures = _
                    GetModuleProcedureList( _
                        FolderName)

                ModuleGenerated = 0

                For Each Info In Procedures

                    If Info.Status = _
                        STATUS_OK Then

                        ModuleGenerated = _
                            ModuleGenerated + 1

                    End If

                Next Info

                Print #FileNo, _
                    "<tr>" & _
                    "<td>" & _
                    "<a href=""" & _
                    FolderName & _
                    "/_INDEX.html"">" & _
                    FolderName & _
                    "</a>" & _
                    "</td>" & _
                    "<td align=""right"">" & _
                    Procedures.Count & _
                    "</td>" & _
                    "<td align=""right"">" & _
                    ModuleGenerated & _
                    "</td>" & _
                    "</tr>"

            End If

        End If

        FolderName = Dir()

    Loop

    Print #FileNo, _
        "</table>"

    WriteHtmlFooter _
        FileNo

    Close #FileNo

End Sub

'===================================================================
' @JPName
' HTMLヘッダ出力
'
' @Category
' HCP
'
' @Input
' FileNo         : 出力ファイル番号
' Title          : HTMLタイトル
' Line_Height    : 行間
' EnableResize   : 列幅変更有無
'
' @Output
' なし
'
' @Summary
' HTML共通ヘッダを出力する
'
' @Remarks
' 共通CSSを出力する
'
' EnableResize=Trueの場合は
' 列幅変更用CSSも出力する
'
'===================================================================
Public Sub WriteHtmlHeader( _
                ByVal FileNo As Integer, _
                ByVal Title As String, _
                Optional ByVal Line_Height As Single = 1.5, _
                Optional ByVal EnableResize As Boolean = False)

    '# HTML開始
    Print #FileNo, "<!DOCTYPE html>"
    
    '# 日本語ページ指定
    Print #FileNo, "<html lang=""ja"">"

    '# HEAD開始
    Print #FileNo, "<head>"

    '# 文字コード
    Print #FileNo, _
        "<meta charset=""shift_jis"">"
    
    '# Google翻訳抑止
    Print #FileNo, _
        "<meta name=""google"" content=""notranslate"">"

    '# 電話番号リンク化抑止
    Print #FileNo, _
        "<meta name=""format-detection"" content=""telephone=no"">"
    
    '# ブラウザ言語指定
    Print #FileNo, _
        "<meta http-equiv=""Content-Language"" content=""ja"">"

    '# タイトル
    Print #FileNo, _
        "<title>" & _
        Title & _
        "</title>"

    '# CSS開始
    Print #FileNo, _
        "<style>"

    '# 翻訳抑止
    Print #FileNo, _
        ".notranslate{" & _
        "translate:no;" & _
        "}"

    '# コメント色付け
    Print #FileNo, _
        ".hcp-comment{" & _
        "color:#008000;" & _
        "font-style:italic;" & _
        "white-space:pre;" & _
        "}"


    '# Body
    Print #FileNo, _
        "body{font-family:Meiryo;font-size:9pt;}"

    '# Table
    Print #FileNo, _
        "table{border-collapse:collapse;}"

    '# Cell
    Print #FileNo, _
        "th,td{border:1px solid #cccccc;" & _
        "padding:1px 4px;" & _
        "line-height:" & _
        Line_Height & ";}"

    '# DataCell
    Print #FileNo, _
        "td{vertical-align:top;}"

    '# HeaderCell
    Print #FileNo, _
        "th{background:#e8f0fe;}"

    '# 列幅変更CSS
    If EnableResize Then

        WriteColumnResizerCss _
            FileNo

    End If

    '# CSS終了
    Print #FileNo, _
        "</style>"

    '# HEAD終了
    Print #FileNo, _
        "</head>"

    '# BODY開始
    Print #FileNo, _
        "<body translate=""no"">"

End Sub

' 共通フッタ
'===================================================================
' @JPName
' HTMLフッタ出力
'
' @Category
' HCP
'
' @Input
' FileNo       : 出力ファイル番号
' EnableResize : 列幅変更有無
'
' @Output
' なし
'
' @Summary
' HTML共通フッタを出力する
'
' @Remarks
' EnableResize=Trueの場合は
' 列幅変更用JavaScriptを出力する
'
' BODYおよびHTMLタグを
' 閉じる
'
'===================================================================
Public Sub WriteHtmlFooter( _
                ByVal FileNo As Integer, _
                Optional ByVal EnableResize As Boolean = False)

    '# 列幅変更Script
    If EnableResize Then

        WriteColumnResizerScript _
            FileNo

    End If

    '# BODY終了
    Print #FileNo, _
        "</body>"

    '# HTML終了
    Print #FileNo, _
        "</html>"

End Sub

'===================================================================
' @JPName
' 列幅変更CSS出力
'
' @Category
' HCP
'
' @Input
' FileNo : 出力ファイル番号
'
' @Output
' なし
'
' @Summary
' テーブル列幅変更用CSSを出力する
'
' @Remarks
' resizeクラスおよび
' resizerクラスを定義する
'===================================================================
Public Sub WriteColumnResizerCss( _
                ByVal FileNo As Integer)

    '# resizeクラス
    Print #FileNo, _
        ".resize{position:relative;}"

    '# リサイズハンドル
    Print #FileNo, _
        ".resizer{" & _
        "position:absolute;" & _
        "right:-10px;" & _
        "top:0;" & _
        "width:20px;" & _
        "height:100%;" & _
        "cursor:col-resize;" & _
        "user-select:none;" & _
        "z-index:100;}"

End Sub

'===================================================================
' @JPName
' 列幅変更Script出力
'
' @Category
' HCP
'
' @Input
' FileNo : 出力ファイル番号
'
' @Output
' なし
'
' @Summary
' テーブル列幅変更用JavaScriptを
' 出力する
'
' @Remarks
' resizeクラスを持つ
' TH要素に対して
' ドラッグによる列幅変更を
' 有効化する
'===================================================================
Public Sub WriteColumnResizerScript( _
                ByVal FileNo As Integer)

    '# Script開始
    Print #FileNo, _
        "<script>"

    '# resize対象取得
    Print #FileNo, _
        "document.querySelectorAll('.resize').forEach(function(eleTh) {"

    '# 開始位置
    Print #FileNo, _
        "let startX;"

    '# 開始幅
    Print #FileNo, _
        "let startWidth;"

    '# リサイズハンドル生成
    Print #FileNo, _
        "const eleResizer = document.createElement('div');"

    Print #FileNo, _
        "eleResizer.classList.add('resizer');"

    '# ドラッグ開始
    Print #FileNo, _
        "eleResizer.addEventListener('mousedown', function(e) {"

    Print #FileNo, _
        "startX = e.pageX;"

    Print #FileNo, _
        "startWidth = eleTh.offsetWidth;"

    Print #FileNo, _
        "document.addEventListener('mousemove', onMouseMove);"

    Print #FileNo, _
        "document.addEventListener('mouseup', onMouseUp);"

    Print #FileNo, _
        "});"

    '# ハンドル追加
    Print #FileNo, _
        "eleTh.appendChild(eleResizer);"

    '# ドラッグ中
    Print #FileNo, _
        "function onMouseMove(e){"

    Print #FileNo, _
        "eleTh.style.width=" & _
        "(startWidth+(e.pageX-startX))+'px';"

    Print #FileNo, _
        "}"

    '# ドラッグ終了
    Print #FileNo, _
        "function onMouseUp(){"

    Print #FileNo, _
        "document.removeEventListener('mousemove',onMouseMove);"

    Print #FileNo, _
        "document.removeEventListener('mouseup',onMouseUp);"

    Print #FileNo, _
        "}"

    Print #FileNo, _
        "});"

    '# Script終了
    Print #FileNo, _
        "</script>"

End Sub


