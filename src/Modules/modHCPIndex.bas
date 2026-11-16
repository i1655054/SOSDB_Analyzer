Attribute VB_Name = "modHCPIndex"
Option Explicit

'===================================================================
'@JPName      _INDEX.html 生成
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
' HCPファイルを走査し毎回再構築する
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
' モジュールフォルダを走査して
' 毎回再構築する
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
' HCP_TARGETSを利用して
' モジュール一覧を取得する
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
'===================================================================
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
' _INDEX.html確認用
'===================================================================
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


'===================================================================
' @JPName
' INDEX整合性確認
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
' モジュールINDEXの存在を確認する
'
' @Remarks
' INDEX欠落時はWARNINGを出力する
'===================================================================
Public Sub CheckHCPIndexes()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim ModuleName As String
    
    Dim Checked As Object

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)
    
    Set Checked = _
        CreateObject( _
            "Scripting.Dictionary")

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_HCP_MODULE).End(xlUp).row

    For r = 2 To LastRow

        ModuleName = _
            Trim$( _
                ws.Cells( _
                    r, _
                    COL_HCP_MODULE).Value)

        If ModuleName <> "" Then

            If Not Checked.Exists( _
                    ModuleName) Then
                
                Checked.Add _
                    ModuleName, _
                    True
                
                If Not ExistsModuleIndex( _
                    ModuleName) Then

                    Trace "WARNING", _
                        "Missing Index : " & _
                        ModuleName

                End If
            
            End If

        End If

    Next r

End Sub


'===================================================================
' @JPName
' INDEX自動修復
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
' 不足しているINDEXを再生成する
'
' @Remarks
' HCP_TARGETSを走査し
' 欠落モジュールINDEXおよび
' ツールINDEXを再生成する
'===================================================================
Public Sub RepairHCPIndexes()

    '# HCP_TARGETS
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# モジュール名
    Dim ModuleName As String

    '# 重複防止
    Dim Checked As Object

    '# 修復件数
    Dim RepairCount As Long

    Trace "INDEX", _
        "RepairHCPIndexes Start"

    Set Checked = _
        CreateObject( _
            "Scripting.Dictionary")

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_HCP_MODULE).End(xlUp).row

    '# INDEX存在確認
    For r = 2 To LastRow

        ModuleName = _
            Trim$( _
                ws.Cells( _
                    r, _
                    COL_HCP_MODULE).Value)

        If ModuleName <> "" Then

            If Not Checked.Exists( _
                    ModuleName) Then

                Checked.Add _
                    ModuleName, _
                    True

                If Not ExistsModuleIndex( _
                        ModuleName) Then

                    Trace "INDEX", _
                        "Missing ModuleIndex : " & _
                        ModuleName

                    RebuildModuleIndex _
                        ModuleName

                    Trace "INDEX", _
                        "Repaired ModuleIndex : " & _
                        ModuleName

                    RepairCount = _
                        RepairCount + 1

                End If

            End If

        End If

    Next r

    '# ツールINDEX再生成
    If Not ExistsToolIndex Then

        Trace "INDEX", _
            "Missing ToolIndex"

    Else

        Trace "INDEX", _
            "ToolIndex Exists"

    End If

    RebuildToolIndex

    Trace "INDEX", _
        "Repaired ToolIndex"

    Trace "INDEX", _
        "RepairHCPIndexes End" & _
        " ModuleRepair=" & _
        RepairCount

End Sub

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

' 共通ヘッダ
Public Sub WriteHtmlHeader( _
                ByVal FileNo As Integer, _
                ByVal Title As String, _
                Optional ByVal Line_Height As Single = 1.5)
                
    Print #FileNo, "<!DOCTYPE html>"
    Print #FileNo, "<html>"

    Print #FileNo, "<head>"

    Print #FileNo, _
        "<meta charset=""shift_jis"">"

    Print #FileNo, _
        "<title>" & _
        Title & _
        "</title>"

    Print #FileNo, _
        "<style>"

    Print #FileNo, _
        "body{font-family:Meiryo;font-size:9pt;}"

    Print #FileNo, _
        "table{border-collapse:collapse;}"

    'Print #FileNo, _
    '    "th,td{border:1px solid #cccccc;" & _
    '    "padding:1px 4px;}"

    Print #FileNo, _
        "th,td{border:1px solid #cccccc;" & _
        "padding:1px 4px;line-height: " & Line_Height & ";}"

    Print #FileNo, _
        "td{vertical-align:top;}"

    Print #FileNo, _
        "th{background:#e8f0fe;}"
    
    ' --- 追加 ここから ---
    Print #FileNo, _
        "colgroup col:nth-of-type(1) {}"

    Print #FileNo, _
        "colgroup col:nth-of-type(2) {}"

    Print #FileNo, _
        ".resize {position: relative;}"
    
    Print #FileNo, _
        ".resizer {" & _
            "position: absolute;" & _
            "right: -10px;" & _
            "top: 0;" & _
            "width: 20px;" & _
            "height: 100%;" & _
            "cursor: col-resize;" & _
            "user-select: none;" & _
            "z-index: 100;}"
    
    ' --- 追加 ここから ---
    
    Print #FileNo, _
        "</style>"

    Print #FileNo, _
        "</head>"

    Print #FileNo, _
        "<body>"

End Sub

' 共通フッタ
Public Sub WriteHtmlFooter( _
                ByVal FileNo As Integer)

    ' --- 修正 追加 ここから ---
    Print #FileNo, _
        "<script>"
    
    Print #FileNo, _
        "document.querySelectorAll('.resize').forEach(function (eleTh) {"

    Print #FileNo, _
        "let startX;"
    Print #FileNo, _
        "let startWidth;"
        
    Print #FileNo, _
        "const eleResizer = document.createElement('div');"
    Print #FileNo, _
        "eleResizer.classList.add('resizer');"
    Print #FileNo, _
        "eleResizer.addEventListener('mousedown', function (e) {"
    Print #FileNo, _
        "const styleTh = window.getComputedStyle(eleTh);"
    Print #FileNo, _
        "const paddingLeft = parseFloat(styleTh.paddingLeft);"
    Print #FileNo, _
        "const paddingRight = parseFloat(styleTh.paddingRight);"
    Print #FileNo, _
        "const borderLeft = parseFloat(styleTh.borderLeftWidth);"
    Print #FileNo, _
        "const borderRight = parseFloat(styleTh.borderRightWidth);"
    Print #FileNo, _
        "startX = e.pageX;"
    Print #FileNo, _
        "startWidth = eleTh.offsetWidth - (paddingLeft + paddingRight + borderLeft + borderRight);"
    Print #FileNo, _
        "document.addEventListener('mousemove', onMouseMove);"
    Print #FileNo, _
        "document.addEventListener('mouseup', onMouseUp);"
    Print #FileNo, _
        "});"
    Print #FileNo, _
        "eleTh.appendChild(eleResizer);"
    
    Print #FileNo, _
        "function onMouseMove(e) {"
    Print #FileNo, _
        "const newWidth = startWidth + (e.pageX - startX);"
    Print #FileNo, _
        "eleTh.style.width = newWidth + 'px';"
    Print #FileNo, _
        "}"
    
    Print #FileNo, _
        "function onMouseUp() {"
    Print #FileNo, _
        "document.removeEventListener('mousemove', onMouseMove);"
    Print #FileNo, _
        "document.removeEventListener('mouseup', onMouseUp);"
    Print #FileNo, _
        "}"
    Print #FileNo, _
        "});"

    
    Print #FileNo, _
        "</script>"
    
    ' --- 修正 追加 ここまで ---
    
    Print #FileNo, _
        "</body>"

    Print #FileNo, _
        "</html>"

End Sub

