Attribute VB_Name = "modHCPPath"
Option Explicit

'===================================================================
'@JPName      フォルダ・パス管理
'===================================================================

'--------------------------------------------------
' パス管理
'--------------------------------------------------

'
'
' @JPName
' HCP出力フォルダ取得
'
Public Function GetHCPOutputFolder( _
                ByVal ModuleName As String) _
                As String

    Dim Tool As clsToolInfo

    Dim FolderPath As String

    Set Tool = _
        GetCurrentToolInfo()

    FolderPath = _
        GetOutputFolder() & _
        "HCP\" & _
        Tool.ToolName & "\" & _
        ModuleName & "\"

    CreateFolderIfNotExists _
        GetOutputFolder() & _
        "HCP"

    CreateFolderIfNotExists _
        GetOutputFolder() & _
        "HCP\" & _
        Tool.ToolName

    CreateFolderIfNotExists _
        FolderPath

    GetHCPOutputFolder = _
        FolderPath

End Function

'===================================================================
' @JPName
' INDEX相対パス取得
'
' @Category
' HCP
'
' @Input
' Level : 親階層数
'
' @Output
' String
'
' @Summary
' 指定階層からの
' _INDEX.html相対パスを返却する
'
' @Remarks
' Level=0 : _INDEX.html
' Level=1 : ../_INDEX.html
' Level=2 : ../../_INDEX.html
'===================================================================
Public Function GetIndexRelativePath( _
                ByVal Level As Long) _
                As String

    '# 走査用
    Dim i As Long

    For i = 1 To Level

        GetIndexRelativePath = _
            GetIndexRelativePath & "../"

    Next i

    GetIndexRelativePath = _
        GetIndexRelativePath & _
        "_INDEX.html"

End Function

'--------------------------------------------------
' フォルダ管理
'--------------------------------------------------

'
' @JPName
' フォルダ作成
'
Public Sub CreateFolderIfNotExists( _
                ByVal FolderPath As String)

    If Dir( _
        FolderPath, _
        vbDirectory) = "" Then

        MkDir FolderPath

    End If

End Sub


'===================================================================
' @JPName
' HCPツールフォルダ取得
'
' @Category
' HCP
'
' @Input
' なし
'
' @Output
' String
'
' @Summary
' 現在選択中ツールの
' HCP出力ルートを返却する
'
' @Remarks
' Tool単位の_INDEX.html生成で使用
'===================================================================
Public Function GetHCPToolFolder() _
                As String

    '# ツール情報
    Dim Tool As clsToolInfo

    '# フォルダパス
    Dim FolderPath As String

    Set Tool = _
        GetCurrentToolInfo()

    FolderPath = _
        GetOutputFolder() & _
        "HCP\" & _
        Tool.ToolName & "\"

    CreateFolderIfNotExists _
        GetOutputFolder() & _
        "HCP"

    CreateFolderIfNotExists _
        FolderPath

    GetHCPToolFolder = _
        FolderPath

End Function



