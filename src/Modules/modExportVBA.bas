Attribute VB_Name = "modExportVBA"
Option Explicit

'===================================================================
'@JPName     VBAソース管理
'@Category   Export
'@Summary    VBAソースエクスポートおよびGit連携機能を提供する
'===================================================================

' 標準モジュール (*.bas)
' VBComponent.Type
' vbext_ct_StdModule
Const VBEXT_CT_STDMODULE As Long = 1

' クラスモジュール (*.cls)
' VBComponent.Type
' vbext_ct_ClassModule
Const VBEXT_CT_CLASSMODULE As Long = 2

' ユーザーフォーム (*.frm)
' VBComponent.Type
' vbext_ct_MSForm
Const VBEXT_CT_MSFORM As Long = 3

' ドキュメントモジュール
' (ThisWorkbook、Sheet1 など)
' VBComponent.Type
' vbext_ct_Document
Const VBEXT_CT_DOCUMENT As Long = 100

'**************************************
'====  VBA ソース自動エクスポート  ====
'
' Excelブック内のVBAプロジェクトを、
' ボタン一発で .bas .cls .frm に出力する
'**************************************

' @JPName
' VBAソース出力
'
' @Category
' Export
'
' @Input
' TargetBook(Workbook)
' ToolID(Long)
'
' @Output
' なし
'
' @Summary
' 指定ブックのVBAソースを
' GitHub管理用フォルダへ出力する
'
' @Remarks
' 標準モジュール(.bas)
' クラスモジュール(.cls)
' フォーム(.frm)
' ドキュメントモジュール(.cls)
'
' をGit管理用フォルダへ
' 種別毎に出力する
'
' 既存ファイルが存在する場合は
' 上書き出力する
'
Public Sub ExportVBABook( _
                ByVal TargetBook As Workbook, _
                ByVal ToolID As Long)
    
    ' VBAコンポーネント
    Dim VBComp As Object

    ' Gitルートフォルダ
    Dim RootPath As String
    
    ' 標準モジュール出力先
    Dim ModulesPath As String
    
    ' クラスモジュール出力先
    Dim ClassesPath As String
    
    ' フォーム出力先
    Dim FormsPath As String
    
    ' ドキュメントモジュール出力先
    Dim DocumentsPath As String
    
    ' FileSystemObject
    Dim FSO As Object

    ' エクスポートファイル名
    Dim ExportFile As String

    '# Git管理フォルダ取得
    RootPath = GetGitRootPathByToolID(ToolID)
    
    '# モジュール出力先設定
    ModulesPath = RootPath & "\src\Modules\"
    ClassesPath = RootPath & "\src\Classes\"
    FormsPath = RootPath & "\src\Forms\"
    DocumentsPath = RootPath & "\src\Documents\"

    '# フォルダ操作オブジェクト生成
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    '# srcフォルダ存在確認
    If Not FSO.FolderExists(RootPath & "\src") Then
        FSO.CreateFolder RootPath & "\src"
    End If
    
    '# Modulesフォルダ作成
    If Not FSO.FolderExists(ModulesPath) Then
        FSO.CreateFolder ModulesPath
    End If

    '# Classesフォルダ作成
    If Not FSO.FolderExists(ClassesPath) Then
        FSO.CreateFolder ClassesPath
    End If

    '# Formsフォルダ作成
    If Not FSO.FolderExists(FormsPath) Then
        FSO.CreateFolder FormsPath
    End If

    '# Documentsフォルダ作成
    If Not FSO.FolderExists(DocumentsPath) Then
        FSO.CreateFolder DocumentsPath
    End If

    #If DEBUG_MODE Then
        Debug.Print "ModulesPath=" & ModulesPath
        Debug.Print "ClassesPath=" & ClassesPath
        Debug.Print "FormsPath=" & FormsPath
        Debug.Print "DocumentsPath=" & DocumentsPath
    #End If

    '# VBAコンポーネント出力開始
    For Each VBComp In TargetBook.VBProject.VBComponents

        '# コンポーネント種別毎に出力先決定
        Select Case VBComp.Type

            Case VBEXT_CT_STDMODULE
                ExportFile = ModulesPath & VBComp.Name & ".bas"

            Case VBEXT_CT_CLASSMODULE
                ExportFile = ClassesPath & VBComp.Name & ".cls"

            Case VBEXT_CT_MSFORM
                ExportFile = FormsPath & VBComp.Name & ".frm"

            Case VBEXT_CT_DOCUMENT
                ExportFile = DocumentsPath & VBComp.Name & ".cls"

            Case Else
                ExportFile = ""

        End Select

        If ExportFile <> "" Then

            '# 既存ソース削除
            If Dir(ExportFile) <> "" Then

                Kill ExportFile

            End If

            #If DEBUG_MODE Then
                Debug.Print "--------------------------------"
                Debug.Print "VBComp.Name=" & VBComp.Name
                Debug.Print "ExportFile=" & ExportFile
                Debug.Print "Exists=" & (Dir$(Left$(ExportFile, InStrRev(ExportFile, "\") - 1), vbDirectory) <> "")
            #End If
            
            '# VBAソースエクスポート
            VBComp.Export ExportFile

        End If

    Next VBComp

End Sub

' @JPName
' Git管理フォルダ取得
'
' @Category
' Export
'
' @Input
' ToolID(Long)
'
' @Output
' GitRootPath(String)
'
' @Summary
' ToolIDに対応する
' GitHub管理フォルダを取得する
'
' @Remarks
' OneDrive\ドキュメント\GitHub配下の
' リポジトリ構成を前提とする
'
Public Function GetGitRootPathByToolID( _
            ByVal ToolID As Long) As String

    ' ツール情報
    Dim Tool As clsToolInfo

    '# ツール情報取得
    Set Tool = GetToolInfo(ToolID)

    '# Git管理フォルダ生成
    GetGitRootPathByToolID = _
        Environ$("OneDriveCommercial") & _
        "\ドキュメント\GitHub\" & _
        Tool.GitRepository

End Function

'==================================================
' Git連携
'==================================================

' @JPName
' Gitフォルダ起動
'
' @Category
' Export
'
' @Input
' RootPath(String)
'
' @Output
' なし
'
' @Summary
' 指定フォルダをカレントとして
' PowerShellを起動する
'
' @Remarks
' PowerShellを
' -NoExit オプションで起動する
'
' 起動後のカレントディレクトリを
' Git管理フォルダへ変更する
'
' Git Add
' Git Commit
' Git Push
'
' 等の実行支援として利用する
'
Public Sub OpenGitPowerShell( _
                    ByVal RootPath As String)

    ' PowerShell起動コマンド
    Dim Cmd As String

    '# PowerShell起動コマンド生成
    Cmd = _
        "powershell.exe -NoExit -Command " & _
        """Set-Location '" & RootPath & "'"""

    '# GitフォルダでPowerShell起動
    Shell Cmd, vbNormalFocus

End Sub
