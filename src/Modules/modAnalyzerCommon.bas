Attribute VB_Name = "modAnalyzerCommon"
Option Explicit

Private Const GIT_REPOSITORY As String = "SOSDB_VBA"

' 対象ブック取得
Public Function GetTargetBook( _
                Optional ByVal BookName As String = _
                "●SOSDB問診票登録.xlsm") _
                As Workbook

    On Error Resume Next
    Set GetTargetBook = Workbooks(BookName)
    On Error GoTo 0

End Function

' 対象ブックの状態確認
Public Function CheckTargetBook( _
                    ByVal TargetBook As Workbook) As Boolean

    If TargetBook Is Nothing Then

        MsgBox _
            "●SOSDB問診票登録.xlsm が開かれていません。", _
            vbExclamation

        Exit Function

    End If

    CheckTargetBook = True

End Function

' SOSDB_VBA のルート取得
Public Function GetGitRootPath() As String

    If Left$(LCase$(ThisWorkbook.Path), 5) = "https" Then

        GetGitRootPath = _
            Environ$("OneDriveCommercial") & _
            "\ドキュメント\GitHub\" & GIT_REPOSITORY

    Else

        GetGitRootPath = _
            CreateObject("Scripting.FileSystemObject") _
                .GetParentFolderName(ThisWorkbook.Path)

    End If

End Function

' 分析資料保存先
Public Function GetAnalysisFolder() As String

    GetAnalysisFolder = _
        GetGitRootPath & _
        "\doc\analysis"

End Function

' Ver比較フォルダ
Public Function GetCompareFolder() As String

    GetCompareFolder = _
        GetGitRootPath & _
        "\..\Compare"

End Function

Public Function IsTargetBookOpen() As Boolean

    IsTargetBookOpen = _
        Not (GetTargetBook() Is Nothing)

End Function

