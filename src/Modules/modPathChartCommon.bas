Attribute VB_Name = "modPathChartCommon"
Option Explicit

'===================================================================
' @JPName
' TextPathChart生成
'
' @Category
' PathChart
'
' @Summary
' PathChart情報をツリー形式テキストへ変換し
' PathChartTreeシートへ出力する
'
' @Remarks
' TreeNode生成
' 英語Tree生成
' 日本語Tree生成
' Treeシート出力
'
' を提供する
'===================================================================

' @JPName
' 子関数一覧取得
'
' @Category
' PathChart
'
' @Input
' ParentProc(String)
'
' @Output
' Children(Collection)
'
' @Summary
' 指定プロシージャの子プロシージャ一覧を取得する
'
' @Remarks
' 関数依存関係シートを参照し
' 出力条件に一致する子プロシージャのみ返す
'
Public Function GetChildrenList( _
                ByVal ParentProc As String) _
                As Collection

    ' 子Procedure一覧
    Dim Col As Collection

    ' 関数依存関係シート
    Dim wsDep As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 子Procedure名
    Dim ChildProc As String
    
    ' 呼出元モジュール
    Dim ParentModule As String
    
    ' 呼出先モジュール
    Dim ChildModule As String

    '-----------------------------
    '# 初期化
    '-----------------------------
    '# 子Procedure一覧初期化
    Set Col = New Collection

    '-----------------------------
    '# 関数依存関係取得
    '-----------------------------
    '# 関数依存関係シート取得
    Set wsDep = _
        gContext.ManagementBook.Worksheets( _
            SHEET_FUNCTION_DEPENDENCY)

    '# 関数依存関係最終行取得
    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    '-----------------------------
    '# 子Procedure探索
    '-----------------------------
    For r = 2 To LastRow

        '# 親Procedure一致判定
        If StrComp( _
                Trim$(wsDep.Cells(r, 1).Value), _
                ParentProc, _
                vbTextCompare) = 0 Then

            '# 呼出関係情報取得
            ChildProc = _
                Trim$(wsDep.Cells(r, 3).Value)

            ParentModule = _
                Trim$(wsDep.Cells(r, 6).Value)

            ChildModule = _
                Trim$(wsDep.Cells(r, 7).Value)

            '# 出力対象判定
            If IsTargetRelation( _
                    ParentProc, _
                    ChildProc, _
                    ParentModule, _
                    ChildModule) Then

                '# 子Procedure追加
                Col.Add ChildProc

            End If

        End If

    Next r

    '# 子Procedure一覧返却
    Set GetChildrenList = Col

End Function

'====================================================
' 出力対象判定
'====================================================

' @JPName
' 出力対象判定
'
' @Category
' PathChart
'
' @Input
' ParentProc(String)
' ChildProc(String)
' ParentModule(String)
' ChildModule(String)
'
' @Output
' Boolean
'
' @Summary
' PathChart出力対象の呼出関係か判定する
'
' @Remarks
' 再帰呼出除外
' 外部モジュール除外
' 共通関数除外
' を考慮して判定する
'
Public Function IsTargetRelation( _
                ByVal ParentProc As String, _
                ByVal ChildProc As String, _
                ByVal ParentModule As String, _
                ByVal ChildModule As String) _
                As Boolean

    '------------------------------
    '# 初期状態
    '------------------------------
    
    '# 出力対象初期設定
    IsTargetRelation = True

    '------------------------------
    '# 再帰呼出判定
    '------------------------------
    If Not gRecursive Then

        '# 再帰呼出除外
        If StrComp( _
                ParentProc, _
                ChildProc, _
                vbTextCompare) = 0 Then

            IsTargetRelation = False
            Exit Function

        End If

    End If

    '------------------------------
    ' 外部モジュール除外
    '------------------------------
    '# 外部モジュール判定
    If Not gExternal Then

        '# 外部モジュール除外
        If StrComp( _
                ParentModule, _
                ChildModule, _
                vbTextCompare) <> 0 Then

            IsTargetRelation = False
            Exit Function

        End If

    End If

    '------------------------------
    ' 共通関数除外
    '------------------------------
    '# 共通関数判定
    If gSkipCommon Then

        '# 共通関数除外
        If IsCommonFunction( _
                ChildProc) Then

            IsTargetRelation = False
            Exit Function

        End If

    End If

End Function

'====================================================
' 共通関数判定
'====================================================

' @JPName
' 共通関数判定
'
' @Category
' PathChart
'
' @Input
' ProcName(String)
'
' @Output
' IsCommon(Boolean)
'
' @Summary
' 共通関数として除外対象か判定する
'
' @Remarks
' PathChart出力時に除外対象とする
' 共通ユーティリティ関数か判定する
'
Public Function IsCommonFunction( _
                ByVal ProcName As String) As Boolean

    '------------------------------
    '# 共通関数判定
    '------------------------------
    Select Case UCase$(ProcName)

        Case "CHK_PERSON", _
             "CHK_TARGET", _
             "CHK_PRIORITY", _
             "CHK_EQ_PRODUCT", _
             "CHK_ESCALE_DAY", _
             "CHK_TORBLENUM"
             '"VALIDATEANDHIGHLIGHT", _
             '"RESTORECOLOR", _
             '"GETDEFAULTCOLOR", _
             '"HIGHLIGHTCELL"

            IsCommonFunction = True

    End Select

End Function

'====================================================
' TreeNode生成
'====================================================

' @JPName
' ツリーノード生成
'
' @Category
' PathChart
'
' @Input
' ProcName(String)
' ParentProcName(String)
' Level(Long)
' TreeText(String)
'
' @Output
' TreeNode(clsTreeNode)
'
' @Summary
' PathChartおよびTextPathChart用の
' ツリーノードを生成する
'
' @Remarks
' ProcListの情報を取得して
' JPName、Category、ModuleName、Summaryを設定する
'
Public Function CreateTreeNode( _
                ByVal ProcName As String, _
                ByVal ParentProcName As String, _
                ByVal Level As Long, _
                ByVal TreeText As String) _
                As clsTreeNode

    ' ツリーノード
    Dim Node As clsTreeNode
    
    ' Procedure情報
    Dim Proc As clsProcInfo

    '--------------------------
    '# TreeNode生成
    '--------------------------
    Set Node = New clsTreeNode

    '--------------------------
    '# TreeNode基本情報
    '--------------------------
    '# TreeNode基本情報設定
    Node.Level = Level
    Node.ProcName = ProcName
    Node.ParentProcName = ParentProcName
    Node.TreeText = TreeText

    '--------------------------
    ' Procedure情報
    '--------------------------
    '# Procedure情報取得
    Set Proc = _
        ResolveProcedure( _
            "", _
            ProcName)

    '# Procedure詳細情報設定
    If Not Proc Is Nothing Then

        Node.JPName = Proc.JPName
        Node.Category = Proc.Category
        Node.ModuleName = Proc.ModuleName
        Node.Summary = Proc.Summary

    End If

    '# TreeNode返却
    Set CreateTreeNode = Node

End Function

