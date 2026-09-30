Attribute VB_Name = "modPathChartCommon"
Option Explicit

'====================================================
' 子Procedure一覧取得
'====================================================

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
' Collection
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

    Dim Col As Collection

    Dim wsDep As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim ChildProc As String
    Dim ParentModule As String
    Dim ChildModule As String

    Set Col = New Collection

    Set wsDep = _
        ThisWorkbook.Worksheets( _
            "関数依存関係")

    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
                Trim$(wsDep.Cells(r, 1).Value), _
                ParentProc, _
                vbTextCompare) = 0 Then

            ChildProc = _
                Trim$(wsDep.Cells(r, 3).Value)

            ParentModule = _
                Trim$(wsDep.Cells(r, 6).Value)

            ChildModule = _
                Trim$(wsDep.Cells(r, 7).Value)

            If IsTargetRelation( _
                    ParentProc, _
                    ChildProc, _
                    ParentModule, _
                    ChildModule) Then

                Col.Add ChildProc

            End If

        End If

    Next r

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

    IsTargetRelation = True

    '------------------------------
    ' 再帰除外
    '------------------------------
    If Not gRecursive Then

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
    If Not gExternal Then

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
    If gSkipCommon Then

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
' Boolean
'
' @Summary
' 共通関数として除外対象か判定する
'
Public Function IsCommonFunction( _
                ByVal ProcName As String) As Boolean

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
' clsTreeNode
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

    Dim Node As clsTreeNode
    Dim Proc As clsProcInfo

    Set Node = New clsTreeNode

    Node.Level = Level
    Node.ProcName = ProcName
    Node.ParentProcName = ParentProcName
    Node.TreeText = TreeText

    '--------------------------
    ' Procedure情報
    '--------------------------
    Set Proc = GetProcInfo(ProcName)

    If Not Proc Is Nothing Then

        Node.JPName = Proc.JPName
        Node.Category = Proc.Category
        Node.ModuleName = Proc.ModuleName
        Node.Summary = Proc.Summary

    End If

    Set CreateTreeNode = Node

End Function

