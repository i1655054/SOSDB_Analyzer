Attribute VB_Name = "modPathChartCommon"
Option Explicit

'====================================================
' 子Procedure一覧取得
'====================================================
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

        If IsCommonFunctionSafe( _
                ChildProc) Then

            IsTargetRelation = False
            Exit Function

        End If

    End If

End Function

'====================================================
' 共通関数判定
'====================================================
Public Function IsCommonFunctionSafe( _
                ByVal ProcName As String) _
                As Boolean

    Select Case UCase$(ProcName)

        Case "CHK_PERSON", _
             "CHK_TARGET", _
             "CHK_PRIORITY", _
             "CHK_EQ_PRODUCT", _
             "CHK_ESCALE_DAY", _
             "CHK_TORBLENUM"

            IsCommonFunctionSafe = True

    End Select

End Function

'====================================================
' TreeNode生成
'====================================================
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

'====================================================
' 日本語名取得
'====================================================
Public Function GetJPNameSafe( _
                ByVal ProcName As String) _
                As String

    Dim Proc As clsProcInfo

    Set Proc = GetProcInfo(ProcName)

    If Proc Is Nothing Then Exit Function

    GetJPNameSafe = Proc.JPName

End Function

'====================================================
' Tree表示文字列
'====================================================
Public Function GetTreeDisplayText( _
                ByVal ProcName As String) _
                As String

    GetTreeDisplayText = ProcName

End Function

'====================================================
' 固定長補完
'====================================================
Public Function PadRight( _
                ByVal Text As String, _
                ByVal Width As Long) _
                As String

    If Len(Text) >= Width Then

        PadRight = Text

    Else

        PadRight = _
            Text & _
            Space$(Width - Len(Text))

    End If

End Function

'====================================================
' Procedureヘッダ生成
'====================================================
Public Function GetProcHeaderText( _
                ByVal ProcName As String) _
                As String

    Dim Proc As clsProcInfo
    Dim Buf As String

    Set Proc = GetProcInfo(ProcName)

    If Proc Is Nothing Then

        GetProcHeaderText = _
            "Procedure : " & ProcName

        Exit Function

    End If

    Buf = ""
    Buf = Buf & _
        "Procedure : " & Proc.ProcName & vbCrLf

    Buf = Buf & _
        "JPName    : " & Proc.JPName & vbCrLf

    Buf = Buf & _
        "Category  : " & Proc.Category & vbCrLf

    Buf = Buf & _
        "Module    : " & Proc.ModuleName & vbCrLf

    Buf = Buf & _
        "Summary   : " & Proc.Summary

    GetProcHeaderText = Buf

End Function

