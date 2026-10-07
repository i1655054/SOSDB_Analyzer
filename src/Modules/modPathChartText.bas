Attribute VB_Name = "modPathChartText"
Option Explicit

' TreeNode一覧
Private gTreeNodes As Collection

'==================================================
' Tree表示文字
'==================================================

' 縦罫線
Private Const TREE_BAR As String = "│"

' ノード記号
Private Const TREE_NODE_LAST As String = "└─"
Private Const TREE_NODE_NEXT As String = "├─"

' 空白
Private Const TREE_SPACE As String = "　"

' 階層継続(日本語Tree)
Private Const TREE_CONTINUE As String = "│　"

' 次階層Prefix
Private Const TREE_NEXT_LEVEL As String = "│　"

' 最終階層Prefix
Private Const TREE_LAST_LEVEL As String = "　　"

'==================================================
' Tree表示
'==================================================

' 日本語Tree行文字色 RGB(80,80,80)
Private Const TREE_JP_FONT_COLOR As Long = &H505050

' Treeフォント
Private Const TREE_FONT_NAME As String = "ＭＳ ゴシック"

' Tree列幅
Private Const TREE_COLUMN_WIDTH As Double = 80

'==================================================
' テキストチャート生成
'==================================================

' @JPName
' テキストチャート出力
'
' @Category
' PathChart
'
' @Input
' StartProc(String)
' MaxDepth(Long)
'
' @Output
' PathChartTreeシート
'
' @Summary
' PathChart情報をツリー形式の
' テキストとして出力する
'
' @Remarks
' TreeNode一覧を生成し
' PathChartTreeシートへ出力する
'
Public Sub ExportPathChartToText( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    Set gTreeNodes = New Collection

    AddRootNode StartProc

    BuildTreeText _
            StartProc, _
            "", _
            1, _
            MaxDepth

    OutputTreeNodeList

End Sub

'==================================================
' ルートノード追加
'==================================================

' @JPName
' ルートノード追加
'
' @Category
' PathChart
'
' @Input
' StartProc(String)
'
' @Output
' なし
'
' @Summary
' ツリー構造のルートノードを生成し
' TreeNode一覧へ追加する
'
Private Sub AddRootNode( _
                ByVal StartProc As String)

    Dim Node As clsTreeNode

    Set Node = CreateTreeNode( _
                    StartProc, _
                    "", _
                    0, _
                    StartProc)

    If Trim$(Node.ModuleName) <> "" Then

        If gShowModule Then
    
            Node.TreeText = _
                Node.TreeText & _
                " <" & _
                Node.ModuleName & _
                ">"

        End If

    End If
    
    If Trim$(Node.JPName) <> "" Then

        Node.JPLine = _
            TREE_BAR & _
            Node.JPName

    End If

    gTreeNodes.Add Node

End Sub

'==================================================
' 木構造生成
'==================================================

' @JPName
' ツリー構造生成
'
' @Category
' PathChart
'
' @Input
' ParentProc(String)
' Prefix(String)
' Level(Long)
' MaxDepth(Long)
'
' @Output
' なし
'
' @Summary
' 関数依存関係を再帰的にたどり
' TreeNode一覧を生成する
'
' @Remarks
' 英語Tree行および日本語Tree行を生成する
'
Private Sub BuildTreeText( _
                ByVal ParentProc As String, _
                ByVal Prefix As String, _
                ByVal Level As Long, _
                ByVal MaxDepth As Long, _
                Optional ByVal PathText As String = "")

    '------------------------------------
    ' 日本語Tree行生成
    '
    ' Prefix   : 上位階層構造
    ' JPParent : 兄弟継続/最終ノード
    '               TREE_CONTINUE = 後続兄弟あり
    '               TREE_LAST_LEVEL = 最終ノード
    ' JPChild  : 子ノード有無
    '               TREE_BAR = 子ノードあり
    '               TREE_SPACE = 子ノードなし
    '------------------------------------
    Dim JPText As String
    Dim JPParent As String
    Dim JPChild As String

    Dim Children As Collection

    Dim ChildProc As Variant
    Dim ChildCount As Long

    Dim TreeText As String
    Dim NextPrefix As String

    Dim Node As clsTreeNode

    Dim i As Long
    Dim IsLastChild As Boolean

    Dim HasChildren As Boolean

    If PathText = "" Then

        PathText = _
            "|" & UCase$(ParentProc) & "|"

    End If

    If MaxDepth > 0 Then

        If Level > MaxDepth Then
            Exit Sub
        End If

    End If

    Set Children = GetChildrenList(ParentProc)

    ChildCount = Children.Count

    If ChildCount = 0 Then Exit Sub

    For i = 1 To ChildCount

        ChildProc = CStr(Children(i))

        '============================
        ' 循環参照防止
        '============================
        If InStr( _
                1, _
                PathText, _
                "|" & UCase$(ChildProc) & "|", _
                vbTextCompare) > 0 Then

            GoTo NextChild

        End If

        IsLastChild = (i = ChildCount)

        '---------------------------
        ' 英語Tree行生成
        '---------------------------
        If IsLastChild Then

            TreeText = _
                Prefix & _
                TREE_NODE_LAST & _
                TREE_SPACE & _
                ChildProc

            NextPrefix = _
                Prefix & _
                TREE_LAST_LEVEL & _
                TREE_SPACE

        Else

            TreeText = _
                Prefix & _
                TREE_NODE_NEXT & _
                TREE_SPACE & _
                ChildProc

            NextPrefix = _
                Prefix & _
                TREE_NEXT_LEVEL & _
                TREE_SPACE

        End If

        Set Node = _
            CreateTreeNode( _
                ChildProc, _
                ParentProc, _
                Level, _
                TreeText)

        If Len(Node.ModuleName) > 0 Then

            If gShowModule Then

                Node.TreeText = _
                    Node.TreeText & _
                    " <" & _
                    Node.ModuleName & _
                    ">"

            End If

        End If

        HasChildren = _
            (GetChildrenList(ChildProc).Count > 0)

        '---------------------------
        ' 日本語Tree行生成
        '---------------------------
        If IsLastChild Then
            JPParent = TREE_LAST_LEVEL
        Else
            JPParent = TREE_CONTINUE
        End If

        If HasChildren Then
            JPChild = TREE_BAR
        Else
            JPChild = TREE_SPACE
        End If

        JPText = _
            Prefix & _
            JPParent & _
            TREE_SPACE & _
            JPChild & _
            Node.JPName

        Node.JPLine = JPText

        gTreeNodes.Add Node

        BuildTreeText _
            ChildProc, _
            NextPrefix, _
            Level + 1, _
            MaxDepth, _
            PathText & _
            UCase$(ChildProc) & "|"

NextChild:

    Next i

End Sub

'==================================================
' Treeシート出力
'==================================================

' @JPName
' ツリー一覧出力
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' PathChartTree(Worksheet)
'
' @Summary
' TreeNode一覧をPathChartTreeシートへ出力する
'
' @Remarks
' 英語Tree行と日本語Tree行を
' 交互に出力する
'
Public Sub OutputTreeNodeList()

    Dim ws As Worksheet
    Dim Node As clsTreeNode

    Dim RowNo As Long

    Dim LastRow As Long
    Dim LastCol As Long

    On Error Resume Next

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PATH_CHART_TREE)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                    gContext.ManagementBook.Worksheets.Count))

        ws.Name = SHEET_PATH_CHART_TREE

    End If

    ws.Cells.Clear

    '------------------------------------
    ' ヘッダ
    '------------------------------------
    ws.Cells(1, 1).Value = "Tree"
    ws.Cells(1, 2).Value = "Level"
    ws.Cells(1, 3).Value = "Procedure"
    ws.Cells(1, 4).Value = "JPName"
    ws.Cells(1, 5).Value = "ParentProcedure"
    ws.Cells(1, 6).Value = "Module"
    ws.Cells(1, 7).Value = "Category"

    RowNo = 2

    '------------------------------------
    ' データ出力
    '------------------------------------
    For Each Node In gTreeNodes

        '---------------------------
        ' 英語行
        '---------------------------
        ws.Cells(RowNo, 1).Value = _
            Node.TreeText

        ws.Cells(RowNo, 2).Value = _
            Node.Level

        ws.Cells(RowNo, 3).Value = _
            Node.ProcName

        ws.Cells(RowNo, 4).Value = _
            Node.JPName

        ws.Cells(RowNo, 5).Value = _
            Node.ParentProcName

        ws.Cells(RowNo, 6).Value = _
            Node.ModuleName

        ws.Cells(RowNo, 7).Value = _
            Node.Category

        RowNo = RowNo + 1

        '---------------------------
        ' 日本語行
        '---------------------------
        If Trim$(Node.JPLine) <> "" Then

            ws.Cells(RowNo, 1).Value = _
                Node.JPLine

            ws.Cells(RowNo, 1).Font.Color = _
                TREE_JP_FONT_COLOR

            RowNo = RowNo + 1

        End If

    Next Node

    '------------------------------------
    ' 体裁
    '------------------------------------
    ws.Rows(1).Font.Bold = True

    ws.Columns("A:G").AutoFit

    ws.Columns("A").ColumnWidth = TREE_COLUMN_WIDTH

    ws.Cells.Font.Name = _
        TREE_FONT_NAME

    '------------------------------------
    ' オートフィルタ
    '------------------------------------
    LastRow = ws.Cells( _
                ws.Rows.Count, _
                1).End(xlUp).row

    LastCol = ws.Cells( _
                1, _
                ws.Columns.Count).End(xlToLeft).Column

    ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(LastRow, LastCol)).AutoFilter

End Sub


