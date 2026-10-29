Attribute VB_Name = "modHCPAnalyzer"
Option Explicit
'===================================================================
'@JPName     HCP生成
'@Category   HCP
'@Summary    プロシージャ解析結果からHCP(Hierarchical Control Process)
'            を生成しシートへ出力する
'===================================================================

' HCP表示列幅
Private Const HCP_COL_WIDTH As Double = 120

' 業務説明列幅
Private Const BUSINESS_COL_WIDTH As Double = 40

' @JPName
' HCP生成
'
' @Category
' HCP
'
' @Input
' ModuleName(String)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定プロシージャのHCPを生成し
' HCPシートへ出力する
'
' @Remarks
' プロシージャ解析、HCP変換、
' HCP出力を順次実行する
' DEV_MODE時は中間情報も出力する
'
Public Sub CreateHCP( _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    ' 対象プロシージャコード
    Dim CodeText As String

    ' Proc解析結果
    Dim ProcNodes As Collection
    
    ' HCP変換結果
    Dim HCPNodes As Collection

    ' Procedure一覧
    Dim FuncList As Object
    
    Trace "HCP", _
        ModuleName & "." & ProcName
    
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    
    Trace "HCP", "Step=10 BuildFuncList"
    Set FuncList = BuildFuncList()
    
    Trace "HCP", "Step=20 GetProcedureCode"
    CodeText = GetProcedureCode( _
                    ModuleName, _
                    ProcName)
    
    Trace "HCP", _
        "CodeLength=" & Len(CodeText)

    Trace "HCP", "Step=30 AnalyzeProcedure"
    Set ProcNodes = AnalyzeProcedure( _
                        CodeText, _
                        ModuleName, _
                        ProcName, _
                        FuncList)

    ' 旧シート削除
    Trace "HCP", "Step=40 DeleteSheetIfExists SHEET_PROC_NODES"
    DeleteSheetIfExists SHEET_PROC_NODES
    Trace "HCP", "Step=50 DeleteSheetIfExists SHEET_PROC_CHART"
    DeleteSheetIfExists SHEET_PROC_CHART
    Trace "HCP", "Step=60 DeleteSheetIfExists SHEET_HCP_NODES"
    DeleteSheetIfExists SHEET_HCP_NODES

    Trace "HCP", "Step=70 ConvertProcToHCP"
    Set HCPNodes = ConvertProcToHCP( _
                        ProcNodes)

    If DEV_MODE Then

        ExportProcNodes ProcNodes

        ExportHCPNodes HCPNodes

    End If

    Trace "HCP", "Step=80 ExportHCPText"
    ExportHCPText _
        HCPNodes, _
        ModuleName, _
        ProcName

    Application.ScreenUpdating = True
    
    Exit Sub

ExitProc:

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    Trace "HCP", "CreateHCP " & Err.Description
    
    MsgBox _
        Err.Description, _
        vbExclamation, _
        "HCP生成エラー"

    Resume ExitProc

End Sub

' @JPName
' HCP変換
'
' @Category
' HCP
'
' @Input
' ProcNodes(Collection)
'
' @Output
' HCPNodes(Collection)
'
' @Summary
' プロシージャ解析ノードを
' HCPノードへ変換する
'
' @Remarks
' ProcNodeをHCPNodeへ変換する
' 制御構造に応じてTRUEノードおよび
' MERGEノードを自動生成する
' 変換後にParentID、リンク情報、
' Level情報を再構築する
'
Public Function ConvertProcToHCP( _
                ByVal ProcNodes As Collection) _
                As Collection

    Trace "HCP", "ConvertProcToHCP Start"

    ' HCPノード一覧
    Dim HCPNodes As Collection
    
    ' 制御構造マージ管理スタック
    Dim MergeStack As Collection

    ' Proc解析ノード
    Dim ProcNode As clsProcNode
    
    ' HCPノード
    Dim HCPNode As clsHCPNode

    ' Tag集約用ノード
    Dim LastTagNode As clsHCPNode

    ' 前回出力テキスト
    Dim LastText As String

    ' ProcNodeID→HCPNodeID変換表
    Dim IDMap As Object

    ' 採番用NodeID
    Dim NewID As Long

    Trace "HCP", "CreateHCPNodes"

    Set HCPNodes = New Collection

    Set MergeStack = New Collection

    Set IDMap = _
        CreateObject("Scripting.Dictionary")

    NewID = 1

    '--------------------------------
    ' Proc → HCP変換
    '--------------------------------
    For Each ProcNode In ProcNodes

        Trace "PROC", _
            ProcNode.NodeType & _
            ":" & _
            ProcNode.SourceLine

        Select Case ProcNode.NodeType

            Case NODE_ENDIF, _
                 NODE_ENDSELECT, _
                 NODE_NEXT, _
                 NODE_LOOP, _
                 NODE_ENDWITH, _
                 NODE_WEND, _
                 NODE_COMPILE_ENDIF

                PopMergeStack _
                    MergeStack, _
                    HCPNodes, _
                    NewID, _
                    ProcNode.SourceLine, _
                    ProcNode.NodeType

                GoTo ContinueLoop

        End Select

        Set HCPNode = New clsHCPNode

        HCPNode.NodeID = NewID

        IDMap.Add _
            CStr(ProcNode.NodeID), _
            NewID

        NewID = NewID + 1

        HCPNode.ParentID = _
            ProcNode.ParentID

        HCPNode.PrevNodeID = _
            ProcNode.PrevNodeID

        HCPNode.NextNodeID = _
            ProcNode.NextNodeID

        HCPNode.Level = _
            ProcNode.Level

        HCPNode.NodeType = _
            ProcNode.NodeType

        HCPNode.BranchType = _
            ProcNode.BranchType

        HCPNode.SourceText = _
            ProcNode.NodeText

        HCPNode.SourceLine = _
            ProcNode.SourceLine

        If ProcNode.NodeType = NODE_ELSE Then

            HCPNode.Text = _
                "FALSE"

            HCPNode.NodeType = _
                NODE_FALSE

        Else

            HCPNode.Text = _
                ProcNode.NodeText

            HCPNode.BusinessText = _
                GetBusinessText(ProcNode)

        End If
        
        '-----------------------------
        ' Debug.Print除外
        '-----------------------------
        If gHCPCompactMode Then

            If IsDebugPrint( _
                    HCPNode.Text) Then

                GoTo ContinueLoop

            End If

        End If
        
        '-----------------------------
        ' 重複CALL除外
        '-----------------------------
        If gHCPCompactMode Then

            If HCPNode.NodeType = NODE_CALL Then

                If HCPNode.Text = LastText Then

                    GoTo ContinueLoop

                End If

            End If

        End If

        '-----------------------------
        ' Tag取得集約
        '-----------------------------
        If gHCPCompactMode Then

            If HCPNode.Text = _
                "Procedure情報取得" Then

                If Not LastTagNode Is Nothing Then

                    If LastTagNode.ParentID = _
                        HCPNode.ParentID Then

                        GoTo ContinueLoop

                    End If

                End If

                Set LastTagNode = HCPNode

            Else

                Set LastTagNode = Nothing

            End If

        End If

        HCPNodes.Add HCPNode

        If IsMergeNode(HCPNode.NodeType) Then

            PushMergeStack _
                MergeStack, _
                HCPNode.NodeType, _
                HCPNode.NodeID

        End If

        If HCPNode.NodeType = NODE_IF Then

            AddTrueNode _
                HCPNodes, _
                NewID, _
                HCPNode.NodeID

        End If

        LastText = HCPNode.Text

ContinueLoop:

    Next ProcNode

    '--------------------------------
    ' Stack整合性チェック
    '--------------------------------
    If MergeStack.Count <> 0 Then

        ' スタック残数確認用Index
        Dim i As Long

        Trace "ERROR", _
            "Stack Remains Count=" & _
            MergeStack.Count

        For i = 1 To MergeStack.Count

            Trace "ERROR", _
                "Remain Stack[" & i & "]=" & _
                MergeStack(i)

        Next i

    End If

    '--------------------------------
    ' ParentID補正
    '--------------------------------
    Trace "HCP", "RebuildParentID"

    RebuildParentID _
        HCPNodes, _
        IDMap

    '--------------------------------
    ' Prev/Next補正
    '--------------------------------
    Trace "HCP", "RebuildLinks"

    RebuildLinks _
        HCPNodes

    Trace "HCP", "FixTrueBranch"

    FixTrueBranch _
        HCPNodes

    Trace "HCP", "FixFalseBranch"

    FixFalseBranch _
        HCPNodes

    '--------------------------------
    ' Level再計算
    '--------------------------------
    Trace "HCP", "RebuildHCPLevel"

    RebuildHCPLevel _
        HCPNodes

    Trace "HCP", _
        "ConvertProcToHCP End Nodes=" & _
        HCPNodes.Count & _
        " Stack=" & _
        MergeStack.Count
        
    Set ConvertProcToHCP = _
        HCPNodes

End Function

' @JPName
' Debug.Print判定
'
' @Category
' HCP
'
' @Input
' Text(String)
'
' @Output
' True : Debug.Print文
' False: Debug.Print文以外
'
' @Summary
' Debug.Print文か判定する
'
' @Remarks
' HCP圧縮モード時の不要ノード除外で使用する
' 文字列内に「Debug.Print」が含まれる場合は
' Trueを返す
'
Private Function IsDebugPrint( _
                ByVal Text As String) _
                As Boolean

    IsDebugPrint = _
        (InStr(1, _
               Text, _
               "Debug.Print", _
               vbTextCompare) > 0)

End Function

' @JPName
' TRUEノード追加
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' NewID(Long)
' IfNodeID(Long)
'
' @Output
' HCPNodes
'
' @Summary
' IF文用のTRUEノードを追加する
'
' @Remarks
' THENブランチ開始位置を
' 表現するノードを生成する
'
Private Sub AddTrueNode( _
                ByVal HCPNodes As Collection, _
                ByRef NewID As Long, _
                ByVal IfNodeID As Long)

    ' TRUEノード
    Dim TrueNode As clsHCPNode

    Set TrueNode = New clsHCPNode

    TrueNode.NodeID = NewID

    NewID = NewID + 1

    TrueNode.ParentID = _
        -IfNodeID

    TrueNode.NodeType = _
        NODE_TRUE

    TrueNode.BranchType = _
        BRANCH_TRUE

    TrueNode.Text = _
        "TRUE"

    HCPNodes.Add TrueNode

End Sub

' @JPName
' マージ対象判定
'
' @Category
' HCP
'
' @Input
' NodeType(String)
'
' @Output
' True : マージ対象
' False: マージ対象外
'
' @Summary
' ノード種別がマージ対象か判定する
'
' @Remarks
' IF、FOR、SELECT、DO、WITH、
' WHILE、条件コンパイル開始ノードを
' マージノード生成対象として判定する
'
Private Function IsMergeNode( _
                ByVal NodeType As String) _
                As Boolean

    Select Case NodeType

        Case NODE_IF, _
             NODE_FOR, _
             NODE_SELECT, _
             NODE_DO, _
             NODE_WITH, _
             NODE_WHILE, _
             NODE_COMPILE_IF

            IsMergeNode = True

    End Select

End Function

' @JPName
' マージスタック取得
'
' @Category
' HCP
'
' @Input
' MergeStack(Collection)
' HCPNodes(Collection)
' NewID(Long)
' SourceLine(String)
' NodeType(String)
'
' @Output
' HCPNodes
'
' @Summary
' マージスタックから情報を取得し
' MERGEノードを生成する
'
' @Remarks
' 制御構造終了時に対応する
' MERGEノードを生成する
' スタック不整合時は
' エラートレースを出力する
'
Private Sub PopMergeStack( _
                ByVal MergeStack As Collection, _
                ByVal HCPNodes As Collection, _
                ByRef NewID As Long, _
                ByVal SourceLine As String, _
                ByVal NodeType As String)

    If MergeStack.Count = 0 Then

        Trace "ERROR", _
            NodeType & _
            " Stack Empty Line=" & _
            SourceLine

        Exit Sub

    End If

    TraceStack _
        "POP", _
        MergeStack(MergeStack.Count), _
        MergeStack.Count

    AddMergeNode _
        HCPNodes, _
        NewID, _
        GetMergeParentID( _
            MergeStack(MergeStack.Count)), _
        SourceLine

    MergeStack.Remove _
        MergeStack.Count

End Sub

' @JPName
' マージスタック登録
'
' @Category
' HCP
'
' @Input
' MergeStack(Collection)
' NodeType(String)
' NodeID(Long)
'
' @Output
' MergeStack
'
' @Summary
' 制御構造開始ノードを
' マージスタックへ登録する
'
' @Remarks
' IF、FOR、SELECTなどの
' 制御構造開始位置を
' マージ判定用スタックへ登録する
'
Private Sub PushMergeStack( _
                ByVal MergeStack As Collection, _
                ByVal NodeType As String, _
                ByVal NodeID As Long)

    MergeStack.Add _
        NodeType & ":" & NodeID

    TraceStack _
        "PUSH", _
        NodeType & ":" & NodeID, _
        MergeStack.Count

End Sub

' @JPName
' マージ親ID取得
'
' @Category
' HCP
'
' @Input
' StackValue(String)
'
' @Output
' ParentID(Long)
'
' @Summary
' スタック情報から
' 親ノードIDを取得する
'
' @Remarks
' スタック格納形式
' 「NodeType:NodeID」から
' NodeID部分を取得する
'
Private Function GetMergeParentID( _
                ByVal StackValue As String) _
                As Long

    ' 分解文字列
    Dim Parts() As String

    Parts = Split(StackValue, ":")

    If UBound(Parts) <> 1 Then
        GetMergeParentID = 0
        Exit Function
    End If

    GetMergeParentID = CLng(Parts(1))

End Function

' @JPName
' 業務説明取得
'
' @Category
' HCP
'
' @Input
' ProcNode(clsProcNode)
'
' @Output
' BusinessText(String)
'
' @Summary
' CALLノードに対応する
' 業務説明文を取得する
'
' @Remarks
' CALLノードのみ対象とする
' JPNameが存在する場合は
' 「JPName＜ModuleName＞」形式で返す
' JPName未設定時は
' CallNameを代用する
'
Private Function GetBusinessText( _
                ByVal ProcNode As clsProcNode) _
                As String

    ' 呼出先Procedure情報
    Dim Info As clsProcInfo

    If ProcNode.NodeType <> NODE_CALL Then Exit Function

    Set Info = _
        ResolveProcedure( _
            ProcNode.ModuleName, _
            ProcNode.CallName)

    If Info Is Nothing Then Exit Function

    If Trim$(Info.JPName) <> "" Then

        GetBusinessText = _
            Info.JPName & _
            "＜" & _
            Info.ModuleName & _
            "＞"

    Else

        GetBusinessText = _
            ProcNode.CallName & _
            "＜" & _
            Info.ModuleName & _
            "＞"

    End If

End Function

' @JPName
' マージノード追加
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' NewID(Long)
' ParentID(Long)
' Text(String)
'
' @Output
' HCPNodes(Collection)
' NewID(Long)
'
' @Summary
' マージノードを生成し追加する
'
' @Remarks
' 制御構造終了位置を表す
' MERGEノードを生成する
' ノードID採番後にコレクションへ登録する
'
Private Sub AddMergeNode( _
                ByVal HCPNodes As Collection, _
                ByRef NewID As Long, _
                ByVal ParentID As Long, _
                ByVal Text As String)

    ' MERGEノード
    Dim HCPNode As clsHCPNode

    Set HCPNode = New clsHCPNode

    HCPNode.NodeID = NewID

    NewID = NewID + 1

    HCPNode.ParentID = ParentID

    HCPNode.NodeType = NODE_MERGE

    HCPNode.Text = Text

    HCPNodes.Add HCPNode

End Sub

' @JPName
' 親子関係再構築
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' IDMap(Dictionary)
'
' @Output
' HCPNodes
'
' @Summary
' ParentIDをHCPノードIDへ
' 再マッピングする
'
' @Remarks
' ProcNode時点のParentIDを
' HCPNode採番後のNodeIDへ変換する
' MERGEノードは補正対象外とする
'
Private Sub RebuildParentID( _
                ByVal HCPNodes As Collection, _
                ByVal IDMap As Object)

    ' HCPノード
    Dim Node As clsHCPNode
    
    ' ProcNode時点の親ID
    Dim OldParentID As Long

    For Each Node In HCPNodes

        ' MERGEはすでに正しいParentIDを持つ
        If Node.NodeType = NODE_MERGE Then
            GoTo ContinueLoop
        End If

        If Node.ParentID <> 0 Then

            OldParentID = Abs(Node.ParentID)

            If IDMap.Exists(CStr(OldParentID)) Then

                Node.ParentID = _
                    CLng(IDMap(CStr(OldParentID)))

            Else

                Node.ParentID = 0

            End If

        End If

ContinueLoop:

    Next Node

End Sub

' @JPName
' ノードリンク再構築
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' HCPNodes
'
' @Summary
' PrevNodeIDおよびNextNodeIDを
' 再設定する
'
' @Remarks
' HCPNodesの格納順を基準に
' PrevNodeIDおよびNextNodeIDを再設定する
'
Private Sub RebuildLinks( _
                ByVal HCPNodes As Collection)

    ' NodeIndex
    Dim i As Long

    ' HCPノード
    Dim Node As clsHCPNode

    For i = 1 To HCPNodes.Count

        Set Node = HCPNodes(i)

        If i = 1 Then

            Node.PrevNodeID = 0

        Else

            Node.PrevNodeID = _
                HCPNodes(i - 1).NodeID

        End If

        If i = HCPNodes.Count Then

            Node.NextNodeID = 0

        Else

            Node.NextNodeID = _
                HCPNodes(i + 1).NodeID

        End If

    Next i

End Sub

' @JPName
' レベル再計算
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' HCPNodes
'
' @Summary
' 親子関係を基に
' HCPレベルを再計算する
'
' @Remarks
' ParentIDの階層を辿り
' HCP表示用のLevelを再計算する
'
Private Sub RebuildHCPLevel( _
                ByVal HCPNodes As Collection)

    ' NodeID管理Dictionary
    Dim NodeMap As Object

    ' HCPノード
    Dim Node As clsHCPNode

    Set NodeMap = _
        CreateObject("Scripting.Dictionary")

    '-------------------------
    ' Node検索用
    '-------------------------
    For Each Node In HCPNodes

        NodeMap.Add _
            CStr(Node.NodeID), _
            Node

    Next Node

    '-------------------------
    ' Level再計算
    '-------------------------
    For Each Node In HCPNodes

        Node.Level = _
            CalcLevel(Node, NodeMap)

    Next Node

End Sub

' @JPName
' レベル算出
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
' NodeMap(Dictionary)
'
' @Output
' Level(Long)
'
' @Summary
' ノード階層レベルを算出する
'
' @Remarks
' ParentIDを再帰的に辿り
' 階層深度を算出する
' 無限ループ防止のため
' 100階層で探索を打ち切る
'
Private Function CalcLevel( _
                ByVal Node As clsHCPNode, _
                ByVal NodeMap As Object) _
                As Long

    ' 親NodeID
    Dim ParentID As Long
    
    ' 算出Level
    Dim Level As Long
    
    ' 無限ループ防止カウンタ
    Dim SafeCount As Long

    ParentID = Node.ParentID

    Do While ParentID <> 0

        Level = Level + 1

        SafeCount = SafeCount + 1

        If SafeCount > 100 Then

            Exit Do

        End If

        If NodeMap.Exists(CStr(ParentID)) Then

            ParentID = _
                NodeMap(CStr(ParentID)).ParentID

        Else

            Exit Do

        End If

    Loop

    CalcLevel = Level

End Function

' @JPName
' TRUE枝補正
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' HCPNodes
'
' @Summary
' TRUEブランチ配下の
' ParentIDを補正する
'
' @Remarks
' TRUEノード配下の枝に対して
' ParentIDをTRUEノードへ付け替える
' IFブロックの階層構造を補正する
'
Private Sub FixTrueBranch( _
                ByVal HCPNodes As Collection)

    ' HCPNodes走査Index
    Dim i As Long
    
    ' 子ノード走査Index
    Dim j As Long

    ' IFノードID
    Dim IfID As Long
    
    ' TRUEノードID
    Dim TrueID As Long

    For i = 2 To HCPNodes.Count

        If HCPNodes(i).NodeType = NODE_TRUE Then

            IfID = HCPNodes(i - 1).NodeID
            TrueID = HCPNodes(i).NodeID

            HCPNodes(i).ParentID = IfID

            For j = i + 1 To HCPNodes.Count

                If HCPNodes(j).NodeType = NODE_TRUE Then
                    GoTo ContinueJ
                End If

                If HCPNodes(j).NodeType = NODE_FALSE Then
                    Exit For
                End If
                

                If HCPNodes(j).ParentID <> IfID _
                And HCPNodes(j).ParentID <> TrueID Then

                    If HCPNodes(j).Level <= _
                            HCPNodes(i - 1).Level Then

                        Exit For

                    End If

                End If
                

                If HCPNodes(j).BranchType = BRANCH_TRUE Then

                    If HCPNodes(j).ParentID = IfID Then

                        HCPNodes(j).ParentID = TrueID

                    End If

                End If
ContinueJ:
            Next j

        End If

    Next i

End Sub

' @JPName
' FALSE枝補正
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' HCPNodes
'
' @Summary
' FALSEブランチ配下の
' ParentIDを補正する
'
' @Remarks
' FALSEブランチ配下のノードへ
' FALSEノードを親として設定する
' MERGE到達時に補正を終了する
'
Private Sub FixFalseBranch( _
                ByVal HCPNodes As Collection)

    ' HCPNodes走査Index
    Dim i As Long

    ' 対象ノード
    Dim Node As clsHCPNode

    ' FALSEノードID
    Dim FalseNodeID As Long

    FalseNodeID = 0

    For i = 1 To HCPNodes.Count

        Set Node = HCPNodes(i)

        If Node.NodeType = NODE_FALSE Then

            FalseNodeID = Node.NodeID

        ElseIf FalseNodeID <> 0 Then

            If Node.NodeType = NODE_MERGE Then

                FalseNodeID = 0

            ElseIf Node.BranchType = _
                    BRANCH_FALSE Then

                Node.ParentID = _
                    FalseNodeID

            End If

        End If

    Next i

End Sub

' @JPName
' HCPノード出力
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' HCPノードシート
'
' @Summary
' HCPノード一覧をシートへ出力する
'
' @Remarks
' HCPノード情報を一覧形式で出力する
' 主にデバッグおよび解析確認用途で利用する
'
Public Sub ExportHCPNodes( _
                ByVal HCPNodes As Collection)

    ' 出力先シート
    Dim ws As Worksheet

    ' HCPノード
    Dim Node As clsHCPNode

    ' 出力行
    Dim RowOut As Long

    On Error Resume Next

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_NODES)
    
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                gContext.ManagementBook.Worksheets.Count))

        ws.Name = SHEET_HCP_NODES

    End If

    ws.Cells.Clear

    ws.Range("A1:J1").Value = _
        Array( _
            "NodeID", _
            "ParentID", _
            "PrevNodeID", _
            "NextNodeID", _
            "Level", _
            "NodeType", _
            "BranchType", _
            "Text", _
            "SourceText", _
            "BusinessText")

    RowOut = 2

    For Each Node In HCPNodes

        ws.Cells(RowOut, 1) = Node.NodeID
        
        ws.Cells(RowOut, 2) = Node.ParentID
        
        ws.Cells(RowOut, 3) = Node.PrevNodeID
        
        ws.Cells(RowOut, 4) = Node.NextNodeID
        
        ws.Cells(RowOut, 5) = Node.Level
        
        ws.Cells(RowOut, 6) = Node.NodeType
        
        ws.Cells(RowOut, 7) = Node.BranchType
        
        ws.Cells(RowOut, 8) = Node.Text
        
        ws.Cells(RowOut, 9) = Node.SourceText
        
        ws.Cells(RowOut, 10) = Node.BusinessText

        RowOut = RowOut + 1

    Next Node
    
    ws.Columns.AutoFit

End Sub

' @JPName
' 子ノード取得
'
' @Category
' HCP
'
' @Input
' ParentID(Long)
' HCPNodes(Collection)
'
' @Output
' Result(Collection)
'
' @Summary
' 指定ノード配下の
' 子ノード一覧を取得する
'
' @Remarks
' TRUEノード、
' FALSEノード、
' その他ノードの順で返却する
' 描画順序制御に利用する
'
Public Function GetChildNodes( _
                ByVal ParentID As Long, _
                ByVal HCPNodes As Collection) _
                As Collection

    ' 結果ノード一覧
    Dim Result As Collection

    ' HCPノード
    Dim Node As clsHCPNode

    Set Result = New Collection

    ' TRUE
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType = NODE_TRUE Then

                Result.Add Node

            End If

        End If

    Next Node

    ' FALSE
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType = NODE_FALSE Then

                Result.Add Node

            End If

        End If

    Next Node

    ' その他
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType <> NODE_TRUE _
            And Node.NodeType <> NODE_FALSE Then

                Result.Add Node

            End If

        End If

    Next Node

    Set GetChildNodes = Result

End Function

' @JPName
' ノード配置
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
' HCPNodes(Collection)
' Level(Long)
' CurrentY(Double)
'
' @Output
' Node.PosX
' Node.PosY
'
' @Summary
' HCPノードの表示位置を設定する
'
' @Remarks
' 子ノードを再帰的に走査し
' 階層レベルに応じた配置座標を設定する
'
Private Sub LayoutTree( _
                ByVal Node As clsHCPNode, _
                ByVal HCPNodes As Collection, _
                ByVal Level As Long, _
                ByRef CurrentY As Double)

    ' 子ノード一覧
    Dim Children As Collection
    
    ' 子ノード
    Dim Child As clsHCPNode

    Node.PosX = _
        20 + Level * 220

    Node.PosY = _
        CurrentY

    CurrentY = _
        CurrentY + 60

    Set Children = _
        GetChildNodes( _
            Node.NodeID, _
            HCPNodes)

    For Each Child In Children

        LayoutTree _
            Child, _
            HCPNodes, _
            Level + 1, _
            CurrentY

    Next Child

End Sub

' @JPName
' HCP記号取得
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
'
' @Output
' Symbol(String)
'
' @Summary
' ノード種別に対応する
' HCP表示記号を取得する
'
' @Remarks
' IF、CALL、FORなどの
' ノード種別に応じた記号を返却する
'
Private Function GetHCPSymbol( _
                ByVal Node As clsHCPNode) _
                As String

    Select Case Node.NodeType

        Case NODE_START
            GetHCPSymbol = "〇│"

        Case NODE_DECLARE
            GetHCPSymbol = "□ "

        Case NODE_REFERENCE
            GetHCPSymbol = "◇ "

        Case NODE_CONTROL
            GetHCPSymbol = "☆ "

        Case NODE_NORMAL
            GetHCPSymbol = "〇 "

        Case NODE_CALL
            GetHCPSymbol = "◎ "

        Case NODE_IF
            GetHCPSymbol = "◆ "
        
        Case NODE_INLINE_IF
        
            If InStr(1, _
                Node.Text, _
                "Exit ", _
                vbTextCompare) > 0 Then

                GetHCPSymbol = "◆→ "

            Else

                GetHCPSymbol = "◆ "

            End If

        Case NODE_TRUE
            GetHCPSymbol = "THEN "

        Case NODE_FALSE
            GetHCPSymbol = "ELSE "

        Case NODE_FOR
            GetHCPSymbol = "＠ "
        
        Case NODE_DO
            GetHCPSymbol = "＠ "

        Case NODE_WITH
            GetHCPSymbol = "＠ "

        Case NODE_WHILE
            GetHCPSymbol = "＠ "

        Case NODE_SELECT
            GetHCPSymbol = "◆ "

        Case NODE_CASE
            GetHCPSymbol = ""

        Case NODE_MERGE
            GetHCPSymbol = ""

        Case NODE_END
            GetHCPSymbol = "┴"
        
        Case NODE_GOTO
            GetHCPSymbol = "→ "
    
        Case NODE_EXIT
            GetHCPSymbol = "→ "
        
        Case NODE_LABEL
            GetHCPSymbol = "▼ "
        
        Case NODE_RESUME
            GetHCPSymbol = "☆ "

        Case Else
            GetHCPSymbol = ""

    End Select

End Function

' @JPName
' HCP出力
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' ModuleName(String)
' ProcName(String)
'
' @Output
' HCPチャートシート
'
' @Summary
' HCPノードをテキスト形式で
' シートへ出力する
'
' @Remarks
' HCP列と業務説明列を生成し
' プロシージャ階層を可視化する
'
Public Sub ExportHCPText( _
                ByVal HCPNodes As Collection, _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    ' 出力先シート
    Dim ws As Worksheet
    
    ' 出力行
    Dim RowOut As Long
    
    ' HCPノード
    Dim Node As clsHCPNode
    
    ' Procedure情報
    Dim ProcInfo As clsProcInfo
    
    ' 開始表示文字列
    Dim StartText As String

    On Error Resume Next

    Set ws = gContext.ManagementBook.Worksheets( _
                SHEET_HCP_CHART)
    
    On Error GoTo 0

    Set ProcInfo = _
            ResolveProcedure( _
                ModuleName, _
                ProcName)
    
    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                gContext.ManagementBook.Worksheets.Count))

        ws.Name = SHEET_HCP_CHART

    End If

    StartText = ProcName

    For Each Node In HCPNodes

        If Node.NodeType = NODE_START Then

            StartText = Node.Text
            
            Exit For

        End If

    Next Node

    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Procedure: " & _
        ModuleName & "." & ProcName & _
        "  [ " & ProcInfo.JPName & "＜" & ProcInfo.ModuleName & "＞ ]"
    
    ws.Cells(2, 1).Value = "HCP"
    ws.Cells(2, 2).Value = "Business"

    RowOut = 3

    ws.Cells(RowOut, 1).Value = "┬"
    RowOut = RowOut + 1

    ws.Cells(RowOut, 1).Value = _
        "〇│" & StartText
    
    If Not ProcInfo Is Nothing Then
        'Debug.Print "Stert:" & ProcInfo.JPName

        ws.Cells(RowOut, 2).Value = _
            ProcInfo.JPName & "＜" & ProcInfo.ModuleName & "＞"
    End If
    
    RowOut = RowOut + 1

    ws.Cells(RowOut, 1).Value = "┴└┐"
    RowOut = RowOut + 1

    RenderNode _
        HCPNodes, _
        0, _
        "　　", _
        ws, _
        RowOut

    ws.Columns("A:B").AutoFit
    
    ws.Columns("A").ColumnWidth = HCP_COL_WIDTH
    
    ws.Columns("B").ColumnWidth = BUSINESS_COL_WIDTH
    
    ws.Columns("A").WrapText = True
    
    gContext.ManagementBook.Activate
    ws.Activate
    
    ws.Range("B1").Select

    With ActiveWindow

        .FreezePanes = False

        .SplitRow = 0
        .SplitColumn = 1

    End With
    
    ActiveWindow.Zoom = 85
    
End Sub

' @JPName
' HCP描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' ParentID(Long)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' 指定ノード配下を
' HCP形式で描画する
'
' @Remarks
' ノード種別に応じて
' 個別レンダリング処理へ振り分ける
'
Private Sub RenderNode( _
                ByVal HCPNodes As Collection, _
                ByVal ParentID As Long, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' 描画対象ノード
    Dim Node As clsHCPNode

    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            Select Case Node.NodeType

                Case NODE_START

                    ' ExportHCPTextで出力済み

                Case NODE_DECLARE

                    ws.Cells(RowOut, 1).Value = _
                        Prefix & _
                        "□ " & _
                        GetDisplayText(Node, Prefix)

                    ws.Cells(RowOut, 1).WrapText = True

                    RowOut = RowOut + 1

                Case NODE_END

                    ws.Cells(RowOut, 1).Value = _
                        Prefix & _
                        "┴" & Node.Text

                    RowOut = RowOut + 1

                Case NODE_IF

                    RenderIfBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_FOR

                    RenderForBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_SELECT

                    RenderSelectBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_DO

                    RenderDoBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_WITH

                    RenderWithBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_WHILE

                    RenderWhileBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_COMPILE_IF

                    RenderCompileBlock _
                        HCPNodes, _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut

                Case NODE_MERGE

                    ' 描画しない

                Case Else
                    
                    OutputHCPNode _
                        Node, _
                        Prefix, _
                        ws, _
                        RowOut
        
            End Select

        End If

    Next Node

End Sub

' @JPName
' IFブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' IfNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' IF構造をHCP形式で描画する
'
' @Remarks
' TRUEブランチおよびFALSEブランチを
' 再帰的に出力し
' MERGEノードをEnd Ifとして表示する
'
Private Sub RenderIfBlock( _
                ByVal HCPNodes As Collection, _
                ByVal IfNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' IF配下ノード
    Dim Node As clsHCPNode

    ' TRUEノード
    Dim TrueNode As clsHCPNode
    
    ' FALSEノード
    Dim FalseNode As clsHCPNode
    
    ' MERGEノード
    Dim MergeNode As clsHCPNode

    '-----------------------------
    ' IF
    '-----------------------------
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "◆┐" & _
        GetDisplayText(IfNode, Prefix)
    
    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '-----------------------------
    ' TRUE/FALSE/MERGE検索
    '-----------------------------
    For Each Node In HCPNodes

        If Node.ParentID = IfNode.NodeID Then

            Select Case Node.NodeType

                Case NODE_TRUE
                    Set TrueNode = Node

                Case NODE_FALSE
                    Set FalseNode = Node

                Case NODE_MERGE
                    Set MergeNode = Node

            End Select

        End If

    Next Node

    '-----------------------------
    ' THEN
    '-----------------------------
    If Not TrueNode Is Nothing Then

        ' THEN見出し
        'ws.Cells(RowOut, 1).Value = _
        '    Prefix & "││THEN"

        'RowOut = RowOut + 1

        ' THEN本体
        RenderNode _
            HCPNodes, _
            TrueNode.NodeID, _
            Prefix & "│", _
            ws, _
            RowOut

        ' THEN終端
        'If Not FalseNode Is Nothing Then

        '    ws.Cells(RowOut, 1).Value = _
        '        Prefix & "│┴"
            
        '    ws.Cells(RowOut, 1).WrapText = True

        '    RowOut = RowOut + 1

        'End If

    End If

    '-----------------------------
    ' IF直下の通常ノード
    '-----------------------------
    For Each Node In HCPNodes

        If Node.ParentID = IfNode.NodeID Then

            Select Case Node.NodeType

                Case NODE_TRUE
                Case NODE_FALSE
                Case NODE_MERGE

                Case NODE_IF

                    RenderIfBlock _
                        HCPNodes, _
                        Node, _
                        Prefix & "│", _
                        ws, _
                        RowOut

                Case NODE_FOR

                    RenderForBlock _
                        HCPNodes, _
                        Node, _
                        Prefix & "│", _
                        ws, _
                        RowOut

                Case NODE_SELECT

                    RenderSelectBlock _
                        HCPNodes, _
                        Node, _
                        Prefix & "│", _
                        ws, _
                        RowOut

                Case Else

                    OutputHCPNode _
                        Node, _
                        Prefix & "│", _
                        ws, _
                        RowOut

            End Select

        End If

    Next Node

    '-----------------------------
    ' ELSE
    '-----------------------------
    If Not FalseNode Is Nothing Then

        'ELSE見出し
        ws.Cells(RowOut, 1).Value = _
            Prefix & "├┐ELSE"
        
        ws.Cells(RowOut, 1).WrapText = True

        RowOut = RowOut + 1

        ' ELSE本体
        RenderNode _
            HCPNodes, _
            FalseNode.NodeID, _
            Prefix & "│", _
            ws, _
            RowOut

    End If

    '-----------------------------
    ' End If
    '-----------------------------
    If Not MergeNode Is Nothing Then

        ' ELSE終端兼 End If
        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(MergeNode, Prefix)
        
        ws.Cells(RowOut, 1).WrapText = True

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' FORブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' ForNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' FOR構造をHCP形式で描画する
'
' @Remarks
' FOR配下のノードを再帰的に出力し
' MERGEノードを終端として描画する
'
Private Sub RenderForBlock( _
                ByVal HCPNodes As Collection, _
                ByVal ForNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' FOR配下ノード
    Dim Node As clsHCPNode
    
    ' MERGEノード
    Dim MergeNode As clsHCPNode

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(ForNode, Prefix)
    
    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    For Each Node In HCPNodes

        If Node.ParentID = ForNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set MergeNode = Node

            ElseIf Node.NodeType = NODE_IF Then

                RenderIfBlock _
                    HCPNodes, _
                    Node, _
                    Prefix & "│", _
                    ws, _
                    RowOut

            ElseIf Node.NodeType = NODE_FOR Then

                RenderForBlock _
                    HCPNodes, _
                    Node, _
                    Prefix & "│", _
                    ws, _
                    RowOut

            ElseIf Node.NodeType = NODE_SELECT Then

                RenderSelectBlock _
                    HCPNodes, _
                    Node, _
                    Prefix & "│", _
                    ws, _
                    RowOut

            Else

                OutputHCPNode _
                    Node, _
                    Prefix & "│", _
                    ws, _
                    RowOut
                
            End If

        End If

    Next Node

    If Not MergeNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(MergeNode, Prefix)
        
        ws.Cells(RowOut, 1).WrapText = True

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' SELECTブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' SelectNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' Select Case構造をHCP形式で描画する
'
' @Remarks
' CASE単位でノードを展開して出力し
' MERGEノードをEnd Selectとして描画する
'
Private Sub RenderSelectBlock( _
                ByVal HCPNodes As Collection, _
                ByVal SelectNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' SELECT配下ノード
    Dim Node As clsHCPNode
    
    ' MERGEノード
    Dim MergeNode As clsHCPNode

    ' CASE数
    Dim CaseCount As Long
    
    ' 現在CASE番号
    Dim CaseIndex As Long

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "◆┐" & _
        GetDisplayText(SelectNode, Prefix)
    
    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '-----------------------------
    ' CASE数取得
    '-----------------------------
    CaseCount = 0

    For Each Node In HCPNodes

        If Node.ParentID = SelectNode.NodeID Then

            If Node.NodeType = NODE_CASE Then

                CaseCount = CaseCount + 1

            ElseIf Node.NodeType = NODE_MERGE Then

                Set MergeNode = Node

            End If

        End If

    Next Node

    CaseIndex = 0

    '-----------------------------
    ' CASE出力
    '-----------------------------
    For Each Node In HCPNodes

        If Node.ParentID = SelectNode.NodeID Then

            Select Case Node.NodeType

                Case NODE_CASE

                    CaseIndex = CaseIndex + 1

                    '-----------------------------
                    ' CASE見出し
                    '-----------------------------
                    If CaseIndex = 1 Then

                        ws.Cells(RowOut, 1).Value = _
                            Prefix & "││" & _
                            GetDisplayText(Node, Prefix)

                    Else

                        ws.Cells(RowOut, 1).Value = _
                            Prefix & "├┐" & _
                            GetDisplayText(Node, Prefix)

                    End If
                    
                    ws.Cells(RowOut, 1).WrapText = True

                    RowOut = RowOut + 1

                    '-----------------------------
                    ' CASE本体
                    '-----------------------------
                    RenderNode _
                        HCPNodes, _
                        Node.NodeID, _
                        Prefix & "│", _
                        ws, _
                        RowOut

                    '-----------------------------
                    ' CASE終端
                    ' 最終CASE以外
                    '-----------------------------
                    'If CaseIndex < CaseCount Then

                    '    ws.Cells(RowOut, 1).Value = _
                    '        Prefix & "│┴"
                        
                    '    ws.Cells(RowOut, 1).WrapText = True

                    '    RowOut = RowOut + 1

                    'End If

            End Select

        End If

    Next Node

    '-----------------------------
    ' End Select
    '-----------------------------
    If Not MergeNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(MergeNode, Prefix)
        
        ws.Cells(RowOut, 1).WrapText = True

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' DOブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' DoNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' Do Loop構造をHCP形式で描画する
'
' @Remarks
' DO配下のノードを再帰的に出力し
' MERGEノードをループ終端として描画する
'
Private Sub RenderDoBlock( _
                ByVal HCPNodes As Collection, _
                ByVal DoNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' DO配下ノード
    Dim Node As clsHCPNode
    
    ' 終端MERGEノード
    Dim EndNode As clsHCPNode

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(DoNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    For Each Node In HCPNodes

        If Node.ParentID = DoNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    RenderNode _
        HCPNodes, _
        DoNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    If Not EndNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(EndNode, Prefix)

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' WITHブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' WithNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' With構造をHCP形式で描画する
'
' @Remarks
' WITH配下のノードを再帰的に出力し
' MERGEノードをEnd Withとして描画する
'
Private Sub RenderWithBlock( _
                ByVal HCPNodes As Collection, _
                ByVal WithNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' WITH配下ノード
    Dim Node As clsHCPNode
    
    ' 終端MERGEノード
    Dim EndNode As clsHCPNode

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(WithNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    For Each Node In HCPNodes

        If Node.ParentID = WithNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    RenderNode _
        HCPNodes, _
        WithNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    If Not EndNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
                EndNode.Text
        
        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' WHILEブロック描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' WhileNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' While構造をHCP形式で描画する
'
' @Remarks
' WHILE配下のノードを再帰的に出力し
' MERGEノードをループ終端として描画する
'
Private Sub RenderWhileBlock( _
                ByVal HCPNodes As Collection, _
                ByVal WhileNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' WHILE配下ノード
    Dim Node As clsHCPNode
    
    ' 終端MERGEノード
    Dim EndNode As clsHCPNode

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(WhileNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    For Each Node In HCPNodes

        If Node.ParentID = WhileNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    RenderNode _
        HCPNodes, _
        WhileNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    If Not EndNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
                EndNode.Text

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' 条件コンパイル描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' CompileNode(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' 条件コンパイル構造をHCP形式で描画する
'
' @Remarks
' #If～#End If配下のノードを出力し
' MERGEノードを条件コンパイル終端として描画する
'
Private Sub RenderCompileBlock( _
                ByVal HCPNodes As Collection, _
                ByVal CompileNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' 条件コンパイル配下ノード
    Dim Node As clsHCPNode
    
    ' 終端MERGEノード
    Dim MergeNode As clsHCPNode

    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＃┐" & _
        GetDisplayText(CompileNode, Prefix)

    RowOut = RowOut + 1

    RenderNode _
        HCPNodes, _
        CompileNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    For Each Node In HCPNodes

        If Node.ParentID = CompileNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set MergeNode = Node
                Exit For

            End If

        End If

    Next Node

    If Not MergeNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(MergeNode, Prefix)

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' ノード出力
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' 単一ノードをシートへ出力する
'
' @Remarks
' HCP記号と表示文字列を組み合わせて出力し
' 業務説明が存在する場合は併せて出力する
'
Private Sub OutputHCPNode( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    ' 出力文字列
    Dim Text As String

    If Node.NodeType = NODE_LABEL Then

        Text = _
            GetHCPSymbol(Node) & _
            GetDisplayText(Node, Prefix)

    Else

        Text = _
            Prefix & _
            GetHCPSymbol(Node) & _
            GetDisplayText(Node, Prefix)

    End If

    ws.Cells(RowOut, 1).Value = Text

    ws.Cells(RowOut, 1).WrapText = True

    If Trim$(Node.BusinessText) <> "" Then

        ws.Cells(RowOut, 2).Value = _
            Node.BusinessText

    End If

    RowOut = RowOut + 1

End Sub

' @JPName
' HCP文字列整形
'
' @Category
' HCP
'
' @Input
' Text(String)
'
' @Output
' FormattedText(String)
'
' @Summary
' HCP表示用に文字列を整形する
'
' @Remarks
' 長いコードは可読性向上のため
' 特定位置で改行を挿入する
'
Private Function FormatHCPText( _
                ByVal Text As String) _
                As String

    Const MAX_LEN As Long = 60

    If Len(Text) <= MAX_LEN Then

        FormatHCPText = Text
        Exit Function

    End If

    Text = Replace( _
                Text, _
                "Array(", _
                "Array(" & vbCrLf)
    
    Text = Replace( _
                Text, _
                ".Worksheets.Add(", _
                ".Worksheets.Add(" & vbCrLf)

    Text = Replace( _
                Text, _
                ".Worksheets(", _
                ".Worksheets(" & vbCrLf)

    Text = Replace( _
                Text, _
                "(After:=", _
                vbCrLf & "After:=")

    FormatHCPText = Text

End Function

' @JPName
' 描画文字列取得
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
'
' @Output
' RenderText(String)
'
' @Summary
' 描画用文字列を取得する
'
' @Remarks
' ノード文字列に対して
' HCP表示向け整形を行う
'
Private Function GetRenderText( _
                ByVal Node As clsHCPNode) _
                As String

    GetRenderText = _
        FormatHCPText(Node.Text)

End Function

' @JPName
' 複数行インデント付与
'
' @Category
' HCP
'
' @Input
' Text(String)
' Prefix(String)
'
' @Output
' Text(String)
'
' @Summary
' 複数行文字列へインデントを付与する
'
' @Remarks
' 2行目以降にHCP階層表示用の
' プレフィックスを付加する
'
Private Function AddIndentToMultiLine( _
                ByVal Text As String, _
                ByVal Prefix As String) _
                As String

    AddIndentToMultiLine = _
        Replace( _
            Text, _
            vbCrLf, _
            vbCrLf & Prefix & "│ ")

End Function

' @JPName
' 表示文字列取得
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
' Prefix(String)
'
' @Output
' DisplayText(String)
'
' @Summary
' HCP表示用文字列を生成する
'
' @Remarks
' 描画文字列整形および
' 複数行インデント付与を行う
'
Private Function GetDisplayText( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String) _
                As String

    GetDisplayText = _
        AddIndentToMultiLine( _
            GetRenderText(Node), _
            Prefix)

End Function

' @JPName
' シート削除
'
' @Category
' HCP
'
' @Input
' SheetName(String)
'
' @Output
' なし
'
' @Summary
' 指定シートが存在する場合に削除する
'
' @Remarks
' シート未存在時は何も行わない
' 確認ダイアログは表示しない
'
Public Sub DeleteSheetIfExists( _
                ByVal SheetName As String)

    ' 削除対象シート
    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SheetName)

    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Application.DisplayAlerts = False

    ws.Delete

    Application.DisplayAlerts = True

End Sub

