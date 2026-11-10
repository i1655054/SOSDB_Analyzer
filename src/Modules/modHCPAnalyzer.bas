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
    
    '# Procedure一覧取得
    Set FuncList = BuildFuncList()
    
    Trace "HCP", "Step=20 GetProcedureCode"
    
    '# Procedureソース取得
    CodeText = GetProcedureCode( _
                    ModuleName, _
                    ProcName)
    
    Trace "HCP", _
        "CodeLength=" & Len(CodeText)

    Trace "HCP", "Step=30 AnalyzeProcedure"
    
    '# Procedure解析
    Set ProcNodes = AnalyzeProcedure( _
                        CodeText, _
                        ModuleName, _
                        ProcName, _
                        FuncList)

    ' 旧シート削除
    Trace "HCP", "Step=40 DeleteSheetIfExists SHEET_PROC_NODES"
    
    '# HCP関連シート初期化
    DeleteSheetIfExists SHEET_PROC_NODES
    
    Trace "HCP", "Step=50 DeleteSheetIfExists SHEET_PROC_CHART"
    
    DeleteSheetIfExists SHEET_PROC_CHART
    
    Trace "HCP", "Step=60 DeleteSheetIfExists SHEET_HCP_NODES"
    
    DeleteSheetIfExists SHEET_HCP_NODES

    Trace "HCP", "Step=70 ConvertProcToHCP"
    
    '# ProcNodeからHCPNodeへ変換
    Set HCPNodes = ConvertProcToHCP( _
                        ProcNodes)

    '# デバッグ情報出力
    If DEV_MODE Then

        ExportProcNodes ProcNodes

        ExportHCPNodes HCPNodes

    End If

    Trace "HCP", "Step=80 ExportHCPText"
    
    '# HCPチャート出力
    ExportHCPText _
        HCPNodes, _
        ModuleName, _
        ProcName

    Trace "HCP", "Step=90 ExportHCPTextFileFromSheet"
    
    '# HCPテキスト出力
    ExportHCPTextFileFromSheet _
        ModuleName, _
        ProcName
    
    Application.ScreenUpdating = True
    
    Exit Sub

ExitProc:

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    Trace "HCP", "CreateHCP " & Err.Description
    
    '# HCP生成エラー内容通知
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

    '# HCP変換管理初期化
    Set HCPNodes = New Collection

    Set MergeStack = New Collection

    '# ProcNode-HCPNode変換表作成
    Set IDMap = _
        CreateObject("Scripting.Dictionary")

    NewID = 1

    '--------------------------------
    ' Proc → HCP変換
    '--------------------------------
    '# ProcNode走査開始
    For Each ProcNode In ProcNodes

        Trace "PROC", _
            ProcNode.NodeType & _
            ":" & _
            ProcNode.SourceLine

        '# 制御構造終了処理
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

        '# HCPNode生成
        Set HCPNode = New clsHCPNode

        HCPNode.NodeID = NewID

        IDMap.Add _
            CStr(ProcNode.NodeID), _
            NewID

        NewID = NewID + 1

        '# ProcNode属性引継ぎ
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
        
        '# コメント情報引継ぎ
        HCPNode.CommentText = _
            ProcNode.CommentText

        '# ELSEノード変換
        If ProcNode.NodeType = NODE_ELSE Then

            HCPNode.Text = _
                "FALSE"

            HCPNode.NodeType = _
                NODE_FALSE

        Else

            HCPNode.Text = _
                ProcNode.NodeText

            '# CALLノード業務説明取得
            HCPNode.BusinessText = _
                GetBusinessText(ProcNode)

        End If
        
        '-----------------------------
        ' Debug.Print除外
        '-----------------------------
        '# Debug.Print除外
        If gHCPCompactMode Then

            If IsDebugPrint( _
                    HCPNode.Text) Then

                GoTo ContinueLoop

            End If

        End If
        
        '-----------------------------
        ' 重複CALL除外
        '-----------------------------
        '# 重複CALL除外
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

            '# Tag取得処理集約
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

        '# HCPNode登録
        HCPNodes.Add HCPNode

        '# 制御構造開始登録
        If IsMergeNode(HCPNode.NodeType) Then

            PushMergeStack _
                MergeStack, _
                HCPNode.NodeType, _
                HCPNode.NodeID

        End If

        '# TRUEノード生成
        If HCPNode.NodeType = NODE_IF _
        Or HCPNode.NodeType = NODE_ELSEIF Then

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
    '# 制御構造整合性確認
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

    '# ParentID再構築
    RebuildParentID _
        HCPNodes, _
        IDMap

    '--------------------------------
    ' Prev/Next補正
    '--------------------------------
    Trace "HCP", "RebuildLinks"

    '# ノードリンク再構築
    RebuildLinks _
        HCPNodes

    Trace "HCP", "FixTrueBranch"

    '# TRUE枝階層補正
    FixTrueBranch _
        HCPNodes

    '---------------------------------------
    '# DEV_MODE時 ProcNode構造ダンプ
    '---------------------------------------
    If DEV_MODE Then

        DumpNodeRange _
            "After FixTrueBranch", _
            HCPNodes, _
            1, _
            DEBUG_NODE_RANGE

    End If

    Trace "HCP", "FixFalseBranch"

    '# FALSE枝階層補正
    FixFalseBranch _
        HCPNodes

    '---------------------------------------
    '# DEV_MODE時 ProcNode構造ダンプ
    '---------------------------------------
    If DEV_MODE Then

        DumpNodeRange _
            "After FixFalseBranch", _
            HCPNodes, _
            1, _
            DEBUG_NODE_RANGE

    End If

    '--------------------------------
    ' Level再計算
    '--------------------------------
    Trace "HCP", "RebuildHCPLevel"

    '# HCPレベル再計算
    RebuildHCPLevel _
        HCPNodes

    '---------------------------------------
    '# DEV_MODE時 ProcNode構造ダンプ
    '---------------------------------------
    If DEV_MODE Then

        DumpNodes _
            "ConvertProcToHCP End", _
            HCPNodes

    End If
    
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

    '# TRUEノード生成
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

    '# TRUEノード登録
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

    '# マージスタック空判定
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

    '# MERGEノード生成
    AddMergeNode _
        HCPNodes, _
        NewID, _
        GetMergeParentID( _
            MergeStack(MergeStack.Count)), _
        SourceLine

    '# マージスタック削除
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

    '# 制御構造開始位置登録
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

    '# スタック情報分解
    Parts = Split(StackValue, ":")

    If UBound(Parts) <> 1 Then
        GetMergeParentID = 0
        Exit Function
    End If

    '# ParentID返却
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

    '# CALLノード以外除外
    If ProcNode.NodeType <> NODE_CALL Then Exit Function

    '# 呼出先Procedure解決
    Set Info = _
        ResolveProcedure( _
            ProcNode.ModuleName, _
            ProcNode.CallName)

    If Info Is Nothing Then Exit Function

    '# JPName優先表示
    If Trim$(Info.JPName) <> "" Then

        GetBusinessText = _
            Info.JPName & _
            "＜" & _
            Info.ModuleName & _
            "＞"

    Else

        '# Procedure名表示
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

    '# MERGEノード生成
    Set HCPNode = New clsHCPNode

    HCPNode.NodeID = NewID

    NewID = NewID + 1

    HCPNode.ParentID = ParentID

    HCPNode.NodeType = NODE_MERGE

    HCPNode.Text = Text
    
    '# MERGEノード登録
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

    '# ParentID再マッピング
    For Each Node In HCPNodes

        ' MERGEはすでに正しいParentIDを持つ
        If Node.NodeType = NODE_MERGE Then
            GoTo ContinueLoop
        End If

        If Node.ParentID <> 0 Then

            OldParentID = Abs(Node.ParentID)

            '# HCPNode親ID変換
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

    '# ノード前後関係再構築
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
    '# Node検索表生成
    For Each Node In HCPNodes

        NodeMap.Add _
            CStr(Node.NodeID), _
            Node

    Next Node

    '-------------------------
    ' Level再計算
    '-------------------------
    '# HCPレベル再計算
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

    '# 親ノード階層探索
    Do While ParentID <> 0

        Level = Level + 1

        SafeCount = SafeCount + 1

        '# 無限ループ防止
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

    '# TRUE枝補正開始
    For i = 2 To HCPNodes.Count

        If HCPNodes(i).NodeType = NODE_TRUE Then

            IfID = HCPNodes(i - 1).NodeID
            TrueID = HCPNodes(i).NodeID

            '# TRUEノード親設定
            HCPNodes(i).ParentID = IfID

            For j = i + 1 To HCPNodes.Count

                If HCPNodes(j).NodeType = NODE_TRUE Then
                
                    GoTo ContinueJ
                
                End If

                If HCPNodes(j).NodeType = NODE_ELSEIF Then

                    HCPNodes(j).ParentID = IfID

                    HCPNodes(j).BranchType = BRANCH_NONE
                    
                    Exit For

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
                

                '# TRUE配下親補正
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

    '# FALSE枝補正開始
    For i = 1 To HCPNodes.Count

        Set Node = HCPNodes(i)
        
        '# FALSEノード検出
        If Node.NodeType = NODE_FALSE Then

            FalseNodeID = Node.NodeID

        '# FALSE配下親補正
        ElseIf FalseNodeID <> 0 Then

            If Node.NodeType = NODE_ELSEIF Then

                FalseNodeID = 0

            ElseIf Node.NodeType = NODE_MERGE Then

                FalseNodeID = 0
                
            ElseIf Node.BranchType = BRANCH_FALSE Then

                Node.ParentID = FalseNodeID

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

    '# HCPNodesシート取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_NODES)
    
    On Error GoTo 0

    '# HCPNodesシート作成
    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                gContext.ManagementBook.Worksheets.Count))

        ws.Name = SHEET_HCP_NODES

    End If

    '# HCPNodesシート初期化
    ws.Cells.Clear

    '# HCPNodesヘッダ出力
    ws.Range("A1:K1").Value = _
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
            "BusinessText", _
            "CommentText")

    RowOut = 2

    '# HCPノード一覧出力
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
        
        ws.Cells(RowOut, 11) = Node.CommentText

        RowOut = RowOut + 1

    Next Node
    
    '# HCPNodesシート体裁設定
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

    '# 子ノード一覧初期化
    Set Result = New Collection

   '# TRUEノード収集
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType = NODE_TRUE Then

                Result.Add Node

            End If

        End If

    Next Node

    '# FALSEノード収集
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType = NODE_FALSE Then

                Result.Add Node

            End If

        End If

    Next Node

    '# 通常ノード収集
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then

            If Node.NodeType <> NODE_TRUE _
            And Node.NodeType <> NODE_FALSE Then

                Result.Add Node

            End If

        End If

    Next Node

    '# 子ノード一覧返却
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

    '# ノード表示位置設定
    Node.PosX = _
        20 + Level * 220

    Node.PosY = _
        CurrentY

    CurrentY = _
        CurrentY + 60

    '# 子ノード取得
    Set Children = _
        GetChildNodes( _
            Node.NodeID, _
            HCPNodes)

    '# 子ノード配置
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

    '# HCP記号判定
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

    '# HCPチャートシート取得
    Set ws = gContext.ManagementBook.Worksheets( _
                SHEET_HCP_CHART)
    
    On Error GoTo 0

    '# Procedure情報取得
    Set ProcInfo = _
            ResolveProcedure( _
                ModuleName, _
                ProcName)
    
    '# HCPチャートシート作成
    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add( _
                After:=gContext.ManagementBook.Worksheets( _
                gContext.ManagementBook.Worksheets.Count))

        ws.Name = SHEET_HCP_CHART

    End If

    '---------------------------------------
    '# DEV_MODE時 ProcNode構造ダンプ
    '---------------------------------------
    If DEV_MODE Then

        DumpNodes _
            "ExportHCPText Input", _
            HCPNodes

    End If
    
    StartText = ProcName

    '# 開始ノード検索
    For Each Node In HCPNodes

        If Node.NodeType = NODE_START Then

            StartText = Node.Text
            
            Exit For

        End If

    Next Node

    '# HCPチャート初期化
    ws.Cells.Clear

    '# Procedure情報出力
    ws.Cells(1, 1).Value = "Procedure: " & _
        ModuleName & "." & ProcName & _
        "  [ " & ProcInfo.JPName & "＜" & ProcInfo.ModuleName & "＞ ]"
    
    '# HCPヘッダ出力
    ws.Cells(2, 1).Value = "HCP"
    ws.Cells(2, 2).Value = "Business"

    RowOut = 3

    '# 開始ノード出力
    ws.Cells(RowOut, 1).Value = "┬"
    RowOut = RowOut + 1

    ws.Cells(RowOut, 1).Value = _
        "〇│" & StartText
    
    '# 開始Procedure業務説明出力
    If Not ProcInfo Is Nothing Then
        'Debug.Print "Stert:" & ProcInfo.JPName

        ws.Cells(RowOut, 2).Value = _
            ProcInfo.JPName & "＜" & ProcInfo.ModuleName & "＞"
    End If
    
    RowOut = RowOut + 1

    ws.Cells(RowOut, 1).Value = "┴└┐"
    RowOut = RowOut + 1

    '# HCP描画開始
    RenderNode _
        HCPNodes, _
        0, _
        "　　", _
        ws, _
        RowOut

    '# HCPチャート体裁設定
    ws.Columns("A:B").AutoFit
    
    ws.Columns("A").ColumnWidth = HCP_COL_WIDTH
    
    ws.Columns("B").ColumnWidth = BUSINESS_COL_WIDTH
    
    ws.Columns("A").WrapText = True
    
    gContext.ManagementBook.Activate
    
    '# HCPチャート表示
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

    '# 指定階層ノード描画
    For Each Node In HCPNodes

        If Node.ParentID = ParentID Then
        
            '# ノード種別毎に描画
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

                Case NODE_MERGE

                    ' 描画しない

                Case Else
                    
                    RenderControlNode _
                        HCPNodes, _
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

    ' ELSEIFノード
    Dim ElseIfNode As clsHCPNode

    ' MERGEノード
    Dim MergeNode As clsHCPNode

    '-----------------------------
    '# IF見出し
    '-----------------------------
    OutputComment _
        IfNode, _
        Prefix, _
        ws, _
        RowOut

    '# IF条件出力
    If IfNode.NodeType = NODE_ELSEIF Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & _
            "├┐" & _
            GetDisplayText(IfNode, Prefix)

    Else

        ws.Cells(RowOut, 1).Value = _
            Prefix & _
            "◆┐" & _
            GetDisplayText(IfNode, Prefix)

    End If

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '-----------------------------
    '# 配下ノード検索
    '-----------------------------
    For Each Node In HCPNodes

        If Node.ParentID = IfNode.NodeID Then
        
            Select Case Node.NodeType

                Case NODE_TRUE

                    Set TrueNode = Node

                Case NODE_FALSE

                    Set FalseNode = Node

                Case NODE_ELSEIF

                    Set ElseIfNode = Node

                Case NODE_MERGE

                    Set MergeNode = Node

            End Select

        End If

    Next Node

    '-----------------------------
    '# TRUE枝描画
    '-----------------------------
    If Not TrueNode Is Nothing Then

        RenderNode _
            HCPNodes, _
            TrueNode.NodeID, _
            Prefix & "│", _
            ws, _
            RowOut

    End If

    '-----------------------------
    '# IF直下ノード描画
    '-----------------------------
    For Each Node In HCPNodes

        If Node.ParentID = IfNode.NodeID Then

            Select Case Node.NodeType

                Case NODE_TRUE
                Case NODE_FALSE
                Case NODE_ELSEIF
                Case NODE_MERGE

                Case Else

                    RenderControlNode _
                        HCPNodes, _
                        Node, _
                        Prefix & "│", _
                        ws, _
                        RowOut

            End Select

        End If

    Next Node

    '-----------------------------
    '# ELSEIF描画
    '-----------------------------
    If Not ElseIfNode Is Nothing Then

        RenderIfBlock _
            HCPNodes, _
            ElseIfNode, _
            Prefix, _
            ws, _
            RowOut

    End If

    '-----------------------------
    '# FALSE枝描画
    '-----------------------------
    If Not FalseNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "├┐ELSE"

        ws.Cells(RowOut, 1).WrapText = True

        RowOut = RowOut + 1

        RenderNode _
            HCPNodes, _
            FalseNode.NodeID, _
            Prefix & "│", _
            ws, _
            RowOut

    End If

    '-----------------------------
    '# End If描画
    '-----------------------------
    If IfNode.NodeType = NODE_IF Then

        If Not MergeNode Is Nothing Then

            ws.Cells(RowOut, 1).Value = _
                Prefix & _
                "│┴End If"
                '"│┴" & _
                'GetDisplayText(MergeNode, Prefix)

            ws.Cells(RowOut, 1).WrapText = True

            RowOut = RowOut + 1

        End If

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

    OutputComment _
        ForNode, _
        Prefix, _
        ws, _
        RowOut
    
    
    '# FOR開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(ForNode, Prefix)
    
    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '# FOR配下ノード描画
    For Each Node In HCPNodes

        If Node.ParentID = ForNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set MergeNode = Node

            Else

                RenderControlNode _
                    HCPNodes, _
                    Node, _
                    Prefix & "│", _
                    ws, _
                    RowOut
                
            End If

        End If

    Next Node

    '# FOR終端描画
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

    OutputComment _
        SelectNode, _
        Prefix, _
        ws, _
        RowOut
       
    '# SELECT開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "◆┐" & _
        GetDisplayText(SelectNode, Prefix)
    
    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '-----------------------------
    '# CASE数集計
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
    '# CASEブロック描画
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
                    '# CASE配下描画
                    '-----------------------------
                    RenderNode _
                        HCPNodes, _
                        Node.NodeID, _
                        Prefix & "│", _
                        ws, _
                        RowOut

            End Select

        End If

    Next Node

    '-----------------------------
    '# End Select描画
    '-----------------------------
    If Not MergeNode Is Nothing Then
    
        ws.Cells(RowOut, 1).Value = _
            Prefix & "┴" & _
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

    OutputComment _
        DoNode, _
        Prefix, _
        ws, _
        RowOut
    
    '# DO開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(DoNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    
    '# 終端MERGEノード検索
    For Each Node In HCPNodes

        If Node.ParentID = DoNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    '# DO配下ノード描画
    RenderNode _
        HCPNodes, _
        DoNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    '# DO終端出力
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

    OutputComment _
        WithNode, _
        Prefix, _
        ws, _
        RowOut
    
    '# WITH開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(WithNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '# 終端MERGEノード検索
    For Each Node In HCPNodes

        If Node.ParentID = WithNode.NodeID Then
        
            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    '# WITH配下ノード描画
    RenderNode _
        HCPNodes, _
        WithNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    '# WITH終端出力
    If Not EndNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
                GetDisplayText(EndNode, Prefix)
        
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

    OutputComment _
        WhileNode, _
        Prefix, _
        ws, _
        RowOut
    
    '# WHILE開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＠┐" & _
        GetDisplayText(WhileNode, Prefix)

    ws.Cells(RowOut, 1).WrapText = True

    RowOut = RowOut + 1

    '# 終端MERGEノード検索
    For Each Node In HCPNodes

        If Node.ParentID = WhileNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set EndNode = Node

            End If

        End If

    Next Node

    '# WHILE配下ノード描画
    RenderNode _
        HCPNodes, _
        WhileNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    '# WHILE終端出力
    If Not EndNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
                GetDisplayText(EndNode, Prefix)

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

    OutputComment _
        CompileNode, _
        Prefix, _
        ws, _
        RowOut
    
    '# 条件コンパイル開始出力
    ws.Cells(RowOut, 1).Value = _
        Prefix & _
        "＃┐" & _
        GetDisplayText(CompileNode, Prefix)

    RowOut = RowOut + 1

    '# 条件コンパイル配下ノード描画
    RenderNode _
        HCPNodes, _
        CompileNode.NodeID, _
        Prefix & "│", _
        ws, _
        RowOut

    '# 終端MERGEノード検索
    For Each Node In HCPNodes

        If Node.ParentID = CompileNode.NodeID Then

            If Node.NodeType = NODE_MERGE Then

                Set MergeNode = Node
                Exit For

            End If

        End If

    Next Node

    '# 条件コンパイル終端出力
    If Not MergeNode Is Nothing Then

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│┴" & _
            GetDisplayText(MergeNode, Prefix)

        RowOut = RowOut + 1

    End If

End Sub

' @JPName
' 制御ノード描画
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' Node(clsHCPNode)
' Prefix(String)
' ws(Worksheet)
' RowOut(Long)
'
' @Output
' RowOut(Long)
'
' @Summary
' 制御構造ノードを種別に応じて
' 適切な描画ルーチンへ振り分ける
'
' @Remarks
' IF、FOR、DO、WITH、WHILE、
' SELECT、条件コンパイルなどの
' 制御構造を判定し
' 専用描画処理へ委譲する
' 通常ノードはOutputHCPNodeへ
' 出力を委譲する
'
Private Sub RenderControlNode( _
                ByVal HCPNodes As Collection, _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    '# HCP描画ルーティング
    Select Case Node.NodeType

        Case NODE_IF

            '# IFブロック描画
            RenderIfBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_ELSEIF

            RenderIfBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut
        
        Case NODE_FOR

            '# FORブロック描画
            RenderForBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_DO

            '# DOブロック描画
            RenderDoBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_WITH

            '# WITHブロック描画
            RenderWithBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_WHILE

            '# WHILEブロック描画
            RenderWhileBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_SELECT

            '# SELECTブロック描画
            RenderSelectBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case NODE_COMPILE_IF

            '# 条件コンパイルブロック描画
            RenderCompileBlock _
                HCPNodes, _
                Node, _
                Prefix, _
                ws, _
                RowOut

        Case Else

            '# 通常ノード描画
            OutputHCPNode _
                Node, _
                Prefix, _
                ws, _
                RowOut

    End Select

End Sub

' @JPName
' HCPコメント出力
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
' ノードコメントを
' HCPチャートへ出力する
'
' @Remarks
' CommentTextを改行単位で分割し
' 「│#」形式のコメントとして
' HCPチャートへ出力する
'
' 出力例
' │# ProcListシート準備
' │# ProcListヘッダ出力
' │# コメントタグからJPName取得
'
Private Sub OutputComment( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim V As Variant
    
    '# コメント有無確認
    If Trim$(Node.CommentText) = "" Then Exit Sub

    '# HCPコメント出力
    For Each V In Split(Node.CommentText, vbCrLf)

        ws.Cells(RowOut, 1).Value = _
            Prefix & "│#  " & _
            Trim$(CStr(V))
            
        RowOut = RowOut + 1

    Next V

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

    '# ノードコメント出力
    OutputComment _
        Node, _
        Prefix, _
        ws, _
        RowOut
   
    '# HCP表示文字列生成
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

    '# HCPノード出力
    ws.Cells(RowOut, 1).Value = Text

    ws.Cells(RowOut, 1).WrapText = True

    '# 業務説明出力
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

    '# 長文判定
    If Len(Text) <= MAX_LEN Then

        FormatHCPText = Text
        Exit Function

    End If

    '# HCP表示向け改行挿入
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

    '# HCP描画文字列取得
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

    '# 複数行インデント付与
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

    '# HCP表示文字列生成
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

    '# 削除対象シート取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SheetName)

    On Error GoTo 0

    '# シート存在確認
    If ws Is Nothing Then Exit Sub

    '# 削除確認ダイアログ抑止
    Application.DisplayAlerts = False

    '# シート削除
    ws.Delete

    Application.DisplayAlerts = True

End Sub

'
' @JPName
' ノード一覧ダンプ
'
' @Summary
' Collection内の全ノード情報を
' イミディエイトへ出力する
'
Private Sub DumpNodes( _
                ByVal Title As String, _
                ByVal Nodes As Collection)

    Dim Node As Object

    Debug.Print String$(80, "=")
    Debug.Print Title
    Debug.Print String$(80, "=")

    For Each Node In Nodes

        Debug.Print _
            Right$("000" & Node.NodeID, 3), _
            "P=" & Node.ParentID, _
            "L=" & Node.Level, _
            "B=" & Node.BranchType, _
            "[" & Node.NodeType & "]", _
            Node.Text

    Next Node

End Sub

'
' @JPName
' ノード範囲ダンプ
'
' @Summary
' 指定NodeID範囲のみ
' イミディエイトへ出力する
'
Private Sub DumpNodeRange( _
                ByVal Title As String, _
                ByVal Nodes As Collection, _
                ByVal StartID As Long, _
                ByVal EndID As Long)

    Dim Node As Object

    Debug.Print String$(80, "=")
    Debug.Print Title
    Debug.Print String$(80, "=")

    For Each Node In Nodes

        If Node.NodeID >= StartID _
        And Node.NodeID <= EndID Then

            Debug.Print _
                Right$("000" & Node.NodeID, 3), _
                "P=" & Node.ParentID, _
                "L=" & Node.Level, _
                "B=" & Node.BranchType, _
                "[" & Node.NodeType & "]", _
                Node.Text

        End If

    Next Node

End Sub



