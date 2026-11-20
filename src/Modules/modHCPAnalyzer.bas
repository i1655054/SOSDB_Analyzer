Attribute VB_Name = "modHCPAnalyzer"
Option Explicit
'===================================================================
'@JPName     HCP生成エンジン
'@Category   HCP
'@Summary    プロシージャ解析結果からHCP(Hierarchical Control Process)
'            を生成しシートへ出力する
'===================================================================

'--------------------------------------------------
' HCP生成
'--------------------------------------------------

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

    ' RenderBuilder
    Dim RenderBuilder As clsHCPRenderBuilder

    ' TextRenderer
    Dim TextRenderer As clsTextRenderer

    ' HtmlRenderer
    Dim HtmlRenderer As clsHtmlRenderer

    ' ExcelRenderer
    Dim ExcelRenderer As clsExcelRenderer

    '# 開発モード確認
    If gDevMode Then

        CreateDebugSheet

    Else

        DeleteSheetIfExists _
            SHEET_DEBUG_NODES

    End If

    Trace "HCP", _
        ModuleName & "." & ProcName

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Trace "HCP", _
        "Step=10 BuildFuncList"

    '# Procedure一覧取得
    Set FuncList = _
        BuildFuncList()

    Trace "HCP", _
        "Step=20 GetProcedureCode"

    '# Procedureソース取得
    CodeText = _
        GetProcedureCode( _
            ModuleName, _
            ProcName)

    Trace "HCP", _
        "CodeLength=" & _
        Len(CodeText)

    Trace "HCP", _
        "Step=30 AnalyzeProcedure"

    '# Procedure解析
    Set ProcNodes = _
        AnalyzeProcedure( _
            CodeText, _
            ModuleName, _
            ProcName, _
            FuncList)

    '-----------------------------
    '# デバッグシート初期化
    '-----------------------------
    DeleteSheetIfExists _
        SHEET_PROC_NODES

    DeleteSheetIfExists _
        SHEET_PROC_CHART

    DeleteSheetIfExists _
        SHEET_HCP_NODES

    Trace "HCP", _
        "Step=40 ConvertProcToHCP"

    '# ProcNode → HCPNode変換
    Set HCPNodes = _
        ConvertProcToHCP( _
            ProcNodes)

    '-----------------------------
    '# RenderText生成
    '-----------------------------
    Trace "HCP", _
        "Step=50 BuildRenderText"

    Set RenderBuilder = _
        New clsHCPRenderBuilder

    RenderBuilder.BuildRenderText _
        HCPNodes

    '-----------------------------
    '# デバッグ出力
    '-----------------------------
    If gDevMode Then

        ExportProcNodes _
            ProcNodes

        ExportHCPNodes _
            HCPNodes

    End If

    '-----------------------------
    '# TXT出力
    '-----------------------------
    Trace "HCP", _
        "Step=60 TextRenderer"

    Set TextRenderer = _
        New clsTextRenderer

    TextRenderer.Export _
        GetHCPOutputFolder(ModuleName) & _
        ProcName & "." & _
        ModuleName & _
        ".hcp.txt", _
        ProcName, _
        HCPNodes

    '-----------------------------
    '# HTML出力
    '-----------------------------
    Trace "HCP", _
        "Step=70 HtmlRenderer"

    Set HtmlRenderer = _
        New clsHtmlRenderer

    HtmlRenderer.Export _
        GetHCPOutputFolder(ModuleName) & _
        ProcName & "." & _
        ModuleName & _
        ".hcp.html", _
        ModuleName, _
        ProcName, _
        HCPNodes

    '-----------------------------
    '# Excel出力
    '-----------------------------
    If gDevMode Then

        Trace "HCP", _
            "Step=80 ExcelRenderer"

        Set ExcelRenderer = _
            New clsExcelRenderer

        ExcelRenderer.Export _
            HCPNodes, _
            ModuleName, _
            ProcName

        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES).Columns.AutoFit

    End If

    '-----------------------------
    '# INDEX更新
    '-----------------------------
    RebuildModuleIndex _
        ModuleName

    RebuildToolIndex

    Application.ScreenUpdating = True

    Exit Sub

ExitProc:

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    Trace "HCP", _
        "CreateHCP " & _
        Err.Description

    MsgBox _
        Err.Description, _
        vbExclamation, _
        "HCP生成エラー"

    Resume ExitProc

End Sub

'===================================================================
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
' ProcNode一覧を
' HCPNode一覧へ変換する
'
' @Remarks
' AnalyzeProcedureで生成された
' ProcNodeをHCPNodeへ変換する
'
' NODE_COMMENTは
' ProcNode時点で生成済みのため
' コメント専用変換処理は行わない
'
' IF文用TRUEノードおよび
' 制御構造終端MERGEノードを
' 自動生成する
'
' 変換後に
' ParentID
' Prev/Next
' Level
' を再構築する
'===================================================================
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

    ' Procedure情報
    Dim ProcInfo As clsProcInfo

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
    '# Proc → HCP変換
    '--------------------------------
    For Each ProcNode In ProcNodes

        '-----------------------------
        '# 制御構造終了処理
        '-----------------------------
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

        '-----------------------------
        '# HCPNode生成
        '-----------------------------
        Set HCPNode = New clsHCPNode

        HCPNode.NodeID = NewID

        IDMap.Add _
            CStr(ProcNode.NodeID), _
            NewID

        NewID = NewID + 1

        '-----------------------------
        '# ProcNode属性引継ぎ
        '-----------------------------
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

        HCPNode.CallName = _
            ProcNode.CallName

        HCPNode.InternalCalls = _
            ProcNode.InternalCalls

        '-----------------------------
        '# STARTノード業務説明取得
        '-----------------------------
        If HCPNode.NodeType = NODE_START Then

            Set ProcInfo = _
                ResolveProcedure( _
                    ProcNode.ModuleName, _
                    ProcNode.ProcName)

            If Not ProcInfo Is Nothing Then

                HCPNode.BusinessText = _
                    ProcInfo.JPName & _
                    "＜" & _
                    ProcInfo.ModuleName & _
                    "＞"

            End If

        End If

        '-----------------------------
        '# ELSEノード変換
        '-----------------------------
        If ProcNode.NodeType = NODE_ELSE Then

            HCPNode.Text = _
                "FALSE"

            HCPNode.NodeType = _
                NODE_FALSE

        Else

            HCPNode.Text = _
                ProcNode.NodeText

            '# CALLノード業務説明取得
            If HCPNode.NodeType = NODE_CALL Then

                HCPNode.BusinessText = _
                    GetBusinessText( _
                        ProcNode)

            End If

            '# 内部呼出追記
            If Trim$(HCPNode.InternalCalls) <> "" Then

                If Trim$(HCPNode.BusinessText) <> "" Then

                    HCPNode.BusinessText = _
                        HCPNode.BusinessText & _
                        vbCrLf & _
                        "内部呼出:" & _
                        HCPNode.InternalCalls

                Else

                    HCPNode.BusinessText = _
                        "内部呼出:" & _
                        HCPNode.InternalCalls

                End If

            End If

        End If

        '-----------------------------
        '# Debug.Print除外
        '-----------------------------
        If gHCPCompactMode Then

            If IsDebugPrint( _
                HCPNode.Text) Then

                GoTo ContinueLoop

            End If

        End If

        '-----------------------------
        '# 重複CALL除外
        '-----------------------------
        If gHCPCompactMode Then

            If HCPNode.NodeType = _
                NODE_CALL Then

                If HCPNode.Text = _
                    LastText Then

                    GoTo ContinueLoop

                End If

            End If

        End If

        '-----------------------------
        '# Tag取得集約
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

        '-----------------------------
        '# HCPNode登録
        '-----------------------------
        HCPNodes.Add HCPNode

        '-----------------------------
        '# 制御構造開始登録
        '-----------------------------
        If IsMergeNode( _
            HCPNode.NodeType) Then

            PushMergeStack _
                MergeStack, _
                HCPNode.NodeType, _
                HCPNode.NodeID

        End If

        '-----------------------------
        '# TRUEノード生成
        '-----------------------------
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
    '# ParentID再構築
    '--------------------------------
    RebuildParentID _
        HCPNodes, _
        IDMap

    '--------------------------------
    '# Prev/Next再構築
    '--------------------------------
    RebuildLinks _
        HCPNodes

    '--------------------------------
    '# TRUE枝補正
    '--------------------------------
    FixTrueBranch _
        HCPNodes

    '--------------------------------
    '# FALSE枝補正
    '--------------------------------
    FixFalseBranch _
        HCPNodes

    '--------------------------------
    '# Level再計算
    '--------------------------------
    RebuildHCPLevel _
        HCPNodes

    Trace "HCP", _
        "ConvertProcToHCP End Nodes=" & _
        HCPNodes.Count

    Set ConvertProcToHCP = _
        HCPNodes

End Function

'--------------------------------------------------
' HCPノード変換
'--------------------------------------------------

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
    
    ' RenderText設定しない

    '# TRUEノード登録
    HCPNodes.Add TrueNode

End Sub


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

    '# マージ対象判定
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

    '# NodeType:NodeID分解
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

    '# Procedure情報取得
    Set Info = _
        ResolveProcedure( _
            ProcNode.ModuleName, _
            ProcNode.CallName)

    If Info Is Nothing Then Exit Function

    '# JPName優先返却
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


'--------------------------------------------------
' HCP構造補正
'--------------------------------------------------

'===================================================================
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
' ProcNodeIDから
' HCPNodeIDへParentIDを変換する
'
' @Remarks
' Proc解析時のParentIDを
' HCPノード採番後のParentIDへ
' 再構築する
'
'===================================================================
Private Sub RebuildParentID( _
                ByVal HCPNodes As Collection, _
                ByVal IDMap As Object)

    ' HCPノード
    Dim Node As clsHCPNode
    
    ' ProcNode時点の親ID
    Dim OldParentID As Long

    '# ParentID変換開始
    For Each Node In HCPNodes

        '# MERGE除外
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

'===================================================================
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
' PrevNodeIDおよび
' NextNodeIDを再構築する
'
' @Remarks
' HCPNodes格納順から
' 前後リンクを再設定する
'
'===================================================================
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

'===================================================================
' @JPName
' HCPレベル再計算
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
' ParentIDを基準に
' HCP階層レベルを再計算する
'
' @Remarks
' ノード親子関係から
' 階層深度を再算出する
'
'===================================================================
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

'===================================================================
' @JPName
' HCPレベル算出
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
' 指定ノードの
' 階層レベルを算出する
'
' @Remarks
' ParentIDを遡り
' HCP階層を計算する
' 無限ループ防止のため
' 100階層で探索を打ち切る
'
'===================================================================
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

'===================================================================
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
' IFノード直下のTRUEノードへ
' TRUEブランチ配下ノードを
' 再配置する
'
'===================================================================
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

'===================================================================
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
' ELSEノード配下へ
' FALSEブランチを再配置する
'
'===================================================================
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


'--------------------------------------------------
' デバッグ支援
'--------------------------------------------------
'
' @JPName
' デバッグシート作成
'
' @Category
' Debug
'
' @Input
' なし
'
' @Output
' DBG_NODESシート
'
' @Summary
' HCP解析確認用の
' デバッグシートを生成する
'
' @Remarks
' ProcNodeおよびHCPNodeの
' 構造確認で使用する
'
Public Sub CreateDebugSheet()

    Dim ws As Worksheet

    On Error Resume Next

    '# DBG_NODES取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    On Error GoTo 0

    '# DBG_NODES作成
    If ws Is Nothing Then

        Set ws = _
            gContext.ManagementBook.Worksheets.Add

        ws.Name = _
            SHEET_DEBUG_NODES

    End If
    
    '# デバッグシート初期化
    ws.Cells.Clear
    
    '# ヘッダ出力
    ws.Range("A1:H1").Value = Array( _
        "Section", _
        "NodeID", _
        "ParentID", _
        "Level", _
        "BranchType", _
        "NodeType", _
        "Text", _
        "SourceLine")

    '# ヘッダ強調
    ws.Rows(1).Font.Bold = True

    '# 列幅自動調整
    ws.Columns("A:H").AutoFit

End Sub

' @JPName
' DBG_NODES出力行取得
'
' @Category
' Debug
'
' @Input
' なし
'
' @Output
' Row(Long)
'
' @Summary
' DBG_NODESシートの
' 次出力行を取得する
'
' @Remarks
' デバッグ情報追記位置を
' 算出する
'
Private Function GetDebugRow() As Long

    Dim ws As Worksheet

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    GetDebugRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row + 1

End Function

' @Category
' Debug
'
' @Input
' Title(String)
' Nodes(Collection)
'
' @Output
' DBG_NODESシート
'
' @Summary
' ProcNode一覧を
' デバッグシートへ出力する
'
' @Remarks
' ProcNode構造の確認で使用する
'
Public Sub DumpProcNodes( _
                ByVal Title As String, _
                ByVal Nodes As Collection)

    Dim ws As Worksheet
    Dim RowOut As Long
    Dim Node As clsProcNode

    '# DBG_NODES取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    '# 出力開始行取得
    RowOut = GetDebugRow()

    '# ProcNode一覧出力
    For Each Node In Nodes

        ws.Cells(RowOut, 1).Value = Title
        ws.Cells(RowOut, 2).Value = Node.NodeID
        ws.Cells(RowOut, 3).Value = Node.ParentID
        ws.Cells(RowOut, 4).Value = Node.Level
        ws.Cells(RowOut, 5).Value = Node.BranchType
        ws.Cells(RowOut, 6).Value = Node.NodeType
        ws.Cells(RowOut, 7).Value = Node.NodeText
        ws.Cells(RowOut, 8).Value = Node.SourceLine

        RowOut = RowOut + 1

    Next Node

End Sub

' @Category
' Debug
'
' @Input
' Title(String)
' Nodes(Collection)
'
' @Output
' DBG_NODESシート
'
' @Summary
' HCPNode一覧を
' デバッグシートへ出力する
'
' @Remarks
' HCP変換後の構造確認で使用する
'
Public Sub DumpHCPNodes( _
                ByVal Title As String, _
                ByVal Nodes As Collection)

    Dim ws As Worksheet
    Dim RowOut As Long
    Dim Node As clsHCPNode

    '# DBG_NODES取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    '# 出力開始行取得
    RowOut = GetDebugRow()
    
    '# HCPNode一覧出力
    For Each Node In Nodes

        ws.Cells(RowOut, 1).Value = Title
        ws.Cells(RowOut, 2).Value = Node.NodeID
        ws.Cells(RowOut, 3).Value = Node.ParentID
        ws.Cells(RowOut, 4).Value = Node.Level
        ws.Cells(RowOut, 5).Value = Node.BranchType
        ws.Cells(RowOut, 6).Value = Node.NodeType
        ws.Cells(RowOut, 7).Value = Node.Text
        ws.Cells(RowOut, 8).Value = Node.SourceLine

        RowOut = RowOut + 1

    Next Node

End Sub

' @Category
' Debug
'
' @Input
' Title(String)
' Nodes(Collection)
' StartID(Long)
' EndID(Long)
'
' @Output
' DBG_NODESシート
'
' @Summary
' 指定NodeID範囲の
' ProcNodeを出力する
'
Public Sub DumpProcNodeRange( _
                ByVal Title As String, _
                ByVal Nodes As Collection, _
                ByVal StartID As Long, _
                ByVal EndID As Long)

    Dim ws As Worksheet
    Dim RowOut As Long
    Dim Node As clsProcNode

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    RowOut = GetDebugRow()
    
    For Each Node In Nodes

        If Node.NodeID >= StartID _
        And Node.NodeID <= EndID Then

            ws.Cells(RowOut, 1).Value = Title
            ws.Cells(RowOut, 2).Value = Node.NodeID
            ws.Cells(RowOut, 3).Value = Node.ParentID
            ws.Cells(RowOut, 4).Value = Node.Level
            ws.Cells(RowOut, 5).Value = Node.BranchType
            ws.Cells(RowOut, 6).Value = Node.NodeType
            ws.Cells(RowOut, 7).Value = Node.NodeText
            ws.Cells(RowOut, 8).Value = Node.SourceLine

            RowOut = RowOut + 1

        End If

    Next Node

End Sub


' @Category
' Debug
'
' @Input
' Title(String)
' Nodes(Collection)
' StartID(Long)
' EndID(Long)
'
' @Output
' DBG_NODESシート
'
' @Summary
' 指定NodeID範囲の
' HCPNodeを出力する
'
Public Sub DumpHCPNodeRange( _
                ByVal Title As String, _
                ByVal Nodes As Collection, _
                ByVal StartID As Long, _
                ByVal EndID As Long)

    Dim ws As Worksheet
    Dim RowOut As Long
    Dim Node As clsHCPNode

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_DEBUG_NODES)

    RowOut = GetDebugRow()
    
    For Each Node In Nodes

        If Node.NodeID >= StartID _
        And Node.NodeID <= EndID Then

            ws.Cells(RowOut, 1).Value = Title
            ws.Cells(RowOut, 2).Value = Node.NodeID
            ws.Cells(RowOut, 3).Value = Node.ParentID
            ws.Cells(RowOut, 4).Value = Node.Level
            ws.Cells(RowOut, 5).Value = Node.BranchType
            ws.Cells(RowOut, 6).Value = Node.NodeType
            ws.Cells(RowOut, 7).Value = Node.Text
            ws.Cells(RowOut, 8).Value = Node.SourceLine

            RowOut = RowOut + 1

        End If

    Next Node

End Sub

'--------------------------------------------------
' HCPノード出力
'--------------------------------------------------

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
    ws.Range("A1:L1").Value = _
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
            "CallName", _
            "InternalCalls")

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
        
        ws.Cells(RowOut, 11) = Node.CallName

        ws.Cells(RowOut, 12) = Node.InternalCalls
        
        RowOut = RowOut + 1

    Next Node
    
    '# HCPNodesシート体裁設定
    ws.Columns.AutoFit

End Sub

'--------------------------------------------------
' ユーティリティ
'--------------------------------------------------

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


