Attribute VB_Name = "modHCPAnalyzer"
Option Explicit

Public Sub CreateHCP( _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    Dim CodeText As String

    Dim ProcNodes As Collection
    Dim HCPNodes As Collection

    Dim FuncList As Object
    
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    
    Set FuncList = BuildFuncList()
    
    CodeText = GetProcedureCode( _
                    ModuleName, _
                    ProcName)

    Set ProcNodes = AnalyzeProcedure( _
                        CodeText, _
                        ModuleName, _
                        ProcName, _
                        FuncList)

    ' 旧シート削除
    DeleteSheetIfExists SHEET_PROC_NODES
    DeleteSheetIfExists SHEET_PROC_CHART
    DeleteSheetIfExists SHEET_HCP_NODES

    Set HCPNodes = ConvertProcToHCP( _
                        ProcNodes)

    If DEV_MODE Then

        ExportProcNodes ProcNodes

        ExportHCPNodes HCPNodes

    End If

    ExportHCPText _
        HCPNodes, _
        ModuleName, _
        ProcName

    Application.ScreenUpdating = True
    
    MsgBox _
        "HCPチャート生成が完了しました。" & _
        vbCrLf & _
        "[" & ModuleName & "." & ProcName & "]", _
        vbInformation, _
        "HCP生成"
    
    Exit Sub

ExitProc:

    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:

    MsgBox _
        Err.Description, _
        vbExclamation, _
        "HCP生成エラー"

    Resume ExitProc

End Sub


Public Function ConvertProcToHCP( _
                ByVal ProcNodes As Collection) _
                As Collection

    Dim HCPNodes As Collection
    
    Dim MergeStack As Collection

    Dim ProcNode As clsProcNode
    Dim HCPNode As clsHCPNode

    Dim LastTagNode As clsHCPNode

    Dim LastText As String

    Dim IDMap As Object

    Dim NewID As Long

    Set HCPNodes = New Collection
    
    Set MergeStack = New Collection

    Set IDMap = _
        CreateObject("Scripting.Dictionary")

    NewID = 1

    '--------------------------------
    ' Proc → HCP変換
    '--------------------------------
    For Each ProcNode In ProcNodes

        Select Case ProcNode.NodeType

            Case NODE_ENDIF

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

                GoTo ContinueLoop

            Case NODE_ENDSELECT

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

                GoTo ContinueLoop

            Case NODE_NEXT

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                   MergeStack.Count

                GoTo ContinueLoop

            Case NODE_LOOP

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

                GoTo ContinueLoop

            Case NODE_ENDWITH

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

                GoTo ContinueLoop

            Case NODE_WEND

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

                GoTo ContinueLoop

            Case NODE_COMPILE_ENDIF

                AddMergeNode _
                    HCPNodes, _
                    NewID, _
                    GetMergeParentID( _
                        MergeStack(MergeStack.Count)), _
                    ProcNode.SourceLine

                MergeStack.Remove _
                    MergeStack.Count

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
    
            HCPNode.Text = ProcNode.NodeText

            HCPNode.BusinessText = _
                GetBusinessText(ProcNode)

            If HCPNode.BusinessText = "" Then

                HCPNode.BusinessText = ""

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

        If HCPNode.NodeType = NODE_IF Then

            ' 制御ノード登録
            MergeStack.Add NODE_IF & ":" & HCPNode.NodeID
            
            Dim TrueNode As clsHCPNode

            Set TrueNode = New clsHCPNode

            TrueNode.NodeID = NewID

            NewID = NewID + 1

            TrueNode.ParentID = _
                -HCPNode.NodeID

            TrueNode.NodeType = _
                NODE_TRUE

            TrueNode.BranchType = _
                BRANCH_TRUE

            TrueNode.Text = _
                "TRUE"

            HCPNodes.Add TrueNode

        End If
        
        If HCPNode.NodeType = NODE_FOR Then

            ' 制御ノード登録
            MergeStack.Add NODE_FOR & ":" & HCPNode.NodeID

        End If
        
        If HCPNode.NodeType = NODE_SELECT Then

            ' 制御ノード登録
            MergeStack.Add NODE_SELECT & ":" & HCPNode.NodeID

        End If

        If HCPNode.NodeType = NODE_DO Then

            MergeStack.Add _
                NODE_DO & ":" & HCPNode.NodeID

        End If

        If HCPNode.NodeType = NODE_WITH Then

            MergeStack.Add _
                NODE_WITH & ":" & HCPNode.NodeID

        End If

        If HCPNode.NodeType = NODE_WHILE Then

            MergeStack.Add _
                NODE_WHILE & ":" & HCPNode.NodeID

        End If

        If HCPNode.NodeType = NODE_COMPILE_IF Then

            MergeStack.Add _
                NODE_COMPILE_IF & ":" & HCPNode.NodeID

        End If

        LastText = HCPNode.Text

ContinueLoop:

    Next ProcNode

    '--------------------------------
    ' ParentID補正
    '--------------------------------
    RebuildParentID _
        HCPNodes, _
        IDMap

    '--------------------------------
    ' Prev/Next補正
    '--------------------------------
    RebuildLinks _
        HCPNodes
  
    FixTrueBranch _
        HCPNodes

    FixFalseBranch _
        HCPNodes
    
    '--------------------------------
    ' Level再計算
    '--------------------------------

    RebuildHCPLevel _
        HCPNodes
    
    Set ConvertProcToHCP = _
        HCPNodes
    
End Function

Private Function GetMergeParentID( _
                ByVal StackValue As String) _
                As Long

    Dim Parts() As String

    Parts = Split(StackValue, ":")

    If UBound(Parts) <> 1 Then
        GetMergeParentID = 0
        Exit Function
    End If

    GetMergeParentID = CLng(Parts(1))

End Function

Private Function GetBusinessText( _
                ByVal ProcNode As clsProcNode) _
                As String

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

Private Sub AddMergeNode( _
                ByVal HCPNodes As Collection, _
                ByRef NewID As Long, _
                ByVal ParentID As Long, _
                ByVal Text As String)

    Dim HCPNode As clsHCPNode

    Set HCPNode = New clsHCPNode

    HCPNode.NodeID = NewID

    NewID = NewID + 1

    HCPNode.ParentID = ParentID

    HCPNode.NodeType = NODE_MERGE

    HCPNode.Text = Text

    HCPNodes.Add HCPNode

End Sub


Private Sub RebuildParentID( _
                ByVal HCPNodes As Collection, _
                ByVal IDMap As Object)

    Dim Node As clsHCPNode
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


Private Sub RebuildLinks( _
                ByVal HCPNodes As Collection)

    Dim i As Long

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


Private Sub RebuildHCPLevel( _
                ByVal HCPNodes As Collection)

    Dim NodeMap As Object

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


Private Function CalcLevel( _
                ByVal Node As clsHCPNode, _
                ByVal NodeMap As Object) _
                As Long

    Dim ParentID As Long
    Dim Level As Long
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


Private Sub FixTrueBranch( _
                ByVal HCPNodes As Collection)

    Dim i As Long
    Dim j As Long

    Dim IfID As Long
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


Private Sub FixFalseBranch( _
                ByVal HCPNodes As Collection)

    Dim i As Long

    Dim Node As clsHCPNode

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


Public Sub ExportHCPNodes( _
                ByVal HCPNodes As Collection)

    Dim ws As Worksheet

    Dim Node As clsHCPNode

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


Public Function GetChildNodes( _
                ByVal ParentID As Long, _
                ByVal HCPNodes As Collection) _
                As Collection

    Dim Result As Collection

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


Private Sub LayoutTree( _
                ByVal Node As clsHCPNode, _
                ByVal HCPNodes As Collection, _
                ByVal Level As Long, _
                ByRef CurrentY As Double)

    Dim Children As Collection
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
            GetHCPSymbol = "◆ "

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

Public Sub ExportHCPText( _
                ByVal HCPNodes As Collection, _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    Dim ws As Worksheet
    Dim RowOut As Long
    
    Dim Node As clsHCPNode
    Dim StartText As String

    On Error Resume Next

    Set ws = gContext.ManagementBook.Worksheets( _
                SHEET_HCP_CHART)
    
    On Error GoTo 0

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
        ModuleName & "." & ProcName
    
    ws.Cells(2, 1).Value = "HCP"
    ws.Cells(2, 2).Value = "Business"

    RowOut = 3

    ws.Cells(RowOut, 1).Value = "┬"
    RowOut = RowOut + 1

    ws.Cells(RowOut, 1).Value = _
        "〇│" & StartText
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
    
    ws.Columns("A").ColumnWidth = 150
    
    ws.Columns("B").ColumnWidth = 60
    
    ws.Columns("A").WrapText = True
    
    gContext.ManagementBook.Activate
    ws.Activate

End Sub

Private Sub RenderNode( _
                ByVal HCPNodes As Collection, _
                ByVal ParentID As Long, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

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

Private Sub RenderIfBlock( _
                ByVal HCPNodes As Collection, _
                ByVal IfNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode

    Dim TrueNode As clsHCPNode
    Dim FalseNode As clsHCPNode
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

Private Sub RenderForBlock( _
                ByVal HCPNodes As Collection, _
                ByVal ForNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
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

Private Sub RenderSelectBlock( _
                ByVal HCPNodes As Collection, _
                ByVal SelectNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
    Dim MergeNode As clsHCPNode

    Dim CaseCount As Long
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

Private Sub RenderDoBlock( _
                ByVal HCPNodes As Collection, _
                ByVal DoNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
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

Private Sub RenderWithBlock( _
                ByVal HCPNodes As Collection, _
                ByVal WithNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
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

Private Sub RenderWhileBlock( _
                ByVal HCPNodes As Collection, _
                ByVal WhileNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
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

Private Sub RenderCompileBlock( _
                ByVal HCPNodes As Collection, _
                ByVal CompileNode As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

    Dim Node As clsHCPNode
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

Private Sub OutputHCPNode( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String, _
                ByVal ws As Worksheet, _
                ByRef RowOut As Long)

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

Private Function GetRenderText( _
                ByVal Node As clsHCPNode) _
                As String

    GetRenderText = _
        FormatHCPText(Node.Text)

End Function

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

Private Function GetDisplayText( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String) _
                As String

    GetDisplayText = _
        AddIndentToMultiLine( _
            GetRenderText(Node), _
            Prefix)

End Function

Public Sub DeleteSheetIfExists( _
                ByVal SheetName As String)

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

