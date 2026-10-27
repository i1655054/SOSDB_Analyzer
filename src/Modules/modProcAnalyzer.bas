Attribute VB_Name = "modProcAnalyzer"
Option Explicit

Public Const NODE_START  As String = "START"
Public Const NODE_END    As String = "END"

Public Const NODE_DECLARE As String = "DECLARE"
Public Const NODE_CONTROL As String = "CONTROL"
Public Const NODE_MERGE As String = "MERGE"
Public Const NODE_REFERENCE As String = "REFERENCE"

Public Const NODE_CALL   As String = "CALL"

Public Const NODE_IF     As String = "IF"
Public Const NODE_ELSE   As String = "ELSE"
Public Const NODE_ENDIF  As String = "ENDIF"
Public Const NODE_INLINE_IF = "INLINE_IF"

Public Const NODE_TRUE  As String = "TRUE"
Public Const NODE_FALSE As String = "FALSE"

Public Const NODE_FOR    As String = "FOR"
Public Const NODE_NEXT   As String = "NEXT"

Public Const NODE_SELECT As String = "SELECT"
Public Const NODE_CASE   As String = "CASE"
Public Const NODE_ENDSELECT As String = "ENDSELECT"

Public Const NODE_EXIT   As String = "EXIT"

Public Const NODE_GOTO   As String = "GOTO"
Public Const NODE_LABEL  As String = "LABEL"
Public Const NODE_RESUME As String = "RESUME"

Public Const NODE_NORMAL As String = "NORMAL"

Public Const NODE_DO          As String = "DO"
Public Const NODE_LOOP        As String = "LOOP"

Public Const NODE_WITH        As String = "WITH"
Public Const NODE_ENDWITH     As String = "ENDWITH"

Public Const NODE_WHILE       As String = "WHILE"
Public Const NODE_WEND        As String = "WEND"

Public Const NODE_COMPILE_IF      As String = "COMPILE_IF"
Public Const NODE_COMPILE_ELSEIF  As String = "COMPILE_ELSEIF"
Public Const NODE_COMPILE_ELSE    As String = "COMPILE_ELSE"
Public Const NODE_COMPILE_ENDIF   As String = "COMPILE_ENDIF"

Public Const BRANCH_NONE  As String = ""
Public Const BRANCH_TRUE  As String = "TRUE"
Public Const BRANCH_FALSE As String = "FALSE"

Private Const PROC_KIND As Long = 0

'------------------------------------------------
' キーワード判定
'------------------------------------------------
Private Function StartsWithKeyword( _
                ByVal LineText As String, _
                ByVal Keyword As String) _
                As Boolean

    Dim NextChar As String

    If Len(LineText) < Len(Keyword) Then
        Exit Function
    End If

    If StrComp( _
            Left$(LineText, Len(Keyword)), _
            Keyword, _
            vbTextCompare) <> 0 Then
        Exit Function
    End If

    ' 完全一致
    If Len(LineText) = Len(Keyword) Then

        StartsWithKeyword = True
        Exit Function

    End If

    NextChar = Mid$( _
                    LineText, _
                    Len(Keyword) + 1, _
                    1)

    Select Case NextChar

        Case " ", vbTab, "(", ":"

            StartsWithKeyword = True

    End Select

End Function

'------------------------------------------------
' 行種別判定
'------------------------------------------------
Private Function GetNodeType( _
                ByVal LineText As String) _
                As String

    LineText = Trim$(LineText)

    If StartsWithKeyword(LineText, "#ElseIf") Then

        GetNodeType = NODE_COMPILE_ELSEIF

    ElseIf StrComp( _
                LineText, _
                "#Else", _
                vbTextCompare) = 0 Then

        GetNodeType = NODE_COMPILE_ELSE

    ElseIf StrComp( _
                LineText, _
                "#End If", _
                vbTextCompare) = 0 Then

        GetNodeType = NODE_COMPILE_ENDIF

    ElseIf StartsWithKeyword(LineText, "#If") Then

        GetNodeType = NODE_COMPILE_IF
    
    ElseIf StartsWithKeyword(LineText, "Dim") Then

        GetNodeType = NODE_DECLARE

    ElseIf StartsWithKeyword(LineText, "Const") Then

        GetNodeType = NODE_DECLARE

    ElseIf StartsWithKeyword(LineText, "On Error") Then

        GetNodeType = NODE_CONTROL

    ElseIf StartsWithKeyword(LineText, "If") Then

        If IsInlineIf(LineText) Then

            GetNodeType = NODE_INLINE_IF

        Else

            GetNodeType = NODE_IF

        End If

    ElseIf StrComp(LineText, _
                   "End If", _
                   vbTextCompare) = 0 Then

        GetNodeType = NODE_ENDIF

    ElseIf StartsWithKeyword(LineText, "ElseIf") Then

        GetNodeType = NODE_ELSE

    ElseIf StrComp(LineText, _
                   "Else", _
                   vbTextCompare) = 0 Then

        GetNodeType = NODE_ELSE

    ElseIf StartsWithKeyword(LineText, "For") Then

        GetNodeType = NODE_FOR

    ElseIf StartsWithKeyword(LineText, "Next") Then

        GetNodeType = NODE_NEXT

    ElseIf StartsWithKeyword(LineText, "Do") Then

        GetNodeType = NODE_DO

    ElseIf StartsWithKeyword(LineText, "Loop") Then

        GetNodeType = NODE_LOOP

    ElseIf StartsWithKeyword(LineText, "While") Then

        GetNodeType = NODE_WHILE

    ElseIf StrComp(LineText, _
                   "Wend", _
                   vbTextCompare) = 0 Then

        GetNodeType = NODE_WEND

    ElseIf StartsWithKeyword(LineText, "With") Then

        GetNodeType = NODE_WITH

    ElseIf StrComp(LineText, _
                   "End With", _
                   vbTextCompare) = 0 Then

        GetNodeType = NODE_ENDWITH

    ElseIf StartsWithKeyword(LineText, "Select Case") Then

        GetNodeType = NODE_SELECT

    ElseIf StartsWithKeyword(LineText, "Case") Then

        GetNodeType = NODE_CASE

    ElseIf StrComp(LineText, _
                   "End Select", _
                   vbTextCompare) = 0 Then

        GetNodeType = NODE_ENDSELECT

    ElseIf StartsWithKeyword(LineText, "Resume") Then

        GetNodeType = NODE_RESUME

    ElseIf StartsWithKeyword(LineText, "GoTo") Then

        GetNodeType = NODE_GOTO

    ElseIf Right$(LineText, 1) = ":" Then

        GetNodeType = NODE_LABEL

    ElseIf StartsWithKeyword(LineText, "Exit") Then

        GetNodeType = NODE_EXIT

    Else

        GetNodeType = NODE_NORMAL

    End If

End Function


' 呼び出し先取得
Private Function ExtractCallNames( _
                ByVal LineText As String, _
                ByVal FuncList As Object) _
                As Collection

    Dim Names As Collection

    Dim Key As Variant

    Set Names = New Collection

    LineText = Trim$(LineText)

    '----------------------------------
    ' 宣言行除外
    '----------------------------------
    If IsProcedureLine(LineText) Then

        Set ExtractCallNames = Names

        Exit Function

    End If

    '----------------------------------
    ' コメント除外
    '----------------------------------
    If Left$(LineText, 1) = "'" Then

        Set ExtractCallNames = Names

        Exit Function

    End If

    '----------------------------------
    ' CALL解析対象外
    '----------------------------------
    If GetNodeType(LineText) <> NODE_NORMAL Then

        Set ExtractCallNames = Names

        Exit Function

    End If

    '----------------------------------
    ' 登録済みProcedureのみ検索
    '----------------------------------
    For Each Key In FuncList.Keys

        If IsFunctionCall( _
                LineText, _
                CStr(Key)) Then

            Names.Add CStr(Key)

        End If

    Next Key

    Set ExtractCallNames = Names

End Function

Private Function GetCurrentParentID( _
                ByVal ParentStack As Collection) _
                As Long

    If ParentStack.Count = 0 Then

        GetCurrentParentID = 0

    Else

        GetCurrentParentID = _
            ParentStack(ParentStack.Count)

    End If

End Function

'=================================================
' Proc解析
'=================================================
'
' @JPName
' Proc解析
'
' @Category
' ProcAnalyzer
'
' @Input
' CodeText(String)
' ProcName(String)
'
' @Output
' Nodes(Collection)
'
' @Summary
' プロシージャ内部を解析し
' ProcNode一覧を生成する
Public Function AnalyzeProcedure( _
                ByVal CodeText As String, _
                ByVal ModuleName As String, _
                ByVal ProcName As String, _
                ByVal FuncList As Object) _
                As Collection

    Dim Nodes As Collection

    Dim Node As clsProcNode

    Dim PrevNode As clsProcNode

    Dim LogicalLines As Collection
    Dim LineText As Variant

    Dim i As Long

    Dim NodeID As Long
    Dim LastNodeID As Long

    Dim NodeType As String

    Dim CompileLevel As Long
    Dim Level As Long
    Dim OutputLevel As Long

    Dim ParentStack As Collection

    Dim CurrentBranch As String

    Dim CallNames As Collection
    Dim CallName As Variant

    Dim CurrentCaseID As Long
    
    Dim BranchStack As Collection
    
    Dim ProcHeader As String

    Set Nodes = New Collection
    Set ParentStack = New Collection
    Set BranchStack = New Collection

    CompileLevel = 0
    
    NodeID = 1
    Level = 0
    CurrentCaseID = 0

    CurrentBranch = BRANCH_NONE

    ProcHeader = _
        ExtractProcedureHeader(CodeText)
    
    '--------------------------
    ' START
    '--------------------------
    Set Node = New clsProcNode

    Node.NodeID = NodeID
    Node.ParentID = 0
    Node.Level = 0
    Node.BranchType = BRANCH_NONE
    Node.ProcName = ProcName
    Node.ModuleName = ModuleName
    Node.NodeType = NODE_START
    Node.NodeText = ProcHeader
    Node.SourceLine = ""

    AddProcNode _
        Nodes, _
        Node, _
        PrevNode, _
        LastNodeID

    NodeID = NodeID + 1

    '--------------------------
    ' 本体
    '--------------------------
    Set LogicalLines = _
        BuildLogicalLines(CodeText)

    For Each LineText In LogicalLines

        If IsProcedureLine(Trim$(CStr(LineText))) Then
            GoTo NextLine
        End If

        If Left$(Trim$(CStr(LineText)), 1) = "'" Then
            GoTo NextLine
        End If

        i = i + 1

        NodeType = _
            GetNodeType(CStr(LineText))

        Select Case NodeType

            Case NODE_COMPILE_IF

                CompileLevel = _
                   CompileLevel + 1

                GoTo NextLine

            Case NODE_COMPILE_ELSEIF

                GoTo NextLine

            Case NODE_COMPILE_ELSE

                GoTo NextLine

            Case NODE_COMPILE_ENDIF

                If CompileLevel > 0 Then

                    CompileLevel = _
                        CompileLevel - 1

                End If

                GoTo NextLine

        End Select
        
        If CompileLevel > 0 Then

            GoTo NextLine

        End If

        Set CallNames = _
            ExtractCallNames( _
                CStr(LineText), _
                FuncList)

        OutputLevel = Level

        '--------------------------
        ' END SELECT
        '--------------------------
        If NodeType = NODE_ENDSELECT Then

            CurrentCaseID = 0

            If ParentStack.Count > 0 Then
                ParentStack.Remove ParentStack.Count
            End If

            If ParentStack.Count > 0 Then
                ParentStack.Remove ParentStack.Count
            End If

            If Level > 0 Then
                Level = Level - 1
            End If

            OutputLevel = Level

            CurrentBranch = BRANCH_NONE

        End If

        '--------------------------
        ' END IF / NEXT
        '--------------------------
        If NodeType = NODE_ENDIF _
        Or NodeType = NODE_NEXT _
        Or NodeType = NODE_LOOP _
        Or NodeType = NODE_ENDWITH _
        Or NodeType = NODE_WEND Then

            If Level > 0 Then
                Level = Level - 1
            End If

            OutputLevel = Level

            If ParentStack.Count > 0 Then
                ParentStack.Remove ParentStack.Count
            End If

            If BranchStack.Count > 0 Then
            
                CurrentBranch = _
                    BranchStack(BranchStack.Count)
                
                BranchStack.Remove _
                    BranchStack.Count
            Else
                
                CurrentBranch = BRANCH_NONE
            
            End If

        End If

        '--------------------------
        ' ELSE
        '--------------------------
        If NodeType = NODE_ELSE Then

            CurrentBranch = BRANCH_FALSE

            OutputLevel = Level - 1

            If OutputLevel < 0 Then
                OutputLevel = 0
            End If

        End If

        '--------------------------
        ' CASE切替
        '--------------------------
        If NodeType = NODE_CASE Then

            If CurrentCaseID <> 0 Then

                If ParentStack.Count > 0 Then

                    ParentStack.Remove ParentStack.Count

                End If

            End If

        End If
        
        '--------------------------
        ' 制御構造
        '--------------------------
        If NodeType <> NODE_NORMAL Then

            Set Node = New clsProcNode

            Node.NodeID = NodeID
            Node.ParentID = _
                GetCurrentParentID(ParentStack)

            Node.Level = OutputLevel

            Node.BranchType = CurrentBranch
            Node.ProcName = ProcName
            Node.ModuleName = ModuleName
            Node.NodeType = NodeType

            Node.SourceLine = _
                Trim$(CStr(LineText))

            Node.NodeText = _
                Trim$(CStr(LineText))

            Node.LineNo = i

            AddProcNode _
                Nodes, _
                Node, _
                PrevNode, _
                LastNodeID

            Select Case NodeType

                Case NODE_IF

                    ParentStack.Add NodeID
                    
                    BranchStack.Add CurrentBranch

                    Level = Level + 1

                    CurrentBranch = BRANCH_TRUE

                Case NODE_FOR

                    ParentStack.Add NodeID

                    Level = Level + 1

                Case NODE_DO

                    ParentStack.Add NodeID

                    Level = Level + 1

                Case NODE_WITH

                    ParentStack.Add NodeID

                   Level = Level + 1

                Case NODE_WHILE

                    ParentStack.Add NodeID

                    Level = Level + 1

                Case NODE_SELECT

                    ParentStack.Add NodeID

                    Level = Level + 1

                    CurrentCaseID = 0

                Case NODE_CASE

                    ParentStack.Add NodeID

                    CurrentCaseID = NodeID

                    Level = OutputLevel + 1

                Case NODE_LABEL

                    Node.NodeText = _
                        Trim$(CStr(LineText))
                
                Case Else

            End Select

            NodeID = NodeID + 1

        '--------------------------
        ' CALL
        '--------------------------
        ElseIf CallNames.Count > 0 Then

            For Each CallName In CallNames

                Set Node = New clsProcNode
    
                Node.NodeID = NodeID
                Node.ParentID = _
                    GetCurrentParentID(ParentStack)

                Node.Level = Level

                Node.BranchType = CurrentBranch
                Node.ProcName = ProcName
                Node.ModuleName = ModuleName
                Node.NodeType = NODE_CALL

                Node.NodeText = _
                    Trim$(CStr(LineText))

                Node.SourceLine = _
                    Trim$(CStr(LineText))
                
                Node.CallName = _
                    CStr(CallName)

                Node.LineNo = i

                AddProcNode _
                    Nodes, _
                    Node, _
                    PrevNode, _
                    LastNodeID

                NodeID = NodeID + 1

            Next CallName

        '--------------------------
        ' Reference
        '--------------------------
        ElseIf IsReferenceLine( _
                    Trim$(CStr(LineText))) Then

            Set Node = New clsProcNode

            Node.NodeID = NodeID

            Node.ParentID = _
                GetCurrentParentID(ParentStack)

            Node.Level = Level

            Node.BranchType = CurrentBranch

            Node.ProcName = ProcName
            
            Node.ModuleName = ModuleName

            Node.NodeType = NODE_REFERENCE

            Node.NodeText = _
                Trim$(CStr(LineText))

            Node.SourceLine = _
                Trim$(CStr(LineText))

            AddProcNode _
                Nodes, _
                Node, _
                PrevNode, _
                LastNodeID

            NodeID = NodeID + 1
    
        '--------------------------
        ' NORMAL
        '--------------------------
        ElseIf IsOutputNormal( _
                Trim$(CStr(LineText))) Then

            Set Node = New clsProcNode

            Node.NodeID = NodeID

            Node.ParentID = _
                GetCurrentParentID(ParentStack)

            Node.Level = Level

            Node.BranchType = CurrentBranch

            Node.ProcName = ProcName
            
            Node.ModuleName = ModuleName

            Node.NodeType = NODE_NORMAL

            Node.NodeText = _
                Trim$(CStr(LineText))

            Node.SourceLine = _
                Trim$(CStr(LineText))

            Node.LineNo = i

            AddProcNode _
                Nodes, _
                Node, _
                PrevNode, _
                LastNodeID

            NodeID = NodeID + 1

        End If

NextLine:
    
    Next

    '--------------------------
    ' END
    '--------------------------
    Set Node = New clsProcNode

    Node.NodeID = NodeID
    Node.ParentID = 0
    Node.Level = 0
    Node.BranchType = BRANCH_NONE
    Node.ProcName = ProcName
    Node.ModuleName = ModuleName
    Node.NodeType = NODE_END
    Node.NodeText = _
        GetProcedureEndText( _
            ProcHeader)
    Node.SourceLine = ""

    AddProcNode _
        Nodes, _
        Node, _
        PrevNode, _
        LastNodeID

    Set AnalyzeProcedure = Nodes

End Function

Private Function GetProcedureEndText( _
                ByVal ProcHeader As String) _
                As String

    If StartsWithKeyword( _
            ProcHeader, _
            "Public Function") _
    Or StartsWithKeyword( _
            ProcHeader, _
            "Private Function") _
    Or StartsWithKeyword( _
            ProcHeader, _
            "Friend Function") Then

        GetProcedureEndText = _
            "End Function"

    ElseIf StartsWithKeyword( _
                ProcHeader, _
                "Public Property") _
        Or StartsWithKeyword( _
                ProcHeader, _
                "Private Property") _
        Or StartsWithKeyword( _
                ProcHeader, _
                "Friend Property") Then

        GetProcedureEndText = _
            "End Property"

    Else

        GetProcedureEndText = _
            "End Sub"

    End If

End Function

Private Function ExtractProcedureHeader( _
                ByVal CodeText As String) _
                As String

    Dim Lines() As String
    Dim i As Long
    Dim Header As String
    Dim Text As String

    Lines = Split(CodeText, vbCrLf)

    For i = LBound(Lines) To UBound(Lines)

        Text = Trim$(Lines(i))

        If IsProcedureLine(Text) Then

            Header = Text

            Do While _
                Right$(Trim$(Lines(i)), 1) = "_"

                i = i + 1

                Header = _
                    Replace(Header, "_", "") & _
                    " " & _
                    Trim$(Lines(i))

            Loop

            ExtractProcedureHeader = Header

            Exit Function

        End If

    Next i

End Function

Private Function IsOutputNormal( _
                ByVal LineText As String) _
                As Boolean

    Dim Text As String

    Text = Trim$(LineText)

    If Text = "" Then Exit Function
    
    ' コメントは常に除外
    If Left$(Text, 1) = "'" Then Exit Function
    
    ' 終了構文は常に除外
    If UCase$(Text) = "END SUB" Then Exit Function

    If UCase$(Text) = "END FUNCTION" Then Exit Function

    If UCase$(Text) = "END PROPERTY" Then Exit Function
    
    ' 全表示ならここで終了
    If Not gHCPCompactMode Then
    
        IsOutputNormal = True
        Exit Function
    
    End If

    If Left$(Text, 6) = "Option" Then Exit Function

    If Left$(Text, 4) = "Dim " Then Exit Function

    If Left$(Text, 6) = "Const " Then Exit Function

    If Left$(Text, 8) = "On Error" Then Exit Function

    If InStr(1, _
          Text, _
          "Debug.Print", _
          vbTextCompare) > 0 Then
        Exit Function
    End If

    '-------------------------
    ' 書式・表示系除外
    '-------------------------
    If InStr(Text, ".AutoFit") > 0 Then Exit Function

    If InStr(Text, ".AutoFilter") > 0 Then Exit Function

    If InStr(Text, ".Font.") > 0 Then Exit Function

    IsOutputNormal = True

End Function

Private Function IsReferenceLine( _
                ByVal LineText As String) _
                As Boolean

    If InStr(LineText, _
             "MonShinSheet_id(") > 0 Then

        IsReferenceLine = True
        Exit Function

    End If

    If InStr(LineText, _
             "MonshinMap(") > 0 Then

        IsReferenceLine = True
        Exit Function

    End If

End Function

Private Function IsInlineIf( _
                ByVal LineText As String) _
                As Boolean

    Dim PosThen As Long

    LineText = Trim$(LineText)

    If Left$(LineText, 2) <> "If" Then Exit Function

    PosThen = InStr( _
                1, _
                LineText, _
                " Then", _
                vbTextCompare)

    If PosThen = 0 Then Exit Function

    If Trim$(Mid$( _
            LineText, _
            PosThen + 5)) <> "" Then

        IsInlineIf = True

    End If

End Function



Private Function BuildLogicalLines( _
                ByVal CodeText As String) _
                As Collection

    Dim Result As Collection

    Dim Lines() As String

    Dim i As Long

    Dim CurrentLine As String

    Set Result = New Collection

    Lines = Split(CodeText, vbCrLf)

    For i = LBound(Lines) To UBound(Lines)

        Dim LineText As String

        LineText = Trim$(Lines(i))

        If Right$(LineText, 1) = "_" Then

            CurrentLine = _
                CurrentLine & " " & _
                Left$(LineText, _
                      Len(LineText) - 1)

        Else

            CurrentLine = _
                CurrentLine & " " & _
                LineText

            Result.Add Trim$(CurrentLine)

            CurrentLine = ""

        End If

    Next i

    If Len(CurrentLine) > 0 Then

        Result.Add Trim$(CurrentLine)

    End If

    Set BuildLogicalLines = Result

End Function

Public Sub ExportProcNodes( _
                ByVal Nodes As Collection)

    Dim ws As Worksheet

    Dim Node As clsProcNode

    Dim RowOut As Long

    On Error Resume Next

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_NODES)

    On Error GoTo 0

    If gContext Is Nothing Then

        On Error Resume Next

        Set ws = gContext.ManagementBook.Worksheets( _
                    SHEET_PROC_NODES)


        On Error GoTo 0

        If ws Is Nothing Then

            Set ws = _
                gContext.ManagementBook.Worksheets.Add


            ws.Name = _
                SHEET_PROC_NODES

        End If

    Else

        On Error Resume Next

        Set ws = _
            gContext.ManagementBook.Worksheets( _
                SHEET_PROC_NODES)

        On Error GoTo 0

        If ws Is Nothing Then

            Set ws = _
                gContext.ManagementBook.Worksheets.Add

            ws.Name = _
                SHEET_PROC_NODES

        End If

    End If

    '------------------------------------
    ' Clearは使わない
    '------------------------------------
    ws.Range("A:H").Value = ""

    ws.Range("A:H").ClearFormats

    '------------------------------------
    ' Header
    '------------------------------------
    ws.Range("A1:I1").Value = Array( _
        "NodeID", _
        "ParentID", _
        "PrevNodeID", _
        "NextNodeID", _
        "Level", _
        "NodeType", _
        "BranchType", _
        "Text", _
        "SourceText")

    RowOut = 2

    '------------------------------------
    ' Nodes
    '------------------------------------
    For Each Node In Nodes

        ws.Cells(RowOut, 1) = Node.NodeID
        ws.Cells(RowOut, 2) = Node.ParentID
        ws.Cells(RowOut, 3) = Node.PrevNodeID
        ws.Cells(RowOut, 4) = Node.NextNodeID
        ws.Cells(RowOut, 5) = Node.Level
        ws.Cells(RowOut, 6) = Node.NodeType
        ws.Cells(RowOut, 7) = Node.BranchType
        ws.Cells(RowOut, 8) = Node.NodeText
        ws.Cells(RowOut, 9) = Node.SourceLine

        RowOut = RowOut + 1

    Next Node

    '------------------------------------
    ' Format
    '------------------------------------
    ws.Rows(1).Font.Bold = True

    ws.Columns.AutoFit

End Sub



'=================================================
' Function一覧作成
'=================================================
'
' @JPName
' Function一覧作成
'
' @Category
' ProcAnalyzer
'
' @Output
' FuncList(Dictionary)
'
' @Summary
' 対象ブック内の全プロシージャを収集し
' Dictionaryとして返却する
'
Public Function BuildFuncList() _
                As Object

    Dim FuncList As Object

    Dim VBComp As Object

    Dim CodeText As String

    Set FuncList = _
        CreateObject("Scripting.Dictionary")

    For Each VBComp In _
        gContext.TargetBook.VBProject.VBComponents
        
        If VBComp.CodeModule.CountOfLines > 0 Then

            CodeText = _
                VBComp.CodeModule.Lines( _
                    1, _
                    VBComp.CodeModule.CountOfLines)

            GetFunctions _
                CodeText, _
                VBComp.Name, _
                FuncList

        End If

    Next VBComp

    Set BuildFuncList = FuncList

End Function

Public Function GetProcedureCode( _
                ByVal ModuleName As String, _
                ByVal ProcName As String) _
                As String

    Dim VBComp As Object
    Dim CodeMod As Object

    For Each VBComp In _
        gContext.TargetBook.VBProject.VBComponents

        If StrComp( _
                VBComp.Name, _
                ModuleName, _
                vbTextCompare) = 0 Then

            Set CodeMod = _
                VBComp.CodeModule

            Exit For

        End If

    Next VBComp

    If CodeMod Is Nothing Then

        MsgBox _
            "モジュールが存在しません。" & _
            vbCrLf & _
            ModuleName, _
            vbExclamation

        Exit Function

    End If

    GetProcedureCode = _
        CodeMod.Lines( _
            CodeMod.ProcStartLine( _
                ProcName, _
                PROC_KIND), _
            CodeMod.ProcCountLines( _
                ProcName, _
                PROC_KIND))

End Function


Private Sub AddProcNode( _
                ByVal Nodes As Collection, _
                ByVal Node As clsProcNode, _
                ByRef PrevNode As clsProcNode, _
                ByRef LastNodeID As Long)

    Node.PrevNodeID = LastNodeID
    Node.NextNodeID = 0

    If Not PrevNode Is Nothing Then

        PrevNode.NextNodeID = _
            Node.NodeID

    End If

    Nodes.Add Node

    Set PrevNode = Node

    LastNodeID = Node.NodeID

End Sub




Public Function ExtractCalledProcedureName( _
                ByVal SourceLine As String) _
                As String

    Dim PosEqual As Long
    Dim PosSpace As Long
    Dim PosParen As Long

    SourceLine = Trim$(SourceLine)

    If SourceLine = "" Then Exit Function

    ' 代入文
    PosEqual = InStr(SourceLine, "=")

    If PosEqual > 0 Then

        SourceLine = _
            Trim$(Mid$( _
                SourceLine, _
                PosEqual + 1))

    End If

    ' モジュール名付き
    If InStr(SourceLine, " ") > 0 _
    And InStr(SourceLine, "(") > 0 Then

        PosSpace = InStr(SourceLine, " ")
        PosParen = InStr(SourceLine, "(")

        ' Sub 呼び出し
        If PosSpace < PosParen Then

            ExtractCalledProcedureName = _
                Left$(SourceLine, PosSpace - 1)

            GoTo NormalizeName

        End If

    End If

    PosParen = InStr(SourceLine, "(")

    If PosParen > 1 Then

        ExtractCalledProcedureName = _
            Left$(SourceLine, PosParen - 1)

    Else

        PosSpace = InStr(SourceLine, " ")

        If PosSpace > 0 Then

            ExtractCalledProcedureName = _
                Left$(SourceLine, PosSpace - 1)

        Else

            ExtractCalledProcedureName = _
                SourceLine

        End If

    End If

NormalizeName:

    If InStrRev( _
            ExtractCalledProcedureName, _
            ".") > 0 Then

        ExtractCalledProcedureName = _
            Mid$( _
                ExtractCalledProcedureName, _
                InStrRev( _
                    ExtractCalledProcedureName, _
                    ".") + 1)

    End If

End Function


