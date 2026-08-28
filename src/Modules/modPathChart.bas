Attribute VB_Name = "modPathChart"
Option Explicit

Private gNodeNo As Long
Private gMaxLevel As Long

Private gCurrentShapeName As String
Private gCurrentParentShapeName As String

Private gCurrentParentWithChildren As String
Private gCurrentChildIndex As Long

Private gHitCount As Long
Private gCurrentHitNo As Long

Private gLastFindText As String
Private gLastShapeIndex As Long

Private gParentMap As Object
Private gChildMap As Object

Public gExportTimeStamp As String

Public NodeCount As Long
Public EdgeCount As Long
Public MaxDepthFound As Long


Private Function IsCommonFunction( _
                    ByVal FuncName As String) As Boolean

    Select Case UCase$(FuncName)

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

' パスチャート生成

Public Sub CreatePathChart_V3()

    Dim ws As Worksheet
    Dim StartFunc As String
    Dim RootShape As Shape
    Dim RootTop As Double
    Dim RootLeft As Double

    gNodeNo = 0
    
    NodeCount = 0
    EdgeCount = 0
    MaxDepthFound = 0
    
    Set gParentMap = _
            CreateObject("Scripting.Dictionary")

    Set gChildMap = _
            CreateObject("Scripting.Dictionary")
    
    Set ws = Worksheets("PathChart")

    StartFunc = Trim$(ws.Range(CONFIG_START_FUNC).Value)

    If StartFunc = "" Then

        MsgBox "開始関数を指定してください。"
        Exit Sub

    End If

    gMaxLevel = Val(ws.Range(CONFIG_MAX_LEVEL).Value)

    If gMaxLevel <= 0 Then
        gMaxLevel = 99
    End If

    gSkipCommon = _
        (UCase$(Trim$(ws.Range(CONFIG_SKIP_COMMON).Value)) = "Y")
        
    gShowCount = _
        (UCase$(Trim$(ws.Range(CONFIG_SHOW_COUNT).Value)) = "Y")
    
    gPowerPoint = _
        (UCase$(Trim$(ws.Range(CONFIG_POWERPOINT).Value)) = "Y")
    
    gPngOutput = _
        (UCase$(Trim$(ws.Range(CONFIG_PNG_OUTPUT).Value)) = "Y")
    
    gShowModule = _
        (UCase$(Trim$(ws.Range(CONFIG_SHOW_MODULE).Value)) = "Y")
    
    gColorModule = _
        (UCase$(Trim$(ws.Range(CONFIG_COLOR_MODULE).Value)) = "Y")
        
    gExportTimeStamp = GetTimeStamp()

    Call ClearChartShapes(ws)

    RootTop = ws.Range(ROOT_CELL).Top
    RootLeft = ws.Range(ROOT_CELL).Left
    
    Set RootShape = DrawNode( _
                        ws, _
                        StartFunc, _
                        StartFunc, _
                        50, _
                        RootLeft, _
                        RGB(0, 112, 192))

    DrawCallTreeShape _
            StartFunc, _
            RootShape.Name, _
            1, _
            RootTop, _
            RootLeft

    ActiveWindow.Zoom = 60

    If gPowerPoint Then

        ExportPathChartToPowerPoint

    End If
    
    If gPngOutput Then
    
        ExportPathChartToPNG
        
    End If
    
    'Debug.Print gNodeNo
    
    NodeCount = gNodeNo

    'MsgBox "パスチャート作成完了"
    
End Sub

Private Sub DrawCallTree( _
            ByVal ParentFunc As String, _
            ByRef OutRow As Long, _
            ByVal Level As Long)

    Dim wsDep As Worksheet

    Set wsDep = ThisWorkbook.Worksheets("関数依存関係")

    Dim LastRow As Long
    Dim r As Long

    LastRow = wsDep.Cells(wsDep.Rows.Count, "A").End(xlUp).row

    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            Dim ChildFunc As String

            ChildFunc = Trim$(wsDep.Cells(r, 2).Value)

            If ChildFunc <> "" Then

                Worksheets("PathChart").Cells(OutRow, Level + 1).Value = _
                    "└─ " & ChildFunc

                OutRow = OutRow + 1

                Call DrawCallTree( _
                        ChildFunc, _
                        OutRow, _
                        Level + 1)

            End If

        End If

    Next r

End Sub

' ノード作成
Private Function DrawNode( _
            ws As Worksheet, _
            ByVal NodeName As String, _
            ByVal CaptionText As String, _
            ByVal TopPos As Double, _
            ByVal LeftPos As Double, _
            ByVal FillColor As Long) As Shape

    Dim shp As Shape
    
    ' DrawNode制限
    If gMaxNode > 0 Then

        If gNodeNo >= gMaxNode Then

            Exit Function

        End If

    End If

    gNodeNo = gNodeNo + 1

    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                LeftPos, _
                TopPos, _
                180, _
                30)

    shp.Name = "N_" & NodeName & "_" & gNodeNo

    shp.TextFrame.Characters.Text = CaptionText
    
    shp.TextFrame.Characters.Font.Size = 9
    
    shp.OnAction = "SelectNode"
    
    shp.Fill.ForeColor.RGB = FillColor
    
    shp.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    
    shp.line.ForeColor.RGB = RGB(80, 80, 80)
    
    shp.line.Weight = 1.25

    Set DrawNode = shp

End Function

' 再帰描画
Private Sub DrawCallTreeShape( _
            ByVal ParentFunc As String, _
            ByVal ParentShapeName As String, _
            ByVal Level As Long, _
            ByVal TopPos As Double, _
            ByVal LeftPos As Double)
    
    If Level >= gMaxLevel Then Exit Sub
    
    If Level > MaxDepthFound Then

        MaxDepthFound = Level

    End If
    
    If gMaxNode > 0 Then

        If gNodeNo >= gMaxNode Then Exit Sub

    End If

    If gMaxEdge > 0 Then

        If EdgeCount >= gMaxEdge Then Exit Sub

    End If

    Dim wsDep As Worksheet
    
    Set wsDep = Worksheets("関数依存関係")
    
    Dim LastRow As Long
    Dim r As Long

    LastRow = wsDep.Cells( _
                wsDep.Rows.Count, _
                "A").End(xlUp).row

    Dim ChildIndex As Long

    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            Dim ChildFunc As String
            Dim CallCount As Long
            Dim ModuleName As String

            ChildFunc = wsDep.Cells(r, 2).Value
            CallCount = wsDep.Cells(r, 3).Value
            ModuleName = wsDep.Cells(r, 4).Value

            If gSkipCommon Then

                If IsCommonFunction(ChildFunc) Then

                    GoTo NextChild

                End If

            End If

            ChildIndex = ChildIndex + 1
            
            Dim ChildTop As Double
            Dim ChildLeft As Double

            '===============
            ' 縦方向へ並べる
            '===============

            ChildLeft = LeftPos + 220

            ChildTop = TopPos + _
                       ((ChildIndex - 1) * 60)

            Dim CaptionText As String

            CaptionText = ChildFunc

            If gShowCount Then

                CaptionText = _
                    CaptionText & _
                    " (" & CallCount & ")"

            End If

            If gShowModule Then

                CaptionText = _
                    CaptionText & _
                    "[" & ModuleName & "]"

            End If

            Dim FillColor As Long

            If gColorModule Then

                FillColor = _
                    GetModuleColor(ModuleName)

            Else

                    FillColor = RGB(31, 94, 124)

            End If

            Dim shpChild As Shape

            Set shpChild = DrawNode( _
                    Worksheets("PathChart"), _
                    ChildFunc, _
                    CaptionText, _
                    ChildTop, _
                    ChildLeft, _
                    FillColor)

            ' 親子関係保存
            gParentMap(shpChild.Name) = _
                    ParentShapeName
            
            If gChildMap.Exists(ParentShapeName) Then

                gChildMap(ParentShapeName) = _
                    gChildMap(ParentShapeName) & _
                    "|" & _
                    shpChild.Name

            Else

                gChildMap.Add _
                    ParentShapeName, _
                    shpChild.Name

            End If

            'Debug.Print "MAP:" & shpChild.Name & " -> " & ParentShapeName
            
            If gMaxEdge > 0 Then

                If EdgeCount >= gMaxEdge Then Exit Sub

            End If
                      
            ConnectShapes _
                    Worksheets("PathChart"), _
                    ParentShapeName, _
                    shpChild.Name

            EdgeCount = EdgeCount + 1
            
            DrawCallTreeShape _
                    ChildFunc, _
                    shpChild.Name, _
                    Level + 1, _
                    ChildTop, _
                    ChildLeft

        End If

NextChild:

    Next r

End Sub

' コネクタ
Private Sub ConnectShapes( _
            ws As Worksheet, _
            ByVal ParentName As String, _
            ByVal ChildName As String)

    Dim Con As Shape

    Set Con = ws.Shapes.AddConnector( _
                    msoConnectorElbow, _
                    0, 0, 100, 100)

    Con.ConnectorFormat.BeginConnect _
        ws.Shapes(ParentName), 3

    Con.ConnectorFormat.EndConnect _
        ws.Shapes(ChildName), 1

    Con.RerouteConnections

    On Error Resume Next
    Con.Name = _
        "C_" & ParentName & "_" & ChildName
    On Error GoTo 0

End Sub

' 初期化
Private Sub ClearChartShapes(ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1

        If Left$(ws.Shapes(i).Name, 2) = "N_" _
        Or Left$(ws.Shapes(i).Name, 2) = "C_" Then

            ws.Shapes(i).Delete

        End If

    Next i

End Sub

Private Function GetModuleColor( _
                    ByVal ModuleName As String) As Long

    Select Case UCase$(ModuleName)

        Case "MODMONSHINCORE"
            GetModuleColor = RGB(184, 204, 228)

        Case "MODISSUESERVICE"
            GetModuleColor = RGB(198, 224, 180)

        Case "MODISSUEMAPPER"
            GetModuleColor = RGB(248, 203, 173)

        Case "MODINFODISPLAY"
            GetModuleColor = RGB(217, 210, 233)

        Case "MODUTILITY"
            GetModuleColor = RGB(230, 230, 230)

        Case "XMLLOADER"
            GetModuleColor = RGB(255, 242, 204)

        Case Else
            GetModuleColor = RGB(200, 200, 200)

    End Select

End Function

Private Function ShapeExists( _
            ws As Worksheet, _
            ShapeName As String) As Boolean

    Dim shp As Shape

    For Each shp In ws.Shapes

        If shp.Name = ShapeName Then

            ShapeExists = True
            Exit Function

        End If

    Next shp

End Function

Public Sub FindNode()

    'MsgBox "FindNode Start"
    
    FindNodeCore False

End Sub

Public Sub FindNextNode()

    FindNodeCore True

End Sub

Private Sub FindNodeCore(ByVal IsNext As Boolean)

    Dim ws As Worksheet
    Dim shp As Shape

    Dim FindText As String
    Dim ExactMatch As Boolean

    Dim ShapeNo As Long
    Dim ZoomValue As Long

    Set ws = Worksheets("PathChart")

    FindText = Trim$( _
        ws.OLEObjects("ComboBoxSearch") _
          .Object.Value)

    If FindText = "" Then

        MsgBox "検索文字列を入力してください。"
        Exit Sub

    End If

    ExactMatch = _
        (UCase$(Trim$(ws.Range(CELL_EXACT_MATCH).Value)) = "Y")

    ZoomValue = Val(ws.Range(CELL_ZOOM).Value)

    If ZoomValue <= 0 Then
        ZoomValue = 100
    End If

    '=========================
    ' 新規検索ならリセット
    '=========================
    If Not IsNext Then

        gLastShapeIndex = 0
        gCurrentHitNo = 0

    ElseIf gLastFindText <> FindText Then

        gLastShapeIndex = 0
        gCurrentHitNo = 0

    End If

    gLastFindText = FindText

    '=========================
    ' 前回ハイライト解除
    '=========================
    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            shp.line.ForeColor.RGB = RGB(80, 80, 80)
            shp.line.Weight = 1.25

        End If

    Next shp

    '=========================
    ' 総件数を数える
    '=========================
    gHitCount = 0

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            If IsMatchNode( _
                    shp.TextFrame.Characters.Text, _
                    FindText, _
                    ExactMatch) Then

                gHitCount = gHitCount + 1

            End If

        End If

    Next shp

    If gHitCount = 0 Then

        Application.StatusBar = False

        MsgBox "検索結果はありません。"

        Exit Sub

    End If

    '=========================
    ' 検索
    '=========================
    ShapeNo = 0

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            ShapeNo = ShapeNo + 1

            If ShapeNo <= gLastShapeIndex Then
                GoTo ContinueLoop
            End If

            If IsMatchNode( _
                    shp.TextFrame.Characters.Text, _
                    FindText, _
                    ExactMatch) Then
                
                Debug.Print "Current=" & shp.Name
                
                gCurrentShapeName = shp.Name
                               
                If Not gParentMap Is Nothing Then

                    If gParentMap.Exists(shp.Name) Then

                        gCurrentParentShapeName = _
                            gParentMap(shp.Name)

                        Debug.Print "Parent=" & _
                                gCurrentParentShapeName
                        
                    Else

                        gCurrentParentShapeName = ""

                    End If

                End If
                
                
                ' 子巡回用基準ノード
                gCurrentParentWithChildren = shp.Name
                
                gCurrentChildIndex = 0
                
                gCurrentHitNo = gCurrentHitNo + 1

                ' ハイライト
                Call HighlightNode(shp)

                gLastShapeIndex = ShapeNo

                Application.StatusBar = _
                    "検索結果 " & _
                    gCurrentHitNo & "/" & _
                    gHitCount & " : " & _
                    shp.TextFrame.Characters.Text
                
                Exit Sub

            End If

        End If

ContinueLoop:

    Next shp

    '=========================
    ' 最後まで到達
    '=========================
    gLastShapeIndex = 0
    gCurrentHitNo = 0
    gLastFindText = ""

    Application.StatusBar = False

    ' ハイライト解除
    Call ClearHighlight
    
    MsgBox _
        "これ以上見つかりません。" & vbCrLf & _
        "先頭から再検索します。", _
        vbInformation

    ActiveWindow.Zoom = 60

    Application.Goto _
        ws.Range(CELL_NAVIGATION_TOP), _
        True

End Sub

Public Sub MoveToParentNode()

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = Worksheets("PathChart")
    
    'MsgBox "MoveToParentNode"
    
    Debug.Print "CurrentShape=" & gCurrentShapeName
    Debug.Print "CurrentParent=" & gCurrentParentShapeName

    If gCurrentParentShapeName = "" Then

        MsgBox _
            "親ノードはありません。", _
            vbInformation

        Exit Sub

    End If

    Set shp = _
        ws.Shapes(gCurrentParentShapeName)
    
    Call HighlightNode(shp)
    
    gCurrentShapeName = shp.Name
    
    ' 子巡回の基準を親へ変更
    gCurrentParentWithChildren = shp.Name
    
    gCurrentChildIndex = 0
    
    If Not gParentMap Is Nothing Then
    
        If gParentMap.Exists(shp.Name) Then
        
            gCurrentParentShapeName = _
                gParentMap(shp.Name)
        Else
        
            gCurrentParentShapeName = ""
            
        End If
        
    End If
    
    Application.StatusBar = _
            "親ノード : " & _
            shp.TextFrame.Characters.Text
    
End Sub

Public Sub MoveToChildNode()

    Dim ws As Worksheet
    Dim shp As Shape

    Dim Ary As Variant
    
    Set ws = Worksheets("PathChart")

    If gCurrentParentWithChildren = "" Then

        MsgBox "基準ノードがありません。"
        
        Exit Sub

    End If

    If gChildMap Is Nothing Then
    
        MsgBox "パスチャートを再生成してください。"
        
        Exit Sub
        
    End If
    
    Debug.Print "Current=" & gCurrentParentWithChildren

    If Not gChildMap.Exists(gCurrentParentWithChildren) Then

        MsgBox "子ノードはありません。"
        
        Exit Sub

    End If
    
    Debug.Print "ChildMap=" & gChildMap(gCurrentParentWithChildren)

    Ary = Split( _
            gChildMap(gCurrentParentWithChildren), _
            "|")

    gCurrentChildIndex = _
        gCurrentChildIndex + 1

    If gCurrentChildIndex > UBound(Ary) + 1 Then

        gCurrentChildIndex = 1

    End If

    Set shp = _
        ws.Shapes( _
        Ary(gCurrentChildIndex - 1))

    Call HighlightNode(shp)

    gCurrentParentWithChildren = shp.Name

    If Not gParentMap Is Nothing Then
    
        If gParentMap.Exists(shp.Name) Then

            gCurrentParentWithChildren = _
                gParentMap(shp.Name)

        Else

            gCurrentParentWithChildren = ""

        End If
        
    End If

    Application.StatusBar = _
        "子ノード " & _
        gCurrentChildIndex & "/" & _
        UBound(Ary) + 1 & _
        " : " & _
        shp.TextFrame.Characters.Text

End Sub

Public Sub MoveToSiblingNode()

    Dim ws As Worksheet
    Dim shp As Shape

    Dim ParentName As String
    Dim Ary As Variant

    Dim i As Long
    Dim CurrentPos As Long
    
    Set ws = Worksheets("PathChart")

    If gCurrentShapeName = "" Then

        MsgBox "現在ノードがありません。"
        Exit Sub

    End If

    If gParentMap Is Nothing Then

        MsgBox "パスチャートを再生成してください。"
        Exit Sub

    End If

    If Not gParentMap.Exists(gCurrentShapeName) Then

        MsgBox "兄弟ノードはありません。"
        Exit Sub

    End If

    ParentName = gParentMap(gCurrentShapeName)

    If Not gChildMap.Exists(ParentName) Then

        MsgBox "兄弟ノードはありません。"
        Exit Sub

    End If

    Ary = Split(gChildMap(ParentName), "|")

    CurrentPos = -1

    For i = LBound(Ary) To UBound(Ary)

        If Ary(i) = gCurrentShapeName Then

            CurrentPos = i
            Exit For

        End If

    Next i

    If CurrentPos = -1 Then
    
        MsgBox "兄弟ノードが見つかりません。"
        Exit Sub
    End If

    CurrentPos = CurrentPos + 1

    If CurrentPos > UBound(Ary) Then

        CurrentPos = 0

    End If

    Set shp = ws.Shapes(Ary(CurrentPos))

    Call HighlightNode(shp)

    gCurrentShapeName = shp.Name

    If gParentMap.Exists(shp.Name) Then

        gCurrentParentShapeName = _
            gParentMap(shp.Name)

    Else

        gCurrentParentShapeName = ""

    End If

    Application.StatusBar = _
        "兄弟ノード : " & _
        shp.TextFrame.Characters.Text

End Sub

Private Function IsMatchNode( _
                ByVal NodeText As String, _
                ByVal FindText As String, _
                ByVal ExactMatch As Boolean) _
                As Boolean

    If ExactMatch Then

        Dim BaseText As String

        BaseText = Split(NodeText, "(")(0)
        BaseText = Split(BaseText, "[")(0)

        BaseText = Trim$(BaseText)

        IsMatchNode = _
            (StrComp(BaseText, _
                     FindText, _
                     vbTextCompare) = 0)

    Else

        IsMatchNode = _
            (InStr(1, _
                   NodeText, _
                   FindText, _
                   vbTextCompare) > 0)

    End If

End Function

Private Sub AddSearchHistory(ByVal SearchText As String)

    Dim cbo As Object
    Dim i As Long

    Set cbo = _
        Worksheets("PathChart") _
        .OLEObjects("ComboBoxSearch") _
        .Object

    For i = 0 To cbo.ListCount - 1

        If cbo.List(i) = SearchText Then
            cbo.Value = SearchText
            Exit Sub
        End If

    Next i

    cbo.AddItem SearchText

    cbo.Value = SearchText

End Sub

Public Sub ClearSearchHistory()

    Worksheets("PathChart") _
        .OLEObjects("ComboBoxSearch") _
        .Object.Clear

    Application.StatusBar = False

End Sub

' ハイライト処理
Private Sub HighlightNode( _
                ByVal shp As Shape)

    Dim s As Shape
    Dim ws As Worksheet

    Set ws = Worksheets("PathChart")

    For Each s In ws.Shapes

        If Left$(s.Name, 2) = "N_" Then

            s.line.ForeColor.RGB = RGB(80, 80, 80)
            s.line.Weight = 1.25

        End If

    Next s

    shp.line.ForeColor.RGB = RGB(255, 0, 0)

    shp.line.Weight = 3

    ' 移動
    Application.Goto _
        ws.Cells( _
            shp.TopLeftCell.row, _
            shp.TopLeftCell.Column), _
        True

End Sub

' ハイライト解除
Public Sub ClearHighlight()

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = Worksheets("PathChart")

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            shp.line.ForeColor.RGB = RGB(80, 80, 80)
            shp.line.Weight = 1.25

        End If

    Next shp

    Application.StatusBar = False

    gCurrentShapeName = ""
    gCurrentParentShapeName = ""

End Sub

Public Sub RegisterShortcut()

    ' 検索 Ctrl + Shift + F
    Application.OnKey "^+F", "FindNode"

    ' 次検索 Ctrl + Shift + N
    Application.OnKey "^+N", "FindNextNode"
    
    ' 親へ移動 Ctrl + Shift + P
    Application.OnKey "^+P", "MoveToParentNode"
    
    ' 子へ移動 Ctrl + Shift + C
    Application.OnKey "^+C", "MoveToChildNode"
    
    ' 兄弟へ移動 Ctrl + Shift + S
    Application.OnKey "^+S", "MoveToSiblingNode"
    
    ' ハイライト解除 Ctrl + Shift + H
    Application.OnKey "^+H", "ClearHighlight"

End Sub

Public Sub UnRegisterShortcut()

    Application.OnKey "^+F"

    Application.OnKey "^+N"
    
    Application.OnKey "^+P"

    Application.OnKey "^+C"
    
    Application.OnKey "^+S"
    
    Application.OnKey "^+H"

End Sub

Private Sub ExportPathChartToPowerPoint()

    'MsgBox _
    '    "PowerPoint出力機能は未実装です。", _
    '    vbInformation
    
    Call modPowerPointChart.CreatePowerPointPathChart

End Sub

' ノード選択処理
Public Sub SelectNode()

    Dim ws As Worksheet
    Dim shp As Shape

    Dim ShapeName As String

    Set ws = Worksheets("PathChart")

    ShapeName = Application.Caller

    Set shp = ws.Shapes(ShapeName)

    HighlightNode shp

    gCurrentShapeName = shp.Name

    gCurrentParentWithChildren = shp.Name

    gCurrentChildIndex = 0

    If Not gParentMap Is Nothing Then

        If gParentMap.Exists(shp.Name) Then

            gCurrentParentShapeName = _
                gParentMap(shp.Name)

        Else

            gCurrentParentShapeName = ""

        End If

    End If

    Application.StatusBar = _
        "Current : " & _
        shp.TextFrame.Characters.Text
        
    ws.Range(CELL_DUMMY_SELECT).Select
    
    Debug.Print "Select=" & shp.Name

    If Not gParentMap Is Nothing Then

        Debug.Print _
            "Exists=" & _
            gParentMap.Exists(shp.Name)

        If gParentMap.Exists(shp.Name) Then

            Debug.Print _
                "Parent=" & _
                gParentMap(shp.Name)

        End If

    End If
End Sub

Public Function GetChartRange() As Range

    Dim ws As Worksheet
    Dim shp As Shape

    Dim MaxRight As Double
    Dim MaxBottom As Double

    Dim LastCol As Long
    Dim LastRow As Long

    Set ws = Worksheets("PathChart")

    MaxRight = ws.Range(ROOT_CELL).Left
    MaxBottom = ws.Range(ROOT_CELL).Top

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" _
        Or Left$(shp.Name, 2) = "C_" Then

            If shp.Left + shp.Width > MaxRight Then

                MaxRight = shp.Left + shp.Width

            End If

            If shp.Top + shp.Height > MaxBottom Then

                MaxBottom = shp.Top + shp.Height

            End If

        End If

    Next shp

    LastCol = 1

    Do While ws.Columns(LastCol).Left < MaxRight

        LastCol = LastCol + 1

    Loop

    LastRow = 1

    Do While ws.Rows(LastRow).Top < MaxBottom

        LastRow = LastRow + 1

    Loop

    Set GetChartRange = _
        ws.Range( _
            ws.Range(ROOT_CELL), _
            ws.Cells(LastRow + 1, _
                     LastCol + 1))

End Function

Public Sub ExecutePathChart( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    Dim ws As Worksheet

    Set ws = _
        ThisWorkbook.Worksheets("PathChart")

    '--------------------------
    ' PathChartシートへ転記
    '--------------------------

    ws.Range(CONFIG_START_FUNC).Value = _
        StartProc

    ws.Range(CONFIG_MAX_LEVEL).Value = _
        MaxDepth

    frmPathChart.AddLog _
        "PathChartシートへOption転記"

    ApplyPathChartOption

    '--------------------------
    ' ログ
    '--------------------------

    frmPathChart.AddLog _
        "PathChart生成開始"

    '--------------------------
    ' PathChart生成
    '--------------------------

    CreatePathChart_V3

    '--------------------------
    ' 結果表示
    '--------------------------

    frmPathChart.AddLog _
        "PathChart生成完了"

    frmPathChart.lblNodeCount.Caption = _
        "Node数: " & NodeCount

    frmPathChart.lblEdgeCount.Caption = _
        "Edge数: " & EdgeCount

    frmPathChart.lblDepth.Caption = _
        "最大深度: " & MaxDepthFound

    frmPathChart.AddLog _
        "Node数 : " & NodeCount

    frmPathChart.AddLog _
        "Edge数 : " & EdgeCount

    frmPathChart.AddLog _
        "最大深度 : " & MaxDepthFound

    ThisWorkbook.Worksheets("PathChart").Activate
    
    frmPathChart.Hide

End Sub

Public Sub ClearPathChart()

End Sub

Public Sub DrawPathChart()

End Sub

Private Sub ApplyPathChartOption()

    With Worksheets("PathChart")

        .Range(CONFIG_SKIP_COMMON).Value = _
            IIf(gSkipCommon, "Y", "N")

        .Range(CONFIG_SHOW_COUNT).Value = _
            IIf(gShowCount, "Y", "N")

        .Range(CONFIG_POWERPOINT).Value = _
            IIf(gPowerPoint, "Y", "N")

        .Range(CONFIG_PNG_OUTPUT).Value = _
            IIf(gPngOutput, "Y", "N")

        .Range(CONFIG_SHOW_MODULE).Value = _
            IIf(gShowModule, "Y", "N")

        .Range(CONFIG_COLOR_MODULE).Value = _
            IIf(gColorModule, "Y", "N")

    End With

End Sub

Public Sub LoadStartProcedure( _
                ByVal cbo As MSForms.ComboBox)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object

    Set Dic = CreateObject("Scripting.Dictionary")

    Set ws = _
        ThisWorkbook.Worksheets("関数依存関係")

    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    cbo.Clear

    For r = 2 To LastRow

        If Trim$(ws.Cells(r, 1).Value) <> "" Then

            If Not Dic.Exists( _
                    ws.Cells(r, 1).Value) Then

                Dic.Add _
                    ws.Cells(r, 1).Value, _
                    True

                cbo.AddItem _
                    ws.Cells(r, 1).Value

            End If

        End If

    Next r

End Sub

