Attribute VB_Name = "modPathChart"
Option Explicit

#Const DEBUG_MODE = False

Private Const ROOT_CELL As String = "B3"

Private Const NODE_WIDTH As Long = 180
Private Const NODE_HEIGHT As Long = 40

Private Const H_GAP As Long = 260
'Private Const V_GAP As Long = 50
Private Const V_GAP As Long = 55

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

Public gNavigatorStatus As String

'==================================================
' PathChart生成
'==================================================

' @JPName
' パスチャート作成
'
' @Category
' PathChart
'
' @Summary
' 関数依存関係シートを参照し
' 指定関数を起点としたパスチャート図を生成する
'
Public Sub CreatePathChart_V3( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    Dim ws As Worksheet
    Dim RootShape As Shape
    Dim RootLeft As Double
    Dim RootCaption As String

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    gNodeNo = 0

    NodeCount = 0
    EdgeCount = 0
    MaxDepthFound = 0

    gCurrentNodeText = ""
    gNavigatorStatus = ""
    gCurrentShapeName = ""
    gCurrentParentShapeName = ""
    gCurrentParentWithChildren = ""
    gCurrentChildIndex = 0

    Set gParentMap = _
            CreateObject("Scripting.Dictionary")

    Set gChildMap = _
            CreateObject("Scripting.Dictionary")

    Set ws = ManagementBook.Worksheets("PathChart")

    If StartProc = "" Then

        MsgBox _
            "開始Procedureを指定してください。"

        Exit Sub

    End If

    gMaxLevel = MaxDepth

    If gMaxLevel <= 0 Then
        gMaxLevel = 99
    End If

    gExportTimeStamp = GetTimeStamp()

    Call ClearChartShapes(ws)

    RootLeft = ws.Range(ROOT_CELL).Left

#If DEBUG_MODE Then

    Debug.Print _
        "SetIssueToUI=" & _
        GetBranchSize("SetIssueToUI")

    Debug.Print _
        "get_MonshinFile=" & _
        GetBranchSize("get_MonshinFile")

    Debug.Print _
        "Get_MonshinInfo=" & _
        GetBranchSize("Get_MonshinInfo")

#End If

    '------------------------------------
    ' ルートノードキャプション生成
    '------------------------------------
    RootCaption = _
        BuildCaptionText(StartProc)

    Set RootShape = _
        DrawNode( _
            ws, _
            RootCaption, _
            50, _
            RootLeft, _
            RGB(93, 186, 255))

    DrawCallTreeShape _
            StartProc, _
            RootShape.Name, _
            1, _
            50, _
            RootLeft

    ActiveWindow.Zoom = 60

    If gPowerPoint Then

        Call ExportPathChartToPowerPoint( _
                StartProc, _
                MaxDepth)

    End If

    If gPngOutput Then

        Call ExportPathChartToPNG( _
                StartProc, _
                MaxDepth)

    End If

    Call ExportPathChartToText( _
            StartProc, _
            MaxDepth)

    NodeCount = gNodeNo

    Dim ChartRange As Range

    Set ChartRange = GetChartRange()

    ChartRange.Cells( _
        ChartRange.Rows.Count, _
        ChartRange.Columns.Count).Value = " "

End Sub

' @JPName
' ノード再帰描画
'
' @Category
' PathChart
'
' @Summary
' 関数依存関係を再帰的にたどり
' 子ノードおよび接続線を描画する
'
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

    Dim LastRow As Long
    Dim r As Long

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set wsDep = ManagementBook.Worksheets("関数依存関係")

    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    Dim BranchOffset As Long
    Dim BranchSize As Long

    BranchOffset = 0

    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            Dim ChildFunc As String
            Dim CallCount As Long

            Dim ParentModule As String
            Dim ChildModule As String
            Dim ModuleName As String

            Dim Proc As clsProcInfo

            Dim CaptionText As String

            ChildFunc = _
                Trim$(wsDep.Cells(r, 3).Value)

            '====================
            ' 再帰呼出除外
            '====================
            If Not gRecursive Then

                If StrComp( _
                        ParentFunc, _
                        ChildFunc, _
                        vbTextCompare) = 0 Then

                    GoTo NextChild

                End If

            End If

            CallCount = _
                Val(wsDep.Cells(r, 5).Value)

            ParentModule = _
                Trim$(wsDep.Cells(r, 6).Value)

            ChildModule = _
                Trim$(wsDep.Cells(r, 7).Value)

            ModuleName = ChildModule

            '====================
            ' 外部呼出除外
            '====================
            If Not gExternal Then

                If StrComp( _
                        ParentModule, _
                        ChildModule, _
                        vbTextCompare) <> 0 Then

                    GoTo NextChild

                End If

            End If

            '====================
            ' 共通関数除外
            '====================
            If gSkipCommon Then

                If IsCommonFunction(ChildFunc) Then

                    GoTo NextChild

                End If

            End If

            Dim ChildTop As Double
            Dim ChildLeft As Double

            BranchSize = _
                GetBranchSize(ChildFunc)

            ChildLeft = _
                LeftPos + H_GAP

            ChildTop = _
                TopPos + _
                BranchOffset

            BranchOffset = _
                BranchOffset + _
                (BranchSize * V_GAP)

            '====================
            ' Procedure情報取得
            '====================
            Set Proc = _
                GetProcInfo(ChildFunc)

            If Not Proc Is Nothing Then

                ModuleName = _
                    Proc.ModuleName

            End If

            '====================
            ' ノード表示文字列生成
            '====================
            CaptionText = _
                BuildCaptionText( _
                    ChildFunc, _
                    CallCount)

            '====================
            ' モジュール色設定
            '====================
            Dim FillColor As Long

            If gColorModule Then

                FillColor = _
                    GetModuleColor(ModuleName)

            Else

                FillColor = _
                    RGB(31, 94, 124)

            End If

            '====================
            ' 子ノード生成
            '====================
            Dim shpChild As Shape

            Set shpChild = _
                DrawNode( _
                    ManagementBook.Worksheets("PathChart"), _
                    CaptionText, _
                    ChildTop, _
                    ChildLeft, _
                    FillColor)

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

            '====================
            ' 接続線生成
            '====================
            ConnectShapes _
                ManagementBook.Worksheets("PathChart"), _
                ParentShapeName, _
                shpChild.Name

            EdgeCount = EdgeCount + 1

            '====================
            ' 再帰描画
            '====================
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

' @JPName
' 分岐サイズ取得
'
' @Category
' PathChart
'
' @Summary
' ノード配列計算用の分岐サイズを取得する
'
Private Function GetBranchSize( _
                ByVal ParentFunc As String) _
                As Long

    Dim wsDep As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim ChildFunc As String
    Dim size As Long

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Function

    End If

    Set wsDep = _
        ManagementBook.Worksheets("関数依存関係")

    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            ChildFunc = _
                Trim$(wsDep.Cells(r, 3).Value)

            size = _
                size + _
                GetBranchSize(ChildFunc)

        End If

    Next r

    If size = 0 Then

        GetBranchSize = 1

    Else

        GetBranchSize = size

    End If

End Function

'==================================================
' ノード作成
'==================================================

' @JPName
' ノード作成
'
' @Category
' PathChart
'
' @Input
' ws(Worksheet)
' CaptionText(String)
' TopPos(Double)
' LeftPos(Double)
' FillColor(Long)
'
' @Output
' Shape
'
' @Summary
' パスチャート用ノード図形を作成する
'
Private Function DrawNode( _
            ws As Worksheet, _
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
                NODE_WIDTH, _
                NODE_HEIGHT)

    shp.Name = "N_" & Format$(gNodeNo, "000000")

    shp.TextFrame.Characters.Text = CaptionText
    
    shp.TextFrame.Characters.Font.size = 9
    
    'shp.OnAction = "SelectNode"
    
    shp.OnAction = _
        "'" & ThisWorkbook.Name & "'!SelectNode"
    
    shp.Fill.ForeColor.RGB = FillColor
    
    shp.TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    
    shp.line.ForeColor.RGB = RGB(80, 80, 80)
    
    shp.line.Weight = 1.25

    Set DrawNode = shp

End Function

'==================================================
' ノード表示テキスト生成
'
' 形式
'   関数名 <モジュール名>
'   日本語関数名
'==================================================

' @JPName
' ノード表示文字列生成
'
' @Category
' PathChart
'
' @Input
' ProcName(String)
' CallCount(Long)
'
' @Output
' CaptionText(String)
'
' @Summary
' PathChartノードに表示する文字列を生成する
'
' @Remarks
' モジュール名表示および呼出回数表示は
' 設定値により切替える
'
Private Function BuildCaptionText( _
                ByVal ProcName As String, _
                Optional ByVal CallCount As Long = 0) _
                As String

    Dim Proc As clsProcInfo

    Dim CaptionText As String

    Set Proc = GetProcInfo(ProcName)

    CaptionText = ProcName

    If Not Proc Is Nothing Then

        If gShowModule Then

            CaptionText = _
                CaptionText & _
                " <" & _
                Proc.ModuleName & _
                ">"

        End If

        If Trim$(Proc.JPName) <> "" Then

            CaptionText = _
                CaptionText & vbLf & _
                Proc.JPName

        End If

    End If

    If gShowCount Then

        If CallCount > 0 Then

            CaptionText = _
                CaptionText & _
                " (" & CallCount & ")"

        End If

    End If

    BuildCaptionText = CaptionText

End Function

' @JPName
' モジュール色取得
'
' @Category
' PathChart
'
' @Summary
' モジュール名に対応する表示色を取得する
'
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

'==================================================
' 接続線
'==================================================

' @JPName
' 接続線作成
'
' @Category
' PathChart
'
' @Summary
' 親ノードと子ノード間の接続線を作成する
'
Private Sub ConnectShapes( _
            ws As Worksheet, _
            ByVal ParentName As String, _
            ByVal ChildName As String)

    Dim Con As Shape

    Set Con = ws.Shapes.AddConnector( _
                    msoConnectorElbow, _
                    0, 0, 100, 100)

    Con.ConnectorFormat.BeginConnect _
        ws.Shapes(ParentName), 4

    Con.ConnectorFormat.EndConnect _
        ws.Shapes(ChildName), 2

    On Error Resume Next

    Con.Name = _
        "C_" & ParentName & "_" & ChildName

    On Error GoTo 0

End Sub

'==================================================
' PathChart初期化
'==================================================

' @JPName
' チャート初期化
'
' @Category
' PathChart
'
' @Summary
' PathChart上の作成済みノードと接続線を削除する
'
Private Sub ClearChartShapes(ws As Worksheet)

    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1

        If Left$(ws.Shapes(i).Name, 2) = "N_" _
        Or Left$(ws.Shapes(i).Name, 2) = "C_" Then

            ws.Shapes(i).Delete

        End If

    Next i

End Sub

' @JPName
' チャート範囲取得
'
' @Category
' PathChart
'
' @Summary
' PathChart全体を含むセル範囲を取得する
'
Public Function GetChartRange() As Range

    Dim ws As Worksheet
    Dim shp As Shape

    Dim MaxRight As Double
    Dim MaxBottom As Double

    Dim LastCol As Long
    Dim LastRow As Long

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Function

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

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

'==================================================
' PathChart実行
'==================================================

' @JPName
' パスチャート実行
'
' @Category
' PathChart
'
' @Summary
' PathChart生成を実行し結果を表示する
'
Public Sub ExecutePathChart( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    Dim ws As Worksheet

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = _
        ManagementBook.Worksheets("PathChart")

    '--------------------------
    ' ログ
    '--------------------------

    frmPathChart.AddLog _
        "PathChart生成開始"

    '--------------------------
    ' PathChart生成
    '--------------------------

    CreatePathChart_V3 StartProc, MaxDepth

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

    ManagementBook.Worksheets("PathChart").Activate
    
    'frmPathChart.Hide

End Sub

'==================================================
' 検索
'==================================================

' @JPName
' ノード検索
'
' @Category
' PathChart
'
' @Summary
' 指定文字列でノードを検索する
'
Public Sub FindNode()

    'MsgBox "FindNode Start"
    
    FindNodeCore False

End Sub

' @JPName
' 次候補検索
'
' @Category
' PathChart
'
' @Summary
' 次の検索候補ノードを検索する
'
Public Sub FindNextNode()

    FindNodeCore True

End Sub

' @JPName
' ノード検索実行
'
' @Category
' PathChart
'
' @Summary
' PathChart内のノード検索処理を実行する
'
Private Sub FindNodeCore(ByVal IsNext As Boolean)

    Dim ws As Worksheet
    Dim shp As Shape

    Dim FindText As String
    Dim ExactMatch As Boolean

    Dim ShapeNo As Long

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

    FindText = Trim$(gSearchText)

    If FindText = "" Then

        MsgBox "検索文字列を入力してください。"
        Exit Sub

    End If

    ExactMatch = gExactMatch

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
                
                #If DEBUG_MODE Then
                    Debug.Print "Current=" & shp.Name
                #End If
                
                gCurrentShapeName = shp.Name
                               
                If Not gParentMap Is Nothing Then

                    If gParentMap.Exists(shp.Name) Then

                        gCurrentParentShapeName = _
                            gParentMap(shp.Name)

                        #If DEBUG_MODE Then
                            Debug.Print "Parent=" & _
                                    gCurrentParentShapeName
                        #End If
                        
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
                
                gCurrentNodeText = shp.TextFrame.Characters.Text
                
                UpdateCurrentNodeInfo _
                    GetProcedureNameFromNodeText( _
                        gCurrentNodeText)
        
                gLastShapeIndex = ShapeNo

                UpdateNavigatorStatus _
                    "検索結果 " & _
                    gCurrentHitNo & "/" & _
                    gHitCount & " : " & _
                    GetProcedureNameFromNodeText( _
                        shp.TextFrame.Characters.Text)
                
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

    'ActiveWindow.Zoom = 60

End Sub

' @JPName
' ノード一致判定
'
' @Category
' PathChart
'
' @Summary
' ノード文字列が検索条件に一致するか判定する
'
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

'==================================================
' ノード選択
'==================================================

' @JPName
' ノード選択
'
' @Category
' PathChart
'
' @Summary
' 選択されたノードを現在ノードとして設定する
'
Public Sub SelectNode()

    If gParentMap Is Nothing Then

        MsgBox _
            "パスチャート情報が初期化されています。" & vbCrLf & _
            "PathChartを再生成してください。"

        Exit Sub

    End If

    If gChildMap Is Nothing Then

        MsgBox _
            "パスチャート情報が初期化されています。" & vbCrLf & _
            "PathChartを再生成してください。"

        Exit Sub

    End If
    
    Dim ws As Worksheet
    Dim shp As Shape
    Dim ShapeName As String

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

    ShapeName = CStr(Application.Caller)

    On Error Resume Next

    Set shp = ws.Shapes(ShapeName)

    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub

    UpdateCurrentNodeInfo _
        GetProcedureNameFromNodeText( _
            shp.TextFrame.Characters.Text)

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
    
    HighlightNode shp

    UpdateNavigatorStatus _
        "Current : " & _
        GetProcedureNameFromNodeText( _
            shp.TextFrame.Characters.Text)
        
End Sub

'==================================================
' ノード表示文字列からProcedure名を取得
'
' 例:
'   Get_MonshinInfo <modMonshinCore>
'   問診情報取得
'
' → Get_MonshinInfo
'==================================================

' @JPName
' 関数名取得
'
' @Category
' PathChart
'
' @Input
' NodeText(String)
'
' @Output
' ProcedureName(String)
'
' @Summary
' ノード表示文字列から
' プロシージャ名を取得する
'
' @Remarks
' ナビゲータ表示および
' ノード検索で利用する
'
Public Function GetProcedureNameFromNodeText( _
                ByVal NodeText As String) _
                As String

    Dim ProcName As String

    ProcName = _
        Split(NodeText, vbLf)(0)

    ProcName = _
        Split(ProcName, " <")(0)

    GetProcedureNameFromNodeText = _
        Trim$(ProcName)

End Function

' @JPName
' ノード情報更新
'
' @Category
' PathChart
'
' @Summary
' 選択ノードの情報をNavigatorへ表示する
'
Public Sub UpdateCurrentNodeInfo( _
                ByVal ProcName As String)
    
    Dim Proc As clsProcInfo

    Set Proc = GetProcInfo(ProcName)

    On Error Resume Next

    If frmPathChartNavigator.Visible Then

        frmPathChartNavigator.txtCurrentNode.Text = _
            ProcName

        If Not Proc Is Nothing Then

            frmPathChartNavigator.lblJPName.Caption = _
                Proc.JPName

            frmPathChartNavigator.lblCategory.Caption = _
                Proc.Category

            frmPathChartNavigator.lblModule.Caption = _
                Proc.ModuleName
            
            frmPathChartNavigator.txtSummary.Text = _
                Proc.Summary

        Else

            frmPathChartNavigator.lblJPName.Caption = ""

            frmPathChartNavigator.lblCategory.Caption = ""
            
            frmPathChartNavigator.lblModule.Caption = ""

            frmPathChartNavigator.txtSummary.Text = ""

        End If

    End If

    On Error GoTo 0

End Sub

'==================================================
' ノード移動
'==================================================

' @JPName
' 親ノード移動
'
' @Category
' PathChart
'
' @Summary
' 現在ノードの親ノードへ移動する
'
Public Sub MoveToParentNode()

    If gParentMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If

    Dim ws As Worksheet
    Dim shp As Shape

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If
    
    Set ws = ManagementBook.Worksheets("PathChart")
    
    'MsgBox "MoveToParentNode"
    
    #If DEBUG_MODE Then
        Debug.Print "CurrentShape=" & shp.Name
        Debug.Print "CurrentParent=" & gCurrentParentShapeName
        Debug.Print "CurrentParentWithChildren=" & gCurrentParentWithChildren
    #End If
    
    If gCurrentParentShapeName = "" Then

        MsgBox _
            "親ノードはありません。", _
            vbInformation

        Exit Sub

    End If

    Set shp = _
        ws.Shapes(gCurrentParentShapeName)

    gCurrentShapeName = shp.Name

    gCurrentNodeText = _
        shp.TextFrame.Characters.Text

    If gParentMap.Exists(shp.Name) Then

        gCurrentParentShapeName = _
            gParentMap(shp.Name)

    Else

        gCurrentParentShapeName = ""

    End If

    gCurrentParentWithChildren = _
        gCurrentShapeName

    Call HighlightNode(shp)
    
    UpdateCurrentNodeInfo _
        GetProcedureNameFromNodeText( _
            gCurrentNodeText)
    
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
    
    UpdateNavigatorStatus _
            "親ノード : " & _
            GetProcedureNameFromNodeText( _
                shp.TextFrame.Characters.Text)
    
End Sub

' @JPName
' 子ノード移動
'
' @Category
' PathChart
'
' @Summary
' 現在ノードの子ノードへ移動する
'
Public Sub MoveToChildNode()

    If gChildMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If

    Dim ws As Worksheet
    Dim shp As Shape

    Dim Ary As Variant
    
    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If


    Set ws = ManagementBook.Worksheets("PathChart")

    If gCurrentParentWithChildren = "" Then

        MsgBox "基準ノードがありません。"
        
        Exit Sub

    End If

    If gChildMap Is Nothing Then
    
        MsgBox "パスチャートを再生成してください。"
        
        Exit Sub
        
    End If
    
    #If DEBUG_MODE Then
        Debug.Print "Current=" & gCurrentParentWithChildren
    #End If

    If Not gChildMap.Exists(gCurrentParentWithChildren) Then

        MsgBox "子ノードはありません。"
        
        Exit Sub

    End If
    
    #If DEBUG_MODE Then
        Debug.Print "ChildMap=" & gChildMap(gCurrentParentWithChildren)
    #End If

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
    
    gCurrentShapeName = shp.Name

    gCurrentNodeText = _
        shp.TextFrame.Characters.Text

    If Not gParentMap Is Nothing Then

        If gParentMap.Exists(shp.Name) Then

            gCurrentParentShapeName = _
                gParentMap(shp.Name)

        Else

            gCurrentParentShapeName = ""

        End If

    End If

    gCurrentParentWithChildren = shp.Name

    Call HighlightNode(shp)
   
    UpdateCurrentNodeInfo _
        GetProcedureNameFromNodeText( _
            gCurrentNodeText)
    
    gCurrentParentWithChildren = shp.Name

    If Not gParentMap Is Nothing Then
    
        If gParentMap.Exists(shp.Name) Then

            gCurrentParentWithChildren = _
                gParentMap(shp.Name)

        Else

            gCurrentParentWithChildren = ""

        End If
        
    End If

    UpdateNavigatorStatus _
        "子ノード " & _
        gCurrentChildIndex & "/" & _
        UBound(Ary) + 1 & _
        " : " & _
        GetProcedureNameFromNodeText( _
            shp.TextFrame.Characters.Text)

End Sub

' @JPName
' 兄弟ノード移動
'
' @Category
' PathChart
'
' @Summary
' 現在ノードと同一親を持つ兄弟ノードへ移動する
'
Public Sub MoveToSiblingNode()

    If gParentMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If
    
    Dim ws As Worksheet
    Dim shp As Shape

    Dim ParentName As String
    Dim Ary As Variant

    Dim i As Long
    Dim CurrentPos As Long
    
    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If


    Set ws = ManagementBook.Worksheets("PathChart")

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

    gCurrentShapeName = shp.Name

    gCurrentNodeText = shp.TextFrame.Characters.Text
    
    UpdateCurrentNodeInfo _
        GetProcedureNameFromNodeText( _
            gCurrentNodeText)

    If gParentMap.Exists(shp.Name) Then

        gCurrentParentShapeName = _
            gParentMap(shp.Name)
        
        gCurrentParentWithChildren = _
            gParentMap(shp.Name)

    Else

        gCurrentParentShapeName = ""
        
        gCurrentParentWithChildren = ""

    End If
    
    Call HighlightNode(shp)

    gCurrentParentWithChildren = _
        gCurrentParentShapeName


    UpdateNavigatorStatus _
        "兄弟ノード : " & _
        GetProcedureNameFromNodeText( _
            shp.TextFrame.Characters.Text)

End Sub

'==================================================
' ハイライト
'==================================================

' @JPName
' ノード強調表示
'
' @Category
' PathChart
'
' @Summary
' 現在ノードおよび関連ノードを強調表示する
'
Private Sub HighlightNode( _
                ByVal shp As Shape)

    ' 全解除
    Call ResetChartHighlight

    ' 現在ノード強調
    Call HighlightCurrentNode(shp)

    ' 親ノード強調
    Call HighlightParentNode( _
            gCurrentParentShapeName)

    ' 親コネクタ強調
    Call HighlightParentConnector
    
    ' 兄弟ノード強調
    Call HighlightSiblingNodes( _
        gCurrentParentShapeName)
    
    ' 兄弟コネクタ強調
    Call HighlightSiblingConnectors( _
            gCurrentParentWithChildren)
    
    ' 現在ノードの子ノード強調
    Call HighlightChildNodes( _
        gCurrentShapeName)

End Sub

' @JPName
' 現在ノード強調
'
' @Category
' PathChart
'
' @Summary
' 現在選択中ノードを強調表示する
'
Private Sub HighlightCurrentNode( _
                ByVal shp As Shape)

    With shp

        ' 現在ノードだけ赤
        .line.ForeColor.RGB = RGB(255, 0, 0)
        .line.Weight = 3

    End With

End Sub

' @JPName
' 親ノード強調
'
' @Category
' PathChart
'
' @Summary
' 親ノードを強調表示する
'
Private Sub HighlightParentNode( _
                ByVal ParentShapeName As String)

    Dim ws As Worksheet

    If gParentMap Is Nothing Then Exit Sub

    If ParentShapeName = "" Then Exit Sub

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If


    Set ws = ManagementBook.Worksheets("PathChart")

    With ws.Shapes(ParentShapeName)

        .line.ForeColor.RGB = RGB(0, 112, 192)
        .line.Weight = 3

        ' 背景色は変更しない

    End With

End Sub

' @JPName
' 親接続線強調
'
' @Category
' PathChart
'
' @Summary
' 親ノードとの接続線を強調表示する
'
Private Sub HighlightParentConnector()

    If gParentMap Is Nothing Then Exit Sub
    
    If gCurrentParentShapeName = "" Then
        Exit Sub
    End If

    ' 親→現在ノード へのコネクタ赤
    HighlightConnector _
        gCurrentParentShapeName, _
        gCurrentShapeName, _
        RGB(255, 0, 0), _
        2

End Sub

' @JPName
' 子ノード強調
'
' @Category
' PathChart
'
' @Summary
' 現在ノードの子ノードを強調表示する
'
Private Sub HighlightChildNodes( _
                ByVal ParentShapeName As String)

    If gChildMap Is Nothing Then Exit Sub
    
    If gChildMap.Exists(ParentShapeName) Then

    End If

    Dim ws As Worksheet
    Dim Ary As Variant

    Dim ChildName As Variant

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If


    Set ws = ManagementBook.Worksheets("PathChart")

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then Exit Sub

    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    For Each ChildName In Ary
    
        'Debug.Print "Child=" & ChildName
    
        If CStr(ChildName) <> gCurrentShapeName Then
    
        With ws.Shapes(CStr(ChildName))

            ' 子だけ薄黄色
            .line.ForeColor.RGB = RGB(255, 192, 0)

            .line.Weight = 2
            
            .Fill.ForeColor.RGB = RGB(255, 242, 204)
            
            .TextFrame.Characters.Font.Color = RGB(192, 0, 0)

        End With

        HighlightConnector _
            ParentShapeName, _
            CStr(ChildName), _
            RGB(255, 128, 0), _
            3
        
        End If

    Next ChildName

End Sub


' @JPName
' 兄弟ノード強調
'
' @Category
' PathChart
'
' @Summary
' 兄弟ノードを強調表示する
'
Private Sub HighlightSiblingNodes( _
                ByVal ParentShapeName As String)

    Dim ws As Worksheet
    Dim Ary As Variant
    Dim ChildName As Variant

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If


    Set ws = ManagementBook.Worksheets("PathChart")

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then
        Exit Sub
    End If

    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    For Each ChildName In Ary

        If CStr(ChildName) <> _
           gCurrentShapeName Then

            With ws.Shapes(CStr(ChildName))

                .line.ForeColor.RGB = _
                    RGB(112, 173, 71)

                .line.Weight = 2

                .Fill.ForeColor.RGB = _
                    RGB(226, 239, 218)

            End With

            HighlightConnector _
                ParentShapeName, _
                CStr(ChildName), _
                RGB(112, 173, 71), _
                2

        End If

    Next ChildName

End Sub


' @JPName
' 兄弟接続線強調
'
' @Category
' PathChart
'
' @Summary
' 兄弟ノードへの接続線を強調表示する
'
Private Sub HighlightSiblingConnectors( _
                ByVal ParentShapeName As String)

    Dim Ary As Variant
    Dim ChildName As Variant

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then
        Exit Sub
    End If

    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    For Each ChildName In Ary

        If CStr(ChildName) <> gCurrentShapeName Then

            ' 兄弟へのコネクタだけ薄緑
            HighlightConnector _
                ParentShapeName, _
                CStr(ChildName), _
                RGB(112, 173, 71), _
                2

        End If

    Next ChildName

End Sub

' @JPName
' 接続線強調
'
' @Category
' PathChart
'
' @Summary
' 指定ノード間の接続線を強調表示する
'
Private Sub HighlightConnector( _
                ByVal ParentShapeName As String, _
                ByVal ChildShapeName As String, _
                ByVal LineColor As Long, _
                ByVal LineWeight As Double)

    Dim ws As Worksheet
    Dim ConnectorName As String

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

    ConnectorName = _
        "C_" & _
        ParentShapeName & "_" & _
        ChildShapeName

    On Error Resume Next

    With ws.Shapes(ConnectorName)

        .line.ForeColor.RGB = LineColor
        .line.Weight = LineWeight
        
    End With

    On Error GoTo 0

End Sub


' @JPName
' 強調解除
'
' @Category
' PathChart
'
' @Summary
' PathChartの強調表示を初期状態へ戻す
'
Private Sub ResetChartHighlight()

    Dim ws As Worksheet
    Dim shp As Shape

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            shp.line.ForeColor.RGB = RGB(80, 80, 80)
            shp.line.Weight = 1.25

            shp.Fill.ForeColor.RGB = RGB(255, 255, 255)

            shp.TextFrame.Characters.Font.Color = RGB(0, 0, 0)

        End If

        If Left$(shp.Name, 2) = "C_" Then

            shp.line.ForeColor.RGB = RGB(120, 120, 120)
            shp.line.Weight = 1

        End If

    Next shp

End Sub


'==================================================
' ハイライト解除
'==================================================

' @JPName
' 強調表示解除
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChartの強調表示状態を解除する
'
' @Remarks
' 現在選択ノード情報および
' Navigator表示情報も初期化する
'
Public Sub ClearHighlight()

    Call ResetChartHighlight

    gCurrentShapeName = ""

    gCurrentParentShapeName = ""

    gCurrentParentWithChildren = ""

    gCurrentNodeText = ""

    gNavigatorStatus = ""

    Application.StatusBar = False

    UpdateNavigatorStatus ""

    frmPathChartNavigator.txtCurrentNode.Text = ""

    frmPathChartNavigator.lblJPName.Caption = ""

    frmPathChartNavigator.lblCategory.Caption = ""

    frmPathChartNavigator.lblModule.Caption = ""

    frmPathChartNavigator.txtSummary.Text = ""

End Sub

'==================================================
' Navigator
'==================================================

' @JPName
' Navigator状態更新
'
' @Category
' PathChart
'
' @Summary
' Navigatorおよびステータスバーへ状態を表示する
'
Public Sub UpdateNavigatorStatus( _
                ByVal Msg As String)

    gNavigatorStatus = Msg

    On Error Resume Next

    If frmPathChartNavigator.Visible Then

        frmPathChartNavigator.txtStatus.Text = _
            Msg

    End If

    On Error GoTo 0

    Application.StatusBar = Msg

End Sub


'==================================================
' 開始候補
'==================================================

' @JPName
' 開始関数読込
'
' @Category
' PathChart
'
' @Summary
' 開始関数候補をコンボボックスへ設定する
'
Public Sub LoadStartProcedure( _
                ByVal cbo As MSForms.ComboBox)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If
    
    Set Dic = CreateObject("Scripting.Dictionary")

    Set ws = _
        ManagementBook.Worksheets("関数依存関係")

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

'==================================================
' 検索履歴
'==================================================

' @JPName
' 検索履歴初期化
'
' @Category
' PathChart
'
' @Summary
' 検索履歴コレクションを初期化する
'
Public Sub InitSearchHistory()

    If gSearchHistory Is Nothing Then

        Set gSearchHistory = New Collection

    End If

End Sub

' @JPName
' 検索履歴追加
'
' @Category
' PathChart
'
' @Summary
' 検索履歴へキーワードを追加する
'
Public Sub AddSearchHistory( _
                ByVal SearchText As String)

    Dim V As Variant

    If Trim$(SearchText) = "" Then Exit Sub

    InitSearchHistory

    For Each V In gSearchHistory

        If StrComp(V, SearchText, _
                   vbTextCompare) = 0 Then
            Exit Sub
        End If

    Next

    gSearchHistory.Add SearchText

End Sub

' @JPName
' 検索履歴読込
'
' @Category
' PathChart
'
' @Summary
' 検索履歴をコンボボックスへ設定する
'
Public Sub LoadSearchHistory( _
                ByVal cbo As MSForms.ComboBox)

    Dim V As Variant

    InitSearchHistory

    For Each V In gSearchHistory

        cbo.AddItem V

    Next

End Sub

'==================================================
' 検索候補
'==================================================

' @JPName
' 検索候補読込
'
' @Category
' PathChart
'
' @Summary
' PathChart内のノード一覧を検索候補として読み込む
'
Public Sub LoadSearchNodeList( _
                ByVal cbo As MSForms.ComboBox)

    Dim ws As Worksheet
    Dim shp As Shape
    Dim Dic As Object
    Dim NodeText As String
    Dim V As Variant

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If

    Set ws = ManagementBook.Worksheets("PathChart")

    Set Dic = _
        CreateObject("Scripting.Dictionary")

    cbo.Clear

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then
        
        #If DEBUG_MODE Then
            Debug.Print shp.Name
            Debug.Print shp.TextFrame.Characters.Text
        #End If

            NodeText = _
                GetProcedureNameFromNodeText( _
                    shp.TextFrame.Characters.Text)

            NodeText = Trim$(NodeText)

            If NodeText <> "" Then

                If Not Dic.Exists(NodeText) Then

                    Dic.Add _
                        NodeText, _
                        True

                End If

            End If

        End If

    Next shp

    For Each V In Dic.Keys

        cbo.AddItem V

    Next V

End Sub

'==================================================
' ショートカット
'==================================================

' @JPName
' ショートカット登録
'
' @Category
' PathChart
'
' @Summary
' PathChart操作用ショートカットキーを登録する
'
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

' @JPName
' ショートカット解除
'
' @Category
' PathChart
'
' @Summary
' PathChart操作用ショートカットキーを解除する
'
Public Sub UnRegisterShortcut()

    Application.OnKey "^+F"

    Application.OnKey "^+N"
    
    Application.OnKey "^+P"

    Application.OnKey "^+C"
    
    Application.OnKey "^+S"
    
    Application.OnKey "^+H"

End Sub

'==================================================
' 出力
'==================================================

' @JPName
' PowerPoint出力
'
' @Category
' PathChart
'
' @Summary
' PathChartをPowerPointへ出力する
'
Private Sub ExportPathChartToPowerPoint( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    Call modPowerPointChart.CreatePowerPointPathChart( _
            StartProc, _
            MaxDepth)

End Sub

