Attribute VB_Name = "modPathChart"
Option Explicit

' PathChartデバッグ出力有効化
#Const DEBUG_MODE = False

'==================================================
' PathChartレイアウト定数
'==================================================

' ルートノード Top
Private Const ROOT_TOP As Long = 50

' ルートノード配置セル
Private Const ROOT_CELL As String = "B3"

' ノード幅
Private Const NODE_WIDTH As Long = 180

' ノード高さ
Private Const NODE_HEIGHT As Long = 40

' 水平ノード間隔
Private Const H_GAP As Long = 260

' 垂直ノード間隔
Private Const V_GAP As Long = 55

'==================================================
' 表示定数
'==================================================

' 画面表示定数
Private Const DEFAULT_ZOOM As Long = 60

'==================================================
' ノード表示
'==================================================

' 図形のフォントサイズ
Private Const NODE_FONT_SIZE As Long = 8

' 図形の文字色 RGB(0, 0, 0)
Private Const NODE_FONT_COLOR As Long = &H0

' 図形の枠線色 RGB(80, 80, 80)
Private Const NODE_LINE_COLOR As Long = &H505050

' 図形の枠線の太さ
Private Const NODE_LINE_WEIGHT As Double = 1.25

' ノード背景色 RGB(255,255,255)
Private Const NODE_FILL_COLOR As Long = &HFFFFFF


'==================================================
' ハイライト表示
'==================================================

' [現在ノード]
' 現在ノード枠線色 RGB(255,0,0)
Private Const CURRENT_NODE_LINE_COLOR As Long = &HFF

' 現在ノード図形の枠線の太さ
Private Const CURRENT_NODE_LINE_WEIGHT As Double = 3

' [親ノード]
' 親ノード枠線色 RGB(0,112,192)
Private Const PARENT_NODE_LINE_COLOR As Long = &HC07000

' [子ノード]
' 子ノード枠線色 RGB(255,192,0)
Private Const CHILD_LINE_COLOR As Long = &HC0FF

' 子ノード図形の枠線の太さ
Private Const CHILD_LINE_WEIGHT As Double = 2

' 子ノード背景色 RGB(255,242,204)
Private Const CHILD_FILL_COLOR As Long = &HCCF2FF

' 子ノード文字色 RGB(192,0,0)
Private Const CHILD_FONT_COLOR As Long = &HC0

' 子ノード接続線色 RGB(255,128,0)
Private Const CHILD_CONNECTOR_COLOR As Long = &H80FF

' [兄弟ノード]
' 兄弟ノード枠線色 RGB(112,173,71)
Private Const SIBLING_LINE_COLOR As Long = &H47AD70

' 兄弟ノード背景色 RGB(226,239,218)
Private Const SIBLING_FILL_COLOR As Long = &HDAEFE2

'==================================================
' コネクタ表示
'==================================================

' コネクタ標準色 RGB(120,120,120)
Private Const CONNECTOR_LINE_COLOR As Long = &H787878

' コネクタ標準太さ
Private Const CONNECTOR_LINE_WEIGHT As Double = 1

' 親接続線太さ
Private Const PARENT_CONNECTOR_WEIGHT As Double = 2

' 子接続線太さ
Private Const CHILD_CONNECTOR_WEIGHT As Double = 3

' 兄弟接続線太さ
Private Const SIBLING_CONNECTOR_WEIGHT As Double = 2

'==================================================
' ノード配色
'==================================================

' ルートノード色 RGB(93, 186, 255)
Private Const ROOT_NODE_COLOR As Long = &HFFBA5D

' 標準ノード色 RGB(31, 94, 124)
Private Const DEFAULT_NODE_COLOR As Long = &H7C5E1F

' modMonshinCore RGB(184,204,228)
Private Const MODULE_COLOR_MONSHIN As Long = &HE4CCB8

' modIssueService RGB(198,224,180)
Private Const MODULE_COLOR_ISSUE As Long = &HB4E0C6

' modIssueMapper RGB(248,203,173)
Private Const MODULE_COLOR_MAPPER As Long = &HADCBF8

' modInfoDisplay RGB(217,210,233)
Private Const MODULE_COLOR_INFO As Long = &HE9D2D9

' modUtility RGB(230,230,230)
Private Const MODULE_COLOR_UTILITY As Long = &HE6E6E6

' XmlLoader RGB(255,242,204)
Private Const MODULE_COLOR_XML As Long = &HCCF2FF

' 既定ノード色 RGB(200,200,200)
Private Const MODULE_COLOR_DEFAULT As Long = &HC8C8C8

'==================================================
' PathChart生成状態
'==================================================

' 現在採番済みノード数
Private gNodeNo As Long

' 探索上限階層
Private gMaxLevel As Long

'==================================================
' ナビゲーション状態
'==================================================

' 現在選択中ノード
Private gCurrentShapeName As String

' 現在選択中親ノード
Private gCurrentParentShapeName As String

' 子巡回基準ノード
Private gCurrentParentWithChildren As String

' 現在子インデックス
Private gCurrentChildIndex As Long

'==================================================
' 検索状態
'==================================================

' 検索ヒット総件数
Private gHitCount As Long

' 現在ヒット番号
Private gCurrentHitNo As Long

' 前回検索文字列
Private gLastFindText As String

' 前回検索Shape位置
Private gLastShapeIndex As Long

'==================================================
' ノード関係管理
'==================================================

' 子→親対応表
Private gParentMap As Object

' 親→子対応表
Private gChildMap As Object

'==================================================
' 出力情報
'==================================================

' 出力ファイル用タイムスタンプ
Public gExportTimeStamp As String

'==================================================
' 統計情報
'==================================================

' ノード総数
Public NodeCount As Long

' 接続線総数
Public EdgeCount As Long

' 発見最大階層
Public MaxDepthFound As Long

'==================================================
' UI状態
'==================================================

' Navigator表示メッセージ
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
' @Input
' StartProc(String)
' MaxDepth(Long)
'
' @Output
' なし
'
' @Summary
' 関数依存関係シートを参照し
' 指定関数を起点としたパスチャート図を生成する
'
' @Remarks
' パスチャート生成前に内部状態を初期化する
' ノード生成後にPowerPoint、PNG、
' テキスト出力を実行する
'
Public Sub CreatePathChart_V3( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' ルートノード図形
    Dim RootShape As Shape
    
    ' ルートノードX座標
    Dim RootLeft As Double
    
    ' ルート表示文字列
    Dim RootCaption As String

    '------------------------------------
    '# 状態初期化
    '------------------------------------
    '# PathChart状態初期化
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

    '------------------------------------
    '# 親子関係管理初期化
    '------------------------------------
    Set gParentMap = _
            CreateObject("Scripting.Dictionary")

    Set gChildMap = _
            CreateObject("Scripting.Dictionary")

    '------------------------------------
    '# シート準備
    '------------------------------------
    '# PathChartシート取得
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 開始Procedure確認
    If StartProc = "" Then

        MsgBox _
            "開始Procedureを指定してください。"

        Exit Sub

    End If

    '# 探索最大深度設定
    gMaxLevel = MaxDepth

    If gMaxLevel <= 0 Then
        gMaxLevel = 99
    End If

    '# 出力タイムスタンプ生成
    gExportTimeStamp = GetTimeStamp()

    '------------------------------------
    '# PathChart初期化
    '------------------------------------
    Call ClearChartShapes(ws)

    RootLeft = ws.Range(ROOT_CELL).Left

#If DEBUG_MODE Then

    Debug.Print _
        "SetIssueToUI=" & GetBranchSize("SetIssueToUI")

    Debug.Print _
        "get_MonshinFile=" & GetBranchSize("get_MonshinFile")

    Debug.Print _
        "Get_MonshinInfo=" & GetBranchSize("Get_MonshinInfo")

#End If

    '------------------------------------
    '# ルートノード表示文字列生成
    '------------------------------------
    RootCaption = _
        BuildCaptionText(StartProc)

    '# ルートノード生成
    Set RootShape = _
        DrawNode( _
            ws, _
            RootCaption, _
            ROOT_TOP, _
            RootLeft, _
            ROOT_NODE_COLOR)

    '# パスチャート再帰生成
    DrawCallTreeShape _
            StartProc, _
            RootShape.Name, _
            1, _
            ROOT_TOP, _
            RootLeft

    ActiveWindow.Zoom = DEFAULT_ZOOM

    '------------------------------------
    '# 出力処理
    '------------------------------------
    '# PowerPoint出力
    If gPowerPoint Then

        Call ExportPathChartToPowerPoint( _
                StartProc, _
                MaxDepth)

    End If

    '# PNG出力
    If gPngOutput Then

        Call ExportPathChartToPNG( _
                StartProc, _
                MaxDepth)

    End If

    '# パスチャートテキスト出力
    Call ExportPathChartToText( _
            StartProc, _
            MaxDepth)

    NodeCount = gNodeNo

    ' 出力範囲
    Dim ChartRange As Range

    '# 出力範囲取得
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
' @Input
' ParentFunc(String)
' ParentShapeName(String)
' Level(Long)
' TopPos(Double)
' LeftPos(Double)
' PathText(String)
'
' @Output
' なし
'
' @Summary
' 関数依存関係を再帰的にたどり
' 子ノードおよび接続線を描画する
'
' @Remarks
' 再帰ループを検出した場合は探索しない
' 外部呼出や共通関数除外条件を考慮する
' 親子関係情報を内部管理表へ登録する
'
Private Sub DrawCallTreeShape( _
            ByVal ParentFunc As String, _
            ByVal ParentShapeName As String, _
            ByVal Level As Long, _
            ByVal TopPos As Double, _
            ByVal LeftPos As Double, _
            Optional ByVal PathText As String = "")

    ' 関数依存関係シート
    Dim wsDep As Worksheet

    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 子ノードオフセット
    Dim BranchOffset As Long
    
    ' 分岐サイズ
    Dim BranchSize As Long

    '# 再帰判定用パス生成
    If PathText = "" Then

        PathText = _
            "|" & UCase$(ParentFunc) & "|"

    End If

    '# 探索深度上限制御
    If Level >= gMaxLevel Then Exit Sub

    '# 最大深度更新
    If Level > MaxDepthFound Then

        MaxDepthFound = Level

    End If

    '# ノード数上限制御
    If gMaxNode > 0 Then

        If gNodeNo >= gMaxNode Then Exit Sub

    End If

    '# 接続線数上限制御
    If gMaxEdge > 0 Then

        If EdgeCount >= gMaxEdge Then Exit Sub

    End If

    '# 関数依存関係取得
    Set wsDep = _
        gContext.ManagementBook.Worksheets(SHEET_FUNCTION_DEPENDENCY)

    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    BranchOffset = 0

    '# 呼出先探索
    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            ' === 呼出先情報 ===
            ' 呼出先関数
            Dim ChildFunc As String
            
            ' 呼出回数
            Dim CallCount As Long

            ' 呼出元モジュール
            Dim ParentModule As String
            
            ' 呼出先モジュール
            Dim ChildModule As String
            
            ' 表示対象モジュール名
            Dim ModuleName As String

            ' Procedure情報
            Dim Proc As clsProcInfo

            ' ノード表示文字列
            Dim CaptionText As String

            ' 子ノードY座標
            Dim ChildTop As Double
            
            ' 子ノードX座標
            Dim ChildLeft As Double

            ' ノード色
            Dim FillColor As Long

            ' 子ノードShape
            Dim shpChild As Shape

            '# 呼出先情報取得
            ChildFunc = _
                Trim$(wsDep.Cells(r, 3).Value)

            '====================
            '# 再帰ループ防止
            '====================
            If InStr( _
                    1, _
                    PathText, _
                    "|" & UCase$(ChildFunc) & "|", _
                    vbTextCompare) > 0 Then

                GoTo NextChild

            End If

            '====================
            '# 再帰呼出除外
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
            '# 外部呼出除外
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
            '# 共通関数除外
            '====================
            If gSkipCommon Then

                If IsCommonFunction(ChildFunc) Then

                    GoTo NextChild

                End If

            End If

            '====================
            '# 子ノード配置座標計算
            '====================
            '# 分岐サイズ取得
            BranchSize = _
                GetBranchSize(ChildFunc)

            '# 子ノード配置座標計算
            ChildLeft = _
                LeftPos + H_GAP

            ChildTop = _
                TopPos + BranchOffset

            BranchOffset = _
                BranchOffset + _
                (BranchSize * V_GAP)

            '====================
            '# Procedure情報取得
            '====================
            Set Proc = _
                ResolveProcedure( _
                    ParentModule, _
                    ChildFunc)

            If Not Proc Is Nothing Then

                ModuleName = _
                    Proc.ModuleName

            End If

            '====================
            '# ノード表示文字列生成
            '====================
            CaptionText = _
                BuildCaptionText( _
                    ChildFunc, _
                    CallCount)

            '====================
            '# モジュール色設定
            '====================
            If gColorModule Then

                FillColor = _
                    GetModuleColor(ModuleName)

            Else

                FillColor = _
                    DEFAULT_NODE_COLOR

            End If

            '====================
            '# 子ノード生成
            '====================
            Set shpChild = _
                DrawNode( _
                    gContext.ManagementBook.Worksheets(SHEET_PATH_CHART), _
                    CaptionText, _
                    ChildTop, _
                    ChildLeft, _
                    FillColor)

            If shpChild Is Nothing Then

                GoTo NextChild

            End If

            '# 親子関係登録
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
            '# 接続線生成
            '====================
            ConnectShapes _
                gContext.ManagementBook.Worksheets(SHEET_PATH_CHART), _
                ParentShapeName, _
                shpChild.Name

            EdgeCount = EdgeCount + 1

            '====================
            '# 子ノード再帰生成
            '====================
            DrawCallTreeShape _
                ChildFunc, _
                shpChild.Name, _
                Level + 1, _
                ChildTop, _
                ChildLeft, _
                PathText & _
                UCase$(ChildFunc) & "|"

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
' @Input
' ParentFunc(String)
' PathText(String)
'
' @Output
' BranchSize(Long)
'
' @Summary
' ノード配列計算用の分岐サイズを取得する
'
' @Remarks
' 子孫ノード数を再帰的に集計する
' ノード配置時の縦方向サイズ算出に利用する
'
Private Function GetBranchSize( _
                ByVal ParentFunc As String, _
                Optional ByVal PathText As String = "") _
                As Long

    ' 関数依存関係シート
    Dim wsDep As Worksheet

    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 呼出先関数
    Dim ChildFunc As String
    
    ' 分岐サイズ
    Dim Size As Long

    '# 再帰判定用パス生成
    If PathText = "" Then

        PathText = _
            "|" & UCase$(ParentFunc) & "|"

    End If

    '# 関数依存関係取得
    Set wsDep = _
        gContext.ManagementBook.Worksheets(SHEET_FUNCTION_DEPENDENCY)

    LastRow = _
        wsDep.Cells( _
            wsDep.Rows.Count, _
            "A").End(xlUp).row

    '====================
    '# 子孫ノード数集計
    '====================
    For r = 2 To LastRow

        If wsDep.Cells(r, 1).Value = ParentFunc Then

            ChildFunc = _
                Trim$(wsDep.Cells(r, 3).Value)

            '====================
            ' 再帰ループ防止
            '====================
            If InStr( _
                    1, _
                    PathText, _
                    "|" & UCase$(ChildFunc) & "|", _
                    vbTextCompare) > 0 Then

                GoTo NextChild

            End If

            '# 分岐サイズ加算
            Size = _
                Size + _
                GetBranchSize( _
                    ChildFunc, _
                    PathText & _
                    UCase$(ChildFunc) & "|")

        End If

NextChild:

    Next r

    '# 分岐サイズ返却
    If Size = 0 Then

        GetBranchSize = 1

    Else

        GetBranchSize = Size

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
' @Remarks
' ノード名の採番を行う
' 図形書式およびクリックイベントを設定する
'
Private Function DrawNode( _
            ws As Worksheet, _
            ByVal CaptionText As String, _
            ByVal TopPos As Double, _
            ByVal LeftPos As Double, _
            ByVal FillColor As Long) As Shape
    
    ' 作成Shape
    Dim shp As Shape
    
    '# ノード生成上限確認
    If gMaxNode > 0 Then

        If gNodeNo >= gMaxNode Then

            Exit Function

        End If

    End If

    '# ノード採番
    gNodeNo = gNodeNo + 1

    '# ノード図形生成
    Set shp = ws.Shapes.AddShape( _
                msoShapeRoundedRectangle, _
                LeftPos, _
                TopPos, _
                NODE_WIDTH, _
                NODE_HEIGHT)

    shp.Name = "N_" & Format$(gNodeNo, "000000")

    '# ノード表示文字列設定
    shp.TextFrame.Characters.Text = CaptionText
    
    shp.TextFrame.Characters.Font.Size = NODE_FONT_SIZE
    
    'shp.OnAction = "SelectNode"
    
    '# ノード選択イベント設定
    shp.OnAction = _
        "'" & ThisWorkbook.Name & "'!SelectNode"
    
    '# ノード書式設定
    shp.Fill.ForeColor.RGB = FillColor
    
    shp.TextFrame.Characters.Font.Color = NODE_FONT_COLOR
    
    shp.line.ForeColor.RGB = NODE_LINE_COLOR
    
    shp.line.Weight = NODE_LINE_WEIGHT

    '# ノード返却
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

    ' Procedure情報
    Dim Proc As clsProcInfo

    ' ノード表示文字列
    Dim CaptionText As String

    '# Procedure情報取得
    Set Proc = _
        ResolveProcedure( _
            "", _
            ProcName)

    '# ノード表示文字列初期化
    CaptionText = ProcName

    If Not Proc Is Nothing Then

        '# モジュール名表示
        If gShowModule Then

            CaptionText = _
                CaptionText & _
                " <" & _
                Proc.ModuleName & _
                ">"

        End If

        '# JPName表示
        If Trim$(Proc.JPName) <> "" Then

            CaptionText = _
                CaptionText & vbLf & _
                Proc.JPName

        End If

    End If

    '# 呼出回数表示
    If gShowCount Then

        If CallCount > 0 Then

            CaptionText = _
                CaptionText & _
                " (" & CallCount & ")"

        End If

    End If

    '# ノード表示文字列返却
    BuildCaptionText = CaptionText

End Function

' @JPName
' モジュール色取得
'
' @Category
' PathChart
'
' @Input
' ModuleName(String)
'
' @Output
' FillColor(Long)
'
' @Summary
' モジュール名に対応する表示色を取得する
'
' @Remarks
' モジュール種別ごとの固定色を返却する
' 未定義モジュールは既定色を返却する
'
Private Function GetModuleColor( _
                    ByVal ModuleName As String) As Long

    Select Case UCase$(ModuleName)

        Case "MODMONSHINCORE"
            GetModuleColor = MODULE_COLOR_MONSHIN

        Case "MODISSUESERVICE"
            GetModuleColor = MODULE_COLOR_ISSUE

        Case "MODISSUEMAPPER"
            GetModuleColor = MODULE_COLOR_MAPPER

        Case "MODINFODISPLAY"
            GetModuleColor = MODULE_COLOR_INFO

        Case "MODUTILITY"
            GetModuleColor = MODULE_COLOR_UTILITY

        Case "XMLLOADER"
            GetModuleColor = MODULE_COLOR_XML

        Case Else
            GetModuleColor = MODULE_COLOR_DEFAULT

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
' @Input
' ws(Worksheet)
' ParentName(String)
' ChildName(String)
'
' @Output
' なし
'
' @Summary
' 親ノードと子ノード間の接続線を作成する
'
' @Remarks
' エルボコネクタを使用する
' 接続線名は親子ノード名から生成する
'
Private Sub ConnectShapes( _
            ws As Worksheet, _
            ByVal ParentName As String, _
            ByVal ChildName As String)

    ' 接続線Shape
    Dim Con As Shape

    '# 接続線生成
    Set Con = ws.Shapes.AddConnector( _
                    msoConnectorElbow, _
                    0, 0, 100, 100)

    '# 接続線書式設定
    Con.line.ForeColor.RGB = _
        CONNECTOR_LINE_COLOR

    Con.line.Weight = _
        CONNECTOR_LINE_WEIGHT
    
    '# 親ノード接続
    Con.ConnectorFormat.BeginConnect _
        ws.Shapes(ParentName), 4

    '# 子ノード接続
    Con.ConnectorFormat.EndConnect _
        ws.Shapes(ChildName), 2

    On Error Resume Next

    '# 接続線名称設定
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
' @Input
' ws(Worksheet)
'
' @Output
' なし
'
' @Summary
' PathChart上の作成済みノードと接続線を削除する
'
' @Remarks
' ノード(N_)および接続線(C_)のみ削除する
'
Private Sub ClearChartShapes(ws As Worksheet)

    ' Shape走査Index
    Dim i As Long

    '# PathChart図形削除
    For i = ws.Shapes.Count To 1 Step -1

        '# PathChartノード・接続線削除
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
' @Input
' なし
'
' @Output
' ChartRange(Range)
'
' @Summary
' PathChart全体を含むセル範囲を取得する
'
' @Remarks
' ノードおよび接続線の配置位置を走査する
' Chart全体を含むセル範囲を返却する
'
Public Function GetChartRange() As Range
    
    ' PathChartシート
    Dim ws As Worksheet
    
    ' Shape
    Dim shp As Shape

    ' 最大右端座標
    Dim MaxRight As Double
    
    ' 最大下端座標
    Dim MaxBottom As Double

    ' 最終列
    Dim LastCol As Long
    
    ' 最終行
    Dim LastRow As Long

    '# PathChartシート取得
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 描画範囲初期化
    MaxRight = ws.Range(ROOT_CELL).Left
    MaxBottom = ws.Range(ROOT_CELL).Top

    '# 図形配置範囲取得
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

    '# 最終列算出
    Do While ws.Columns(LastCol).Left < MaxRight

        LastCol = LastCol + 1

    Loop

    LastRow = 1

    '# 最終行算出
    Do While ws.Rows(LastRow).Top < MaxBottom

        LastRow = LastRow + 1

    Loop

    '# PathChart範囲返却
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
' @Input
' StartProc(String)
' MaxDepth(Long)
'
' @Output
' なし
'
' @Summary
' PathChart生成を実行し結果を表示する
'
' @Remarks
' PathChart生成処理実行後に
' 統計情報およびログを更新する
'
Public Sub ExecutePathChart( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    ' PathChartシート
    Dim ws As Worksheet

    '# PathChartシート取得
    Set ws = _
        gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '--------------------------
    ' ログ
    '--------------------------

    '# PathChart生成開始ログ出力
    frmPathChart.AddLog _
        "PathChart生成開始"

    '--------------------------
    ' PathChart生成
    '--------------------------

    '# PathChart生成実行
    CreatePathChart_V3 StartProc, MaxDepth

    '--------------------------
    ' 結果表示
    '--------------------------

    '# PathChart生成完了ログ出力
    frmPathChart.AddLog _
        "PathChart生成完了"

    '# 統計情報表示更新
    frmPathChart.lblNodeCount.Caption = _
        "Node数: " & NodeCount

    frmPathChart.lblEdgeCount.Caption = _
        "Edge数: " & EdgeCount

    frmPathChart.lblDepth.Caption = _
        "最大深度: " & MaxDepthFound

    '# 統計情報ログ出力
    frmPathChart.AddLog _
        "Node数 : " & NodeCount

    frmPathChart.AddLog _
        "Edge数 : " & EdgeCount

    frmPathChart.AddLog _
        "最大深度 : " & MaxDepthFound

    '# PathChartシート表示
    gContext.ManagementBook.Worksheets(SHEET_PATH_CHART).Activate
    
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 指定文字列でノードを検索する
'
' @Remarks
' 新規検索として検索状態を初期化し
' 検索処理を実行する
'
Public Sub FindNode()

    'MsgBox "FindNode Start"
    
    '# 新規検索実行
    FindNodeCore False

End Sub

' @JPName
' 次候補検索
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
' 次の検索候補ノードを検索する
'
' @Remarks
' 前回検索位置以降から検索を継続する
'
Public Sub FindNextNode()

    '# 次候補検索実行
    FindNodeCore True

End Sub

' @JPName
' ノード検索実行
'
' @Category
' PathChart
'
' @Input
' IsNext(Boolean)
'
' @Output
' なし
'
' @Summary
' PathChart内のノード検索処理を実行する
'
' @Remarks
' 検索条件に一致したノードを探索する
' 一致ノードを強調表示し
' Navigator情報を更新する
'
Private Sub FindNodeCore(ByVal IsNext As Boolean)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' ノードShape
    Dim shp As Shape

    ' 検索文字列
    Dim FindText As String
    
    ' 完全一致検索フラグ
    Dim ExactMatch As Boolean

    ' Shape通番
    Dim ShapeNo As Long

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 検索条件取得
    FindText = Trim$(gSearchText)

    '# 検索条件確認
    If FindText = "" Then

        MsgBox "検索文字列を入力してください。"
        Exit Sub

    End If

    ExactMatch = gExactMatch

    '=========================
    ' 新規検索ならリセット
    '=========================
    '# 検索状態初期化
    If Not IsNext Then

        gLastShapeIndex = 0
        gCurrentHitNo = 0

    ElseIf gLastFindText <> FindText Then

        gLastShapeIndex = 0
        gCurrentHitNo = 0

    End If

    gLastFindText = FindText

    '=========================
    '# 前回ハイライト解除
    '=========================
    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            'shp.line.ForeColor.RGB = RGB(80, 80, 80)
            'shp.line.Weight = 1.25
            shp.line.ForeColor.RGB = NODE_LINE_COLOR
            shp.line.Weight = NODE_LINE_WEIGHT

        End If

    Next shp

    '=========================
    ' 総件数を数える
    '=========================
    gHitCount = 0

    '# 検索ヒット件数集計
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

    '# ヒット件数確認
    If gHitCount = 0 Then

        Application.StatusBar = False

        MsgBox "検索結果はありません。"

        Exit Sub

    End If

    '=========================
    ' 検索
    '=========================
    ShapeNo = 0

    '# 検索開始
    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            ShapeNo = ShapeNo + 1

            If ShapeNo <= gLastShapeIndex Then
                GoTo ContinueLoop
            End If

            '# 一致ノード検出
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

                '# ノード強調表示
                Call HighlightNode(shp)
                
                gCurrentNodeText = shp.TextFrame.Characters.Text
                
                '# Navigator状態更新
                UpdateCurrentNodeInfo _
                    GetProcedureNameFromNodeText( _
                        gCurrentNodeText)
        
                '# 検索状態リセット
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

End Sub

' @JPName
' ノード一致判定
'
' @Category
' PathChart
'
' @Input
' NodeText(String)
' FindText(String)
' ExactMatch(Boolean)
'
' @Output
' IsMatchNode(Boolean)
'
' @Summary
' ノード文字列が検索条件に一致するか判定する
'
' @Remarks
' 完全一致検索および部分一致検索を
' 指定に応じて切り替えて判定する
'
Private Function IsMatchNode( _
                ByVal NodeText As String, _
                ByVal FindText As String, _
                ByVal ExactMatch As Boolean) _
                As Boolean

    If ExactMatch Then

        ' 判定用ノード文字列
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択されたノードを現在ノードとして設定する
'
' @Remarks
' クリックされたノードを現在ノードとして設定する
' ハイライト表示およびNavigator更新を行う
'
Public Sub SelectNode()

    '# 親子関係情報確認
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
    
    ' PathChartシート
    Dim ws As Worksheet
    
    ' 選択ノード
    Dim shp As Shape
    
    ' Shape名
    Dim ShapeName As String

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 選択Shape取得
    ShapeName = CStr(Application.Caller)

    On Error Resume Next

    Set shp = ws.Shapes(ShapeName)

    On Error GoTo 0
    
    If shp Is Nothing Then Exit Sub

    '# Procedure情報表示更新
    UpdateCurrentNodeInfo _
            GetProcedureNameFromNodeText( _
                shp.TextFrame.Characters.Text)

    '# 現在ノード設定
    gCurrentShapeName = shp.Name

    gCurrentParentWithChildren = shp.Name

    gCurrentChildIndex = 0

    If Not gParentMap Is Nothing Then

        '# 親ノード取得
        If gParentMap.Exists(shp.Name) Then

            gCurrentParentShapeName = _
                gParentMap(shp.Name)

        Else

            gCurrentParentShapeName = ""

        End If

    End If
    
    '# ノード強調表示
    HighlightNode shp

    '# Navigator状態更新
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

    ' Procedure名
    Dim ProcName As String

    '# 表示文字列先頭行取得
    ProcName = Split(NodeText, vbLf)(0)

    '# モジュール名除去
    ProcName = Split(ProcName, " <")(0)

    '# Procedure名返却
    GetProcedureNameFromNodeText = Trim$(ProcName)

End Function

' @JPName
' ノード情報更新
'
' @Category
' PathChart
'
' @Input
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 選択ノードの情報をNavigatorへ表示する
'
' @Remarks
' Procedure情報を取得し
' Navigatorの表示内容を更新する
'
Public Sub UpdateCurrentNodeInfo( _
                ByVal ProcName As String)
    
    ' Procedure情報
    Dim Proc As clsProcInfo

    '# Procedure情報取得
    Set Proc = _
        ResolveProcedure( _
            "", _
            ProcName)
    
    On Error Resume Next

    '# Navigator表示更新
    If frmPathChartNavigator.Visible Then

        frmPathChartNavigator.txtCurrentNode.Text = ProcName

        '# Procedure情報表示
        If Not Proc Is Nothing Then

            frmPathChartNavigator.lblJPName.Caption = Proc.JPName

            frmPathChartNavigator.lblCategory.Caption = Proc.Category

            frmPathChartNavigator.lblModule.Caption = Proc.ModuleName
            
            frmPathChartNavigator.txtSummary.Text = Proc.Summary

        '# Navigator表示クリア
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 現在ノードの親ノードへ移動する
'
' @Remarks
' 現在ノードの親ノードへ移動する
' ナビゲーション状態および表示情報を更新する
'
Public Sub MoveToParentNode()

    '# 親ノード存在確認
    If gParentMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If

    ' PathChartシート
    Dim ws As Worksheet
    
    ' 移動先ノード
    Dim shp As Shape

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)
    
    'MsgBox "MoveToParentNode"
       
    '# 親ノード有無確認
    If gCurrentParentShapeName = "" Then

        MsgBox _
            "親ノードはありません。", _
            vbInformation

        Exit Sub

    End If

    '# 親ノード取得
    Set shp = _
        ws.Shapes(gCurrentParentShapeName)

    '# 現在ノード更新
    gCurrentShapeName = shp.Name

    gCurrentNodeText = shp.TextFrame.Characters.Text

    If gParentMap.Exists(shp.Name) Then

        gCurrentParentShapeName = gParentMap(shp.Name)

    Else

        gCurrentParentShapeName = ""

    End If

    gCurrentParentWithChildren = gCurrentShapeName

    '# ノード強調表示
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
        
            gCurrentParentShapeName = gParentMap(shp.Name)
        Else
        
            gCurrentParentShapeName = ""
            
        End If
        
    End If
    
    '# Navigator状態更新
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 現在ノードの子ノードへ移動する
'
' @Remarks
' 現在ノード配下の子ノードを巡回する
' 子ノード選択状態および表示情報を更新する
'
Public Sub MoveToChildNode()

    If gChildMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If

    ' PathChartシート
    Dim ws As Worksheet
    
    ' 移動先ノード
    Dim shp As Shape

    ' 子ノード一覧
    Dim Ary As Variant
    
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    If gCurrentParentWithChildren = "" Then

        MsgBox "基準ノードがありません。"
        
        Exit Sub

    End If

    '# 子ノード管理確認
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

    '# 子ノード一覧取得
    Ary = Split( _
            gChildMap(gCurrentParentWithChildren), _
            "|")
    
    '--------------------------
    '# 子ノード巡回
    '--------------------------

    '# 次子ノード選択
    gCurrentChildIndex = gCurrentChildIndex + 1

    If gCurrentChildIndex > UBound(Ary) + 1 Then

        gCurrentChildIndex = 1

    End If

    '--------------------------
    '# 子ノード取得
    '--------------------------
    Set shp = _
        ws.Shapes( _
        Ary(gCurrentChildIndex - 1))
    
    '--------------------------
    '# 現在ノード更新
    '--------------------------
    gCurrentShapeName = shp.Name

    gCurrentNodeText = _
        shp.TextFrame.Characters.Text

    If Not gParentMap Is Nothing Then

        If gParentMap.Exists(shp.Name) Then

            gCurrentParentShapeName = gParentMap(shp.Name)

        Else

            gCurrentParentShapeName = ""

        End If

    End If

    gCurrentParentWithChildren = shp.Name

    '# ノード強調表示
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

    '# Navigator状態更新
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 現在ノードと同一親を持つ兄弟ノードへ移動する
'
' @Remarks
' 同一親配下の兄弟ノードを巡回する
' ノード選択状態および表示情報を更新する
'
Public Sub MoveToSiblingNode()

    If gParentMap Is Nothing Then

        MsgBox "PathChartを再生成してください。"

        Exit Sub

    End If
    
    ' PathChartシート
    Dim ws As Worksheet
    
    ' 移動先ノード
    Dim shp As Shape

    ' 親ノード名
    Dim ParentName As String
    
    ' 兄弟ノード一覧
    Dim Ary As Variant

    ' 配列走査Index
    Dim i As Long
    
    ' 現在位置
    Dim CurrentPos As Long

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

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

    '# 親ノード取得
    ParentName = gParentMap(gCurrentShapeName)

    If Not gChildMap.Exists(ParentName) Then

        MsgBox "兄弟ノードはありません。"
        Exit Sub

    End If

    '# 兄弟ノード一覧取得
    Ary = Split(gChildMap(ParentName), "|")

    CurrentPos = -1

    '# 現在位置検索
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
    
    '-------------------------
    '# 次兄弟ノード選択
    '-------------------------
    CurrentPos = CurrentPos + 1

    If CurrentPos > UBound(Ary) Then

        CurrentPos = 0

    End If

    '-------------------------
    '# 兄弟ノード取得
    '-------------------------
    Set shp = ws.Shapes(Ary(CurrentPos))

    '--------------------------
    '# 現在ノード更新
    '--------------------------
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
    
    '# ノード強調表示
    Call HighlightNode(shp)

    gCurrentParentWithChildren = _
        gCurrentParentShapeName


    '# Navigator状態更新
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
' @Input
' shp(Shape)
'
' @Output
' なし
'
' @Summary
' 現在ノードおよび関連ノードを強調表示する
'
' @Remarks
' 現在ノード、親ノード、兄弟ノード、
' 子ノードおよび接続線を強調表示する
'
Private Sub HighlightNode( _
                ByVal shp As Shape)

    '# 強調表示初期化
    Call ResetChartHighlight

    '# 現在ノード強調
    Call HighlightCurrentNode(shp)

    '# 親ノード強調
    Call HighlightParentNode( _
            gCurrentParentShapeName)

    '# 親接続線強調
    Call HighlightParentConnector
    
    '# 兄弟ノード強調
    Call HighlightSiblingNodes( _
        gCurrentParentShapeName)
    
    '# 兄弟接続線強調
    Call HighlightSiblingConnectors( _
            gCurrentParentWithChildren)
    
    '# 子ノード強調
    Call HighlightChildNodes( _
        gCurrentShapeName)

End Sub

' @JPName
' 現在ノード強調
'
' @Category
' PathChart
'
' @Input
' shp(Shape)
'
' @Output
' なし
'
' @Summary
' 現在選択中ノードを強調表示する
'
' @Remarks
' 現在選択中ノードを赤枠で強調表示する
'
Private Sub HighlightCurrentNode( _
                ByVal shp As Shape)

    '# 現在ノード強調表示
    With shp

        ' 現在ノードだけ赤
        .line.ForeColor.RGB = CURRENT_NODE_LINE_COLOR
        .line.Weight = CURRENT_NODE_LINE_WEIGHT

    End With

End Sub

' @JPName
' 親ノード強調
'
' @Category
' PathChart
'
' @Input
' ParentShapeName(String)
'
' @Output
' なし
'
' @Summary
' 親ノードを強調表示する
'
' @Remarks
' 現在ノードの親ノードを強調表示する
' 背景色は変更しない
'
Private Sub HighlightParentNode( _
                ByVal ParentShapeName As String)

    ' PathChartシート
    Dim ws As Worksheet

    If gParentMap Is Nothing Then Exit Sub

    If ParentShapeName = "" Then Exit Sub

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 親ノード強調表示
    With ws.Shapes(ParentShapeName)

        .line.ForeColor.RGB = PARENT_NODE_LINE_COLOR
        .line.Weight = CURRENT_NODE_LINE_WEIGHT

        ' 背景色は変更しない

    End With

End Sub

' @JPName
' 親接続線強調
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
' 親ノードとの接続線を強調表示する
'
' @Remarks
' 親ノードと現在ノード間の接続線を
' 強調表示する
'
Private Sub HighlightParentConnector()

    If gParentMap Is Nothing Then Exit Sub
    
    If gCurrentParentShapeName = "" Then
        Exit Sub
    End If

    ' 親→現在ノード へのコネクタ赤
    '# 親接続線強調表示
    HighlightConnector _
        gCurrentParentShapeName, _
        gCurrentShapeName, _
        CURRENT_NODE_LINE_COLOR, _
        PARENT_CONNECTOR_WEIGHT

End Sub

' @JPName
' 子ノード強調
'
' @Category
' PathChart
'
' @Input
' ParentShapeName(String)
'
' @Output
' なし
'
' @Summary
' 現在ノードの子ノードを強調表示する
'
' @Remarks
' 現在ノード配下の子ノードを強調表示し
' 関連接続線も強調表示する
'
Private Sub HighlightChildNodes( _
                ByVal ParentShapeName As String)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' 子ノード一覧
    Dim Ary As Variant

    ' 子ノード名
    Dim ChildName As Variant

    If gContext Is Nothing Then Exit Sub
    If gChildMap Is Nothing Then Exit Sub
    
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then Exit Sub

    '# 子ノード一覧取得
    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    '# 子ノード強調表示
    For Each ChildName In Ary
    
        'Debug.Print "Child=" & ChildName
    
        If CStr(ChildName) <> gCurrentShapeName Then
    
        With ws.Shapes(CStr(ChildName))

            ' 子だけ薄黄色
            .line.ForeColor.RGB = CHILD_LINE_COLOR

            .line.Weight = CHILD_LINE_WEIGHT
            
            .Fill.ForeColor.RGB = CHILD_FILL_COLOR
            
            .TextFrame.Characters.Font.Color = CHILD_FONT_COLOR

        End With

        '# 子接続線強調表示
        HighlightConnector _
            ParentShapeName, _
            CStr(ChildName), _
            CHILD_CONNECTOR_COLOR, _
            CHILD_CONNECTOR_WEIGHT
        
        End If

    Next ChildName

End Sub


' @JPName
' 兄弟ノード強調
'
' @Category
' PathChart
'
' @Input
' ParentShapeName(String)
'
' @Output
' なし
'
' @Summary
' 兄弟ノードを強調表示する
'
' @Remarks
' 現在ノード以外の兄弟ノードを
' 強調表示する
'
Private Sub HighlightSiblingNodes( _
                ByVal ParentShapeName As String)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' 兄弟ノード一覧
    Dim Ary As Variant
    
    ' 兄弟ノード名
    Dim ChildName As Variant

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then
        Exit Sub
    End If

    '# 兄弟ノード一覧取得
    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    '# 兄弟ノード強調表示
    For Each ChildName In Ary

        If CStr(ChildName) <> _
           gCurrentShapeName Then

            With ws.Shapes(CStr(ChildName))

                .line.ForeColor.RGB = SIBLING_LINE_COLOR

                .line.Weight = CHILD_LINE_WEIGHT

                .Fill.ForeColor.RGB = SIBLING_FILL_COLOR

            End With

            '# 兄弟接続線強調表示
            HighlightConnector _
                ParentShapeName, _
                CStr(ChildName), _
                SIBLING_LINE_COLOR, _
                SIBLING_CONNECTOR_WEIGHT

        End If

    Next ChildName

End Sub


' @JPName
' 兄弟接続線強調
'
' @Category
' PathChart
'
' @Input
' ParentShapeName(String)
'
' @Output
' なし
'
' @Summary
' 兄弟ノードへの接続線を強調表示する
'
' @Remarks
' 兄弟ノードへの接続線を
' 強調表示する
'
Private Sub HighlightSiblingConnectors( _
                ByVal ParentShapeName As String)

    ' 兄弟ノード一覧
    Dim Ary As Variant
    
    ' 兄弟ノード名
    Dim ChildName As Variant

    If gChildMap Is Nothing Then Exit Sub

    If Not gChildMap.Exists(ParentShapeName) Then
        Exit Sub
    End If

    Ary = Split( _
            gChildMap(ParentShapeName), _
            "|")

    '# 兄弟接続線強調表示
    For Each ChildName In Ary

        If CStr(ChildName) <> gCurrentShapeName Then

            ' 兄弟へのコネクタだけ薄緑
            HighlightConnector _
                ParentShapeName, _
                CStr(ChildName), _
                SIBLING_LINE_COLOR, _
                SIBLING_CONNECTOR_WEIGHT

        End If

    Next ChildName

End Sub

' @JPName
' 接続線強調
'
' @Category
' PathChart
'
' @Input
' ParentShapeName(String)
' ChildShapeName(String)
' LineColor(Long)
' LineWeight(Double)
'
' @Output
' なし
'
' @Summary
' 指定ノード間の接続線を強調表示する
'
' @Remarks
' 指定された親子ノード間の接続線へ
' 強調表示書式を設定する
'
Private Sub HighlightConnector( _
                ByVal ParentShapeName As String, _
                ByVal ChildShapeName As String, _
                ByVal LineColor As Long, _
                ByVal LineWeight As Double)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' 接続線名
    Dim ConnectorName As String

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 接続線名称生成
    ConnectorName = _
        "C_" & _
        ParentShapeName & "_" & _
        ChildShapeName

    On Error Resume Next

    '# 接続線強調表示
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChartの強調表示を初期状態へ戻す
'
' @Remarks
' ノードおよび接続線の表示属性を
' 初期状態へ戻す
'
Private Sub ResetChartHighlight()

    ' PathChartシート
    Dim ws As Worksheet
    
    ' Shape
    Dim shp As Shape

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# 強調表示初期化
    For Each shp In ws.Shapes

        '# ノード書式初期化
        If Left$(shp.Name, 2) = "N_" Then

            shp.line.ForeColor.RGB = NODE_LINE_COLOR
            shp.line.Weight = NODE_LINE_WEIGHT

            shp.Fill.ForeColor.RGB = NODE_FILL_COLOR

            shp.TextFrame.Characters.Font.Color = NODE_FONT_COLOR

        End If

        '# 接続線書式初期化
        If Left$(shp.Name, 2) = "C_" Then

            shp.line.ForeColor.RGB = CONNECTOR_LINE_COLOR
            shp.line.Weight = CONNECTOR_LINE_WEIGHT

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

    '# 強調表示解除
    Call ResetChartHighlight

    '# ナビゲーション状態初期化
    gCurrentShapeName = ""

    gCurrentParentShapeName = ""

    gCurrentParentWithChildren = ""

    gCurrentNodeText = ""

    gNavigatorStatus = ""

    Application.StatusBar = False

    UpdateNavigatorStatus ""

    '# Navigator表示クリア
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
' @Input
' Msg(String)
'
' @Output
' なし
'
' @Summary
' Navigatorおよびステータスバーへ状態を表示する
'
' @Remarks
' Navigator画面とステータスバーへ
' 状態メッセージを表示する
'
Public Sub UpdateNavigatorStatus( _
                ByVal Msg As String)

    '# Navigator状態保存
    gNavigatorStatus = Msg

    On Error Resume Next

    '# Navigator表示更新
    If frmPathChartNavigator.Visible Then

        frmPathChartNavigator.txtStatus.Text = _
            Msg

    End If

    On Error GoTo 0

    '# ステータスバー更新
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
' @Input
' cbo(MSForms.ComboBox)
'
' @Output
' なし
'
' @Summary
' 開始関数候補をコンボボックスへ設定する
'
' @Remarks
' 関数依存関係シートから
' 開始関数候補一覧を生成する
'
Public Sub LoadStartProcedure( _
                ByVal cbo As MSForms.ComboBox)

    ' 関数依存関係シート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 重複除外管理
    Dim Dic As Object

    Set Dic = CreateObject("Scripting.Dictionary")

    Set ws = _
        gContext.ManagementBook.Worksheets(SHEET_FUNCTION_DEPENDENCY)

    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    cbo.Clear

    '# 開始Procedure一覧取得
    For r = 2 To LastRow

        If Trim$(ws.Cells(r, 1).Value) <> "" Then

            '# 開始Procedure候補追加
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 検索履歴コレクションを初期化する
'
' @Remarks
' 検索履歴コレクション未生成時のみ
' 初期化を実施する
'
Public Sub InitSearchHistory()

    '# 検索履歴初期化
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
' @Input
' SearchText(String)
'
' @Output
' なし
'
' @Summary
' 検索履歴へキーワードを追加する
'
' @Remarks
' 重複登録を防止しながら
' 検索履歴へ追加する
'
Public Sub AddSearchHistory( _
                ByVal SearchText As String)

    ' 検索履歴走査用
    Dim V As Variant

    '# 検索文字列確認
    If Trim$(SearchText) = "" Then Exit Sub

    InitSearchHistory

    '# 検索履歴重複確認
    For Each V In gSearchHistory

        If StrComp(V, SearchText, _
                   vbTextCompare) = 0 Then
            Exit Sub
        End If

    Next

    '# 検索履歴追加
    gSearchHistory.Add SearchText

End Sub

' @JPName
' 検索履歴読込
'
' @Category
' PathChart
'
' @Input
' cbo(MSForms.ComboBox)
'
' @Output
' なし
'
' @Summary
' 検索履歴をコンボボックスへ設定する
'
' @Remarks
' 保持している検索履歴を
' コンボボックスへ設定する
'
Public Sub LoadSearchHistory( _
                ByVal cbo As MSForms.ComboBox)

    ' 検索履歴
    Dim V As Variant

    InitSearchHistory

    '# 検索履歴読込
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
' @Input
' cbo(MSForms.ComboBox)
'
' @Output
' なし
'
' @Summary
' PathChart内のノード一覧を検索候補として読み込む
'
' @Remarks
' PathChart内のノードから
' 重複除外した候補一覧を生成する
'
Public Sub LoadSearchNodeList( _
                ByVal cbo As MSForms.ComboBox)

    ' PathChartシート
    Dim ws As Worksheet
    
    ' ノードShape
    Dim shp As Shape
    
    ' 重複除外管理
    Dim Dic As Object
    
    ' ノード文字列
    Dim NodeText As String
    
    ' Dictionary走査用
    Dim V As Variant

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    Set Dic = _
        CreateObject("Scripting.Dictionary")

    cbo.Clear

    '# 検索候補収集
    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then
        
        #If DEBUG_MODE Then
            Debug.Print shp.Name
            Debug.Print shp.TextFrame.Characters.Text
        #End If

            '# Procedure名取得
            NodeText = _
                GetProcedureNameFromNodeText( _
                    shp.TextFrame.Characters.Text)

            NodeText = Trim$(NodeText)

            If NodeText <> "" Then

                '# 検索候補登録
                If Not Dic.Exists(NodeText) Then

                    Dic.Add _
                        NodeText, _
                        True

                End If

            End If

        End If

    Next shp

    '# 検索候補表示
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChart操作用ショートカットキーを登録する
'
' @Remarks
' PathChart操作用ショートカットキーを
' Excelへ登録する
'
Public Sub RegisterShortcut()

    '# PathChartショートカット登録
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
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' PathChart操作用ショートカットキーを解除する
'
' @Remarks
' 登録済みショートカットキーを解除する
'
Public Sub UnRegisterShortcut()

    '# PathChartショートカット解除
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
' @Input
' StartProc(String)
' MaxDepth(Long)
'
' @Output
' なし
'
' @Summary
' PathChartをPowerPointへ出力する
'
' @Remarks
' PowerPoint出力モジュールを呼び出して
' PathChartをPowerPoint形式で出力する
'
Private Sub ExportPathChartToPowerPoint( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)

    '# PowerPoint出力実行
    Call modPowerPointChart.CreatePowerPointPathChart( _
            StartProc, _
            MaxDepth)

End Sub

