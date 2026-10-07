Attribute VB_Name = "modPowerPointChart"
Option Explicit

'==================================================
' PowerPointレイアウト定数
'==================================================

' スライド関連

' スライド番号
Private Const PPT_SLIDE_INDEX As Long = 1

' 空白スライド
Private Const PPT_LAYOUT_BLANK As Long = 12

' PowerPointテキスト方向(横書き)
Private Const PPT_TEXT_HORIZONTAL As Long = 1

' 縦スライド
Private Const PPT_WIDTH_PORTRAIT As Long = 540
Private Const PPT_HEIGHT_PORTRAIT As Long = 960

' 横スライド
Private Const PPT_WIDTH_LANDSCAPE As Long = 960
Private Const PPT_HEIGHT_LANDSCAPE As Long = 540

' タイトル位置
Private Const PPT_TITLE_LEFT As Long = 10
Private Const PPT_TITLE_TOP As Long = 5
Private Const PPT_TITLE_HEIGHT As Long = 30

' タイトルフォント
Private Const PPT_TITLE_FONT_SIZE As Long = 18

' チャート上余白
Private Const PPT_CHART_TOP As Long = 40

' スライド左右余白
Private Const PPT_SIDE_MARGIN As Long = 20

' スライド下余白
Private Const PPT_BOTTOM_MARGIN As Long = 80

'==================================================
' PowerPoint出力
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
' PowerPointファイル
'
' @Summary
' PathChartをPowerPointファイルとして出力する
'
' @Remarks
' スライドサイズは設定値または
' チャートサイズから自動決定する
'
Public Sub CreatePowerPointPathChart( _
                ByVal StartProc As String, _
                ByVal MaxDepth As Long)
    
    ' PowerPointアプリ
    Dim pptApp As Object
    
    ' プレゼンテーション
    Dim pptPres As Object
    
    ' スライド
    Dim pptSlide As Object

    ' 貼付チャート図形
    Dim pptShape As Object
    
    ' タイトル図形
    Dim pptTitle As Object

    ' PathChartシート
    Dim ws As Worksheet
    
    ' チャート範囲
    Dim rng As Range

    ' チャートサイズ
    Dim ChartWidth As Double
    Dim ChartHeight As Double
    
    ' グリッド線表示状態
    Dim GridLineFlg As Boolean

    ' 出力ファイル
    Dim PptFile As String

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    Set rng = GetChartRange()

    ChartWidth = rng.Width
    ChartHeight = rng.Height
    
    '=================================
    ' グリッド線を一時非表示
    '=================================
    GridLineFlg = ActiveWindow.DisplayGridlines

    ActiveWindow.DisplayGridlines = False

    rng.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    ActiveWindow.DisplayGridlines = _
        GridLineFlg

    '=================================
    ' PowerPoint起動
    '=================================
    Set pptApp = _
        CreateObject("PowerPoint.Application")

    pptApp.Visible = True

    Set pptPres = _
        pptApp.Presentations.Add

    '=================================
    ' 用紙向き
    '=================================
    Select Case gPptOrientation

        Case PPT_ORIENT_PORTRAIT

            ' 縦置き
            pptPres.PageSetup.SlideWidth = PPT_WIDTH_PORTRAIT
            pptPres.PageSetup.SlideHeight = PPT_HEIGHT_PORTRAIT

        Case PPT_ORIENT_LANDSCAPE

            ' 横置き
            pptPres.PageSetup.SlideWidth = PPT_WIDTH_LANDSCAPE
            pptPres.PageSetup.SlideHeight = PPT_HEIGHT_LANDSCAPE

        Case PPT_ORIENT_AUTO

            If ChartHeight > ChartWidth Then

                ' 縦置き
                pptPres.PageSetup.SlideWidth = PPT_WIDTH_PORTRAIT
                pptPres.PageSetup.SlideHeight = PPT_HEIGHT_PORTRAIT

            Else

                ' 横置き
                pptPres.PageSetup.SlideWidth = PPT_WIDTH_LANDSCAPE
                pptPres.PageSetup.SlideHeight = PPT_HEIGHT_LANDSCAPE

            End If

    End Select

    '=================================
    ' スライド追加
    '=================================
    Set pptSlide = _
        pptPres.Slides.Add( _
            PPT_SLIDE_INDEX, _
            PPT_LAYOUT_BLANK)
    '=================================
    ' タイトル
    '=================================
    Set pptTitle = _
        pptSlide.Shapes.AddTextbox( _
            PPT_TEXT_HORIZONTAL, _
            PPT_TITLE_LEFT, _
            PPT_TITLE_TOP, _
            pptPres.PageSetup.SlideWidth - _
            PPT_SIDE_MARGIN, _
            PPT_TITLE_HEIGHT)

    pptTitle.TextFrame.TextRange.Text = _
        StartProc & _
        " (Level=" & _
        MaxDepth & ")"
    
    With pptTitle.TextFrame.TextRange

        .Font.Size = PPT_TITLE_FONT_SIZE
        .Font.Bold = True
        .ParagraphFormat.Alignment = 2

    End With

    '=================================
    ' パスチャート貼付
    '=================================
    pptSlide.Shapes.Paste

    Set pptShape = _
        pptSlide.Shapes( _
            pptSlide.Shapes.Count)

    pptShape.LockAspectRatio = True
    
    '=================================
    ' 幅調整
    '=================================
    If pptShape.Width > _
        (pptPres.PageSetup.SlideWidth - _
        PPT_SIDE_MARGIN) Then

        pptShape.Width = _
            (pptPres.PageSetup.SlideWidth - _
            PPT_SIDE_MARGIN)

    End If
    
    '=================================
    ' 高さ調整
    '=================================
    If pptShape.Height > _
        (pptPres.PageSetup.SlideHeight - _
        PPT_BOTTOM_MARGIN) Then
        
        pptShape.Height = _
            (pptPres.PageSetup.SlideHeight - _
            PPT_BOTTOM_MARGIN)
        
    End If

    '=================================
    ' 中央配置
    '=================================
    pptShape.Left = _
        (pptPres.PageSetup.SlideWidth - _
         pptShape.Width) / 2

    pptShape.Top = PPT_CHART_TOP

    '=================================
    ' 保存
    '=================================
    PptFile = _
        GetOutputFolder() & _
        StartProc & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".pptx"

    pptPres.SaveAs PptFile

    MsgBox _
        "PowerPoint出力完了" & vbCrLf & _
        PptFile, _
        vbInformation

End Sub

' @JPName
' PNG出力
'
' @Category
' PathChart
'
' @Input
' StartProc(String)
' MaxDepth(Long)
'
' @Output
' PNGファイル
'
' @Summary
' PathChartをPNG形式で出力する
'
' @Remarks
' PowerPoint経由で画像ファイルを生成する
'
Public Sub ExportPathChartToPNG( _
        ByVal StartProc As String, _
        ByVal MaxDepth As Long)

    ' PowerPointアプリ
    Dim pptApp As Object
    
    ' プレゼンテーション
    Dim pptPres As Object
    
    ' スライド
    Dim pptSlide As Object

    ' 貼付チャート図形
    Dim pptShape As Object

    ' PathChartシート
    Dim ws As Worksheet
    
    ' チャート範囲
    Dim rng As Range

    ' グリッド線表示状態
    Dim GridLineFlg As Boolean

    ' 出力ファイル
    Dim PngFile As String

    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    Set rng = GetChartRange()

    '=================================
    ' グリッド線を一時非表示
    '=================================
    GridLineFlg = _
        ActiveWindow.DisplayGridlines

    ActiveWindow.DisplayGridlines = False

    rng.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    ActiveWindow.DisplayGridlines = _
        GridLineFlg

    '=================================
    ' PowerPoint起動
    '=================================
    Set pptApp = _
        CreateObject("PowerPoint.Application")

    pptApp.Visible = True

    Set pptPres = _
        pptApp.Presentations.Add

    '=================================
    ' スライド追加
    '=================================
    Set pptSlide = _
        pptPres.Slides.Add( _
            PPT_SLIDE_INDEX, _
            PPT_LAYOUT_BLANK)

    '=================================
    ' パスチャート貼付
    '=================================
    pptSlide.Shapes.Paste

    Set pptShape = _
        pptSlide.Shapes( _
            pptSlide.Shapes.Count)

    '=================================
    ' 保存
    '=================================
    PngFile = _
        GetOutputFolder() & _
        StartProc & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".png"
    
    On Error Resume Next

    pptShape.Export PngFile, 2
    
    pptPres.Close
    
    pptApp.Quit
    
    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    If Err.Number <> 0 Then

        MsgBox _
            "PNG保存失敗" & vbCrLf & _
            Err.Description, _
            vbExclamation

    Else

        MsgBox _
            "PNG保存完了" & vbCrLf & _
            PngFile, _
            vbInformation

    End If

    On Error GoTo 0

End Sub

'==================================================
' 共通
'==================================================

' @JPName
' 出力フォルダ取得
'
' @Category
' PathChart
'
' @Input
' なし
'
' @Output
' FolderPath(String)
'
' @Summary
' PathChart出力先フォルダを取得する
'
' @Remarks
' フォルダが存在しない場合は作成する
'
Public Function GetOutputFolder() As String

    Dim Folder As String

    Folder = _
        Environ$("USERPROFILE") & _
        "\OneDrive - east.ntt.co.jp\" & _
        "ドキュメント\GitHub\SOSDB_Analyzer\"

    If Dir(Folder, vbDirectory) = "" Then

        MkDir Folder

    End If

    GetOutputFolder = Folder

End Function

' @JPName
' タイムスタンプ取得
'
' @Category
' Common
'
' @Input
' なし
'
' @Output
' TimeStamp(String)
'
' @Summary
' ファイル名用タイムスタンプを取得する
'
Public Function GetTimeStamp() As String

    GetTimeStamp = _
        Format(Now, "yyyymmdd_hhnnss")

End Function

