Attribute VB_Name = "modPowerPointChart"
Option Explicit
'===================================================================
' @JPName
' PathChart PowerPoint出力
'
' @Category
' PathChart
'
' @Summary
' PathChartをPowerPointおよびPNG形式で
' 出力する機能を提供する
'
' @Remarks
' PathChartシートを画像として取得し
'
' ・PowerPoint出力
' ・PNG出力
' ・出力ファイル管理
'
' を行う
'===================================================================

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
' 縦スライド幅
Private Const PPT_WIDTH_PORTRAIT As Long = 540
' 縦スライド高さ
Private Const PPT_HEIGHT_PORTRAIT As Long = 960

' 横スライド
' 横スライド幅
Private Const PPT_WIDTH_LANDSCAPE As Long = 960
' 横スライド高さ
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
' PathChartシートを画像として取得し
' PowerPointへ貼り付けて保存する
' スライド向きは設定値または
' チャートサイズに応じて決定する
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

    ' チャート範囲
    Dim rng As Range

    ' チャートサイズ
    Dim ChartWidth As Double
    Dim ChartHeight As Double
    
    ' グリッド線表示状態退避
    Dim GridLineFlg As Boolean

    ' 出力ファイルパス
    Dim PptFile As String

    '# PathChart出力範囲取得
    Set rng = GetChartRange()

    '# チャートサイズ取得
    ChartWidth = rng.Width
    ChartHeight = rng.Height
    
    '=================================
    '# グリッド線一時非表示
    '=================================
    GridLineFlg = ActiveWindow.DisplayGridlines

    ActiveWindow.DisplayGridlines = False

    '# PathChart画像取得
    rng.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    '# グリッド線表示復元
    ActiveWindow.DisplayGridlines = _
        GridLineFlg

    '=================================
    '# PowerPoint起動
    '=================================
    Set pptApp = _
        CreateObject("PowerPoint.Application")

    pptApp.Visible = True

    '# プレゼンテーション生成
    Set pptPres = _
        pptApp.Presentations.Add

    '=================================
    '# スライド向き設定
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

            '# チャートサイズで向き自動判定
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
    '# PowerPointスライド追加
    '=================================
    Set pptSlide = _
        pptPres.Slides.Add( _
            PPT_SLIDE_INDEX, _
            PPT_LAYOUT_BLANK)
    '=================================
    '# スライドタイトル生成
    '=================================
    Set pptTitle = _
        pptSlide.Shapes.AddTextbox( _
            PPT_TEXT_HORIZONTAL, _
            PPT_TITLE_LEFT, _
            PPT_TITLE_TOP, _
            pptPres.PageSetup.SlideWidth - _
            PPT_SIDE_MARGIN, _
            PPT_TITLE_HEIGHT)

    '=================================
    '# タイトル設定
    '=================================
    
    '# スライドタイトル設定
    pptTitle.TextFrame.TextRange.Text = _
        StartProc & _
        " (Level=" & _
        MaxDepth & ")"
    
    '# タイトル書式設定
    With pptTitle.TextFrame.TextRange

        .Font.Size = PPT_TITLE_FONT_SIZE
        .Font.Bold = True
        .ParagraphFormat.Alignment = 2

    End With

    '=================================
    '# PathChart貼付
    '=================================
    pptSlide.Shapes.Paste

    '=================================
    '# 貼付図形取得
    '=================================
    Set pptShape = _
        pptSlide.Shapes( _
            pptSlide.Shapes.Count)

    '# 縦横比固定
    pptShape.LockAspectRatio = True
    
    '=================================
    ' 幅調整
    '=================================
    '# スライド幅内へ縮小
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
    '# スライド高さ内へ縮小
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
    '# スライド中央配置
    pptShape.Left = _
        (pptPres.PageSetup.SlideWidth - _
         pptShape.Width) / 2

    '# チャート配置位置設定
    pptShape.Top = PPT_CHART_TOP

    '=================================
    ' 保存
    '=================================
    '# PowerPoint出力ファイル生成
    PptFile = _
        GetOutputFolder() & _
        StartProc & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".pptx"

    '# PowerPoint保存
    pptPres.SaveAs PptFile

    '# PowerPoint出力完了通知
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
' PathChartシートをPowerPointへ貼り付け
' PNG画像として保存する
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

    ' グリッド線表示状態退避
    Dim GridLineFlg As Boolean

    ' 出力PNGファイルパス
    Dim PngFile As String

    '# PathChartシート取得
    Set ws = gContext.ManagementBook.Worksheets(SHEET_PATH_CHART)

    '# PathChart出力範囲取得
    Set rng = GetChartRange()

    '=================================
    '# グリッド線一時非表示
    '=================================
    GridLineFlg = _
        ActiveWindow.DisplayGridlines

    ActiveWindow.DisplayGridlines = False

    '# PathChart画像取得
    rng.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    '# グリッド線表示復元
    ActiveWindow.DisplayGridlines = _
        GridLineFlg

    '=================================
    '# PowerPoint起動
    '=================================
    Set pptApp = _
        CreateObject("PowerPoint.Application")

    pptApp.Visible = True

    '# プレゼンテーション生成
    Set pptPres = _
        pptApp.Presentations.Add

    '=================================
    '# PowerPointスライド追加
    '=================================
    Set pptSlide = _
        pptPres.Slides.Add( _
            PPT_SLIDE_INDEX, _
            PPT_LAYOUT_BLANK)

    '=================================
    '# PathChart貼付
    '=================================
    pptSlide.Shapes.Paste

    '=================================
    '# 貼付図形取得
    '=================================
    Set pptShape = _
        pptSlide.Shapes( _
            pptSlide.Shapes.Count)

    '=================================
    ' 保存
    '=================================
    '# PNG出力ファイル生成
    PngFile = _
        GetOutputFolder() & _
        StartProc & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".png"
    
    On Error Resume Next

    '=================================
    '# PNG保存
    '=================================
    pptShape.Export PngFile, 2
    
    '=================================
    '# 後処理
    '=================================
    
    '# PowerPoint終了
    pptPres.Close
    
    '# PowerPointアプリ終了
    pptApp.Quit
    
    '# PowerPointオブジェクト解放
    Set pptShape = Nothing
    Set pptSlide = Nothing
    Set pptPres = Nothing
    Set pptApp = Nothing

    '=================================
    '# 結果通知
    '=================================
    
    '# PNG保存結果通知
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

    ' 出力フォルダ
    Dim Folder As String

    '--------------------------
    '# 出力フォルダ生成
    '--------------------------
    Folder = _
        Environ$("USERPROFILE") & _
        "\OneDrive - east.ntt.co.jp\" & _
        "ドキュメント\GitHub\SOSDB_Analyzer\"

    '--------------------------
    '# フォルダ存在確認
    '--------------------------
    
    '# 出力フォルダ存在確認
    If Dir(Folder, vbDirectory) = "" Then

        '# 出力フォルダ作成
        MkDir Folder

    End If

    '--------------------------
    '# 出力フォルダ返却
    '--------------------------
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
' @Remarks
' 出力ファイル名生成時に利用する
'
' yyyymmdd_hhnnss
'
' 形式で返却する
'
Public Function GetTimeStamp() As String

    '# 出力用タイムスタンプ生成
    GetTimeStamp = _
        Format(Now, "yyyymmdd_hhnnss")

End Function

