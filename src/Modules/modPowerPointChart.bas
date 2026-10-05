Attribute VB_Name = "modPowerPointChart"
Option Explicit

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
    
    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    Dim pptShape As Object
    Dim pptTitle As Object

    Dim ws As Worksheet
    Dim rng As Range

    Dim ChartWidth As Double
    Dim ChartHeight As Double
    
    Dim GridLineFlg As Boolean

    Dim PptFile As String

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If
    
    Set ws = ManagementBook.Worksheets("PathChart")

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
            pptPres.PageSetup.SlideWidth = 540
            pptPres.PageSetup.SlideHeight = 960

        Case PPT_ORIENT_LANDSCAPE

            ' 横置き
            pptPres.PageSetup.SlideWidth = 960
            pptPres.PageSetup.SlideHeight = 540

        Case PPT_ORIENT_AUTO

            If ChartHeight > ChartWidth Then

                ' 縦置き
                pptPres.PageSetup.SlideWidth = 540
                pptPres.PageSetup.SlideHeight = 960

            Else

                ' 横置き
                pptPres.PageSetup.SlideWidth = 960
                pptPres.PageSetup.SlideHeight = 540

            End If

    End Select

    '=================================
    ' スライド追加
    '=================================
    Set pptSlide = _
        pptPres.Slides.Add(1, 12)

    '=================================
    ' タイトル
    '=================================
    Set pptTitle = _
        pptSlide.Shapes.AddTextbox( _
            1, _
            10, _
            5, _
            pptPres.PageSetup.SlideWidth - 20, _
            30)

    pptTitle.TextFrame.TextRange.Text = _
        StartProc & _
        " (Level=" & _
        MaxDepth & ")"
    
    With pptTitle.TextFrame.TextRange

        .Font.size = 18
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
        (pptPres.PageSetup.SlideWidth - 20) Then

        pptShape.Width = _
            (pptPres.PageSetup.SlideWidth - 20)

    End If
    
    '=================================
    ' 高さ調整
    '=================================
    If pptShape.Height > _
        (pptPres.PageSetup.SlideHeight - 80) Then
        
        pptShape.Height = _
            (pptPres.PageSetup.SlideHeight - 80)
        
    End If

    '=================================
    ' 中央配置
    '=================================
    pptShape.Left = _
        (pptPres.PageSetup.SlideWidth - _
         pptShape.Width) / 2

    pptShape.Top = 40

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

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    Dim pptShape As Object

    Dim ws As Worksheet
    Dim rng As Range

    Dim GridLineFlg As Boolean

    Dim PngFile As String

    Dim ManagementBook As Workbook

    Set ManagementBook = _
        GetCurrentManagementWorkbook()

    If Not CheckManagementWorkbook( _
            ManagementBook, _
            GetCurrentManagementFile()) Then

        Exit Sub

    End If
    
    Set ws = ManagementBook.Worksheets("PathChart")

    Set rng = GetChartRange()

    GridLineFlg = _
        ActiveWindow.DisplayGridlines

    ActiveWindow.DisplayGridlines = False

    rng.CopyPicture _
        Appearance:=xlScreen, _
        Format:=xlPicture

    ActiveWindow.DisplayGridlines = _
        GridLineFlg

    Set pptApp = _
        CreateObject("PowerPoint.Application")

    pptApp.Visible = True

    Set pptPres = _
        pptApp.Presentations.Add

    Set pptSlide = _
        pptPres.Slides.Add(1, 12)

    pptSlide.Shapes.Paste

    Set pptShape = _
        pptSlide.Shapes( _
            pptSlide.Shapes.Count)

    Dim SaveFolder As String

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

