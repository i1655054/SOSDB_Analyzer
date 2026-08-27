Attribute VB_Name = "modPowerPointChart"
Option Explicit

Public Sub CreatePowerPointPathChart()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    Dim pptShape As Object
    Dim pptTitle As Object

    Dim ws As Worksheet
    Dim rng As Range

    Dim GridLineFlg As Boolean

    Dim PptFile As String

    Set ws = Worksheets("PathChart")

    Set rng = GetChartRange()

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

    Debug.Print pptApp.Version

    Set pptPres = _
        pptApp.Presentations.Add

    ' ppLayoutBlank
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
        ws.Range(CONFIG_START_FUNC).Value & _
        " (Level=" & _
        ws.Range(CONFIG_MAX_LEVEL).Value & ")"

    With pptTitle.TextFrame.TextRange

        .Font.Size = 18
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

    Debug.Print _
        "pptShape.Type:" & _
        pptShape.Type

    Debug.Print _
        "pptShape.Name:" & _
        pptShape.Name

    Debug.Print _
        "pptShape.Width:" & _
        pptShape.Width

    Debug.Print _
        "pptShape.Height:" & _
        pptShape.Height

    pptShape.LockAspectRatio = True

    If pptShape.Width > 1000 Then

        pptShape.Width = 1000

    End If

    '=================================
    ' 中央配置
    '=================================
    pptShape.Left = _
        (pptPres.PageSetup.SlideWidth - _
         pptShape.Width) / 2

    pptShape.Top = 40

    '=================================
    ' PPT保存
    '=================================
    PptFile = _
        GetOutputFolder() & _
        ws.Range(CONFIG_START_FUNC).Value & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".pptx"

    pptPres.SaveAs PptFile

    MsgBox _
        "PowerPoint出力完了" & vbCrLf & _
        PptFile, _
        vbInformation

End Sub

Public Sub ExportPathChartToPNG()

    Dim pptApp As Object
    Dim pptPres As Object
    Dim pptSlide As Object

    Dim pptShape As Object

    Dim ws As Worksheet
    Dim rng As Range

    Dim GridLineFlg As Boolean

    Dim PngFile As String

    Set ws = Worksheets("PathChart")

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
        ws.Range(CONFIG_START_FUNC).Value & _
        "_PathChart_" & _
        gExportTimeStamp & _
        ".png"
    
    Debug.Print PngFile

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

Private Sub ExportPathChartToPPT( _
            ByVal pptSlide As Object)

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = Worksheets("PathChart")

    For Each shp In ws.Shapes

        If Left$(shp.Name, 2) = "N_" Then

            pptSlide.Shapes.AddShape _
                msoShapeRoundedRectangle, _
                shp.Left, _
                shp.Top + 50, _
                shp.Width, _
                shp.Height

            With pptSlide.Shapes( _
                    pptSlide.Shapes.Count)

                .TextFrame.TextRange.Text = _
                    shp.TextFrame.Characters.Text

            End With

        End If

    Next shp

End Sub

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

Public Function GetTimeStamp() As String

    GetTimeStamp = _
        Format(Now, "yyyymmdd_hhnnss")

End Function

