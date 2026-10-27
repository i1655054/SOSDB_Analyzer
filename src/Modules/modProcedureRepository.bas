Attribute VB_Name = "modProcedureRepository"
Option Explicit

' InitializeProcedureRepository
Public Sub InitializeProcedureRepository()

    If gContext Is Nothing Then Exit Sub
    
    If gContext.ProcInfoLoaded Then Exit Sub

    BuildProcInfoDictionary

    If Not gContext.ProcInfoDic Is Nothing Then
        
        gContext.ProcInfoLoaded = True
    
    End If

End Sub

' BuildProcInfoDictionary
' ProcListシートから辞書を生成します。
Public Sub BuildProcInfoDictionary()

    Dim ws As Worksheet

    Dim LastRow As Long
    Dim r As Long

    Dim Info As clsProcInfo
    Dim Key As String

    On Error Resume Next
    
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    On Error GoTo 0

    If ws Is Nothing Then

        MsgBox _
            "ProcListが存在しません。" & vbCrLf & _
            "先に解析実行を行ってください。", _
            vbExclamation

        Exit Sub

    End If

    Set gContext.ProcInfoDic = _
        CreateObject("Scripting.Dictionary")

    Set gContext.ProcIndexDic = _
        CreateObject("Scripting.Dictionary")

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    For r = 2 To LastRow

        Set Info = New clsProcInfo

        Info.ModuleName = _
            Trim(ws.Cells(r, COL_MODULE).Value)

        Info.ProcName = _
            Trim(ws.Cells(r, COL_PROC).Value)

        Info.ProcType = _
            Trim(ws.Cells(r, COL_TYPE).Value)

        Info.JPName = _
            Trim(ws.Cells(r, COL_JPNAME).Value)

        Info.Category = _
            Trim(ws.Cells(r, COL_CATEGORY).Value)

        Info.InputText = _
            Trim(ws.Cells(r, COL_INPUT).Value)

        Info.OutputText = _
            Trim(ws.Cells(r, COL_OUTPUT).Value)

        Info.Summary = _
            Trim(ws.Cells(r, COL_SUMMARY).Value)

        Info.Remarks = _
            Trim(ws.Cells(r, COL_REMARKS).Value)

        Info.UsageCount = _
            CLng(Val(ws.Cells(r, COL_USAGECOUNT).Value))

        Info.Status = _
            Trim(ws.Cells(r, COL_STATUS).Value)

        Key = _
            Info.ModuleName & "|" & _
            Info.ProcName

        If Not gContext.ProcInfoDic.Exists(Key) Then

            gContext.ProcInfoDic.Add Key, Info

        Else

            Set gContext.ProcInfoDic(Key) = Info

        End If

        BuildProcIndex Info

    Next r

End Sub

' BuildProcIndex
' プロシージャ名だけから候補モジュールを取得する索引です。
Private Sub BuildProcIndex( _
                ByVal Info As clsProcInfo)

    Dim Col As Collection

    If gContext.ProcIndexDic.Exists( _
            Info.ProcName) Then

        Set Col = _
            gContext.ProcIndexDic( _
                Info.ProcName)

    Else

        Set Col = New Collection

        gContext.ProcIndexDic.Add _
            Info.ProcName, _
            Col

    End If

    If Not ExistsModule(Col, Info.ModuleName) Then
    
        Col.Add Info.ModuleName
    
    End If

End Sub


' ResolveProcedure
' 呼出元モジュールを優先して候補を決定します。
Public Function ResolveProcedure( _
                ByVal CallerModule As String, _
                ByVal ProcName As String) _
                As clsProcInfo
  
    Dim Key As String
    
    Dim Col As Collection
    Dim V As Variant

    If gContext Is Nothing Then Exit Function

    If gContext.ProcInfoDic Is Nothing Then Exit Function

    If gContext.ProcIndexDic Is Nothing Then Exit Function

    If Trim(ProcName) = "" Then Exit Function
    
    Key = _
        CallerModule & "|" & ProcName

    If gContext.ProcInfoDic.Exists(Key) Then

        Set ResolveProcedure = _
            gContext.ProcInfoDic(Key)

        Exit Function

    End If

    If Not gContext.ProcIndexDic.Exists( _
            ProcName) Then Exit Function

    Set Col = _
        gContext.ProcIndexDic(ProcName)

    If Col.Count = 1 Then

        Key = _
            Col(1) & "|" & ProcName

        Set ResolveProcedure = _
            gContext.ProcInfoDic(Key)

        Exit Function

    End If

    If Col.Count > 1 Then

        Debug.Print _
            "ResolveProcedure Ambiguous : " _
            & ProcName _
            & " CallerModule=" _
            & CallerModule

        For Each V In Col

            Debug.Print _
                "  Candidate : " & CStr(V)

        Next V
    
    End If

End Function


' ClearProcedureRepository
' 対象ツール変更時や再解析時に実行します。
Public Sub ClearProcedureRepository()

    If gContext Is Nothing Then Exit Sub
    
    Set gContext.ProcInfoDic = Nothing
    Set gContext.ProcIndexDic = Nothing

    gContext.ProcInfoLoaded = False

End Sub

Private Function ExistsModule( _
            ByVal Col As Collection, _
            ByVal ModuleName As String) _
            As Boolean

    Dim V As Variant

    For Each V In Col

        If StrComp( _
            CStr(V), _
            ModuleName, _
            vbTextCompare) = 0 Then

            ExistsModule = True
            Exit Function

        End If

    Next V

End Function

