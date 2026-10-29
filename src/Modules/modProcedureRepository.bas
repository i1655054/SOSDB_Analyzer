Attribute VB_Name = "modProcedureRepository"
Option Explicit

'==================================================
' ProcedureRepository
'
' ProcListシートからProcedure情報を読み込み
' Procedure検索用Dictionaryを管理する
'==================================================

' @JPName
' ProcedureRepository初期化
'
' @Category
' ProcedureRepository
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ProcedureRepositoryを初期化する
'
' @Remarks
' Procedure情報未読込時のみ
' ProcInfo辞書を生成する
'
Public Sub InitializeProcedureRepository()

    If gContext Is Nothing Then Exit Sub
    
    If gContext.ProcInfoLoaded Then Exit Sub

    BuildProcInfoDictionary

    If Not gContext.ProcInfoDic Is Nothing Then
        
        gContext.ProcInfoLoaded = True
    
    End If

End Sub

' @JPName
' Procedure辞書作成
'
' @Category
' ProcedureRepository
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ProcListシートからProcedure情報辞書を生成する
'
' @Remarks
' ProcInfoDicおよびProcIndexDicを生成し
' Procedure情報を格納する
'
Public Sub BuildProcInfoDictionary()

    ' ProcListシート
    Dim ws As Worksheet

    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' Procedure情報
    Dim Info As clsProcInfo
    
    ' Dictionary格納キー
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

            ' 既存情報更新
            Set gContext.ProcInfoDic(Key) = Info

        End If

        BuildProcIndex Info

    Next r

End Sub

' @JPName
' Procedure索引作成
'
' @Category
' ProcedureRepository
'
' @Input
' Info(clsProcInfo)
'
' @Output
' なし
'
' @Summary
' Procedure名検索用索引を作成する
'
' @Remarks
' Procedure名から候補モジュール一覧を
' 取得するための索引を生成する
'
Private Sub BuildProcIndex( _
                ByVal Info As clsProcInfo)

    ' モジュール名一覧
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


' @JPName
' Procedure解決
'
' @Category
' ProcedureRepository
'
' @Input
' CallerModule(String)
' ProcName(String)
'
' @Output
' ProcInfo(clsProcInfo)
'
' @Summary
' 呼出先Procedure情報を取得する
'
' @Remarks
' 呼出元モジュールを優先して
' Procedure候補を決定する
'
Public Function ResolveProcedure( _
                ByVal CallerModule As String, _
                ByVal ProcName As String) _
                As clsProcInfo
  
    ' Procedure検索キー
    Dim Key As String
    
    ' 候補モジュール一覧
    Dim Col As Collection
    
    ' 候補モジュール名
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


' @JPName
' ProcedureRepository初期化解除
'
' @Category
' ProcedureRepository
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ProcedureRepositoryをクリアする
'
' @Remarks
' 再解析時にProcedure情報および
' 索引情報を破棄する
'
Public Sub ClearProcedureRepository()

    If gContext Is Nothing Then Exit Sub
    
    Set gContext.ProcInfoDic = Nothing
    Set gContext.ProcIndexDic = Nothing

    gContext.ProcInfoLoaded = False

End Sub

' @JPName
' モジュール存在判定
'
' @Category
' ProcedureRepository
'
' @Input
' Col(Collection)
' ModuleName(String)
'
' @Output
' Boolean
'
' @Summary
' 指定モジュールが一覧内に存在するか判定する
'
' @Remarks
' 大文字小文字を区別せずに比較する
'
Private Function ExistsModule( _
            ByVal Col As Collection, _
            ByVal ModuleName As String) _
            As Boolean

    ' コレクション要素
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

