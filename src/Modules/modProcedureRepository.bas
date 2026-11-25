Attribute VB_Name = "modProcedureRepository"
Option Explicit

'modProcedureRepositoryデバッグ出力有効化
#Const DEBUG_MODE = False

'===================================================================
' @JPName
' ProcedureRepository
'
' @Category
' ProcedureRepository
'
' @Summary
' ProcListシートからProcedure情報を読込み
' Procedure検索用Repositoryを管理する
'
' @Remarks
' Procedure辞書管理
' Procedure索引管理
' Procedure解決
'
' を提供する
'===================================================================

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
' 呼出元モジュールを優先して
' Procedure候補を決定する
'
' 同名Procedureが複数存在する場合は
' Procedure索引を利用して解決する
'
' 解決不能時はNothingを返却する
'
'
Public Sub InitializeProcedureRepository()

    '--------------------------
    '# 初期化状態確認
    '--------------------------
    '# Context存在確認
    If gContext Is Nothing Then Exit Sub
    
    '# ProcedureRepository初期化済確認
    If gContext.ProcInfoLoaded Then Exit Sub

    '--------------------------
    '# ProcedureRepository生成
    '--------------------------
    '# Procedure辞書生成
    BuildProcInfoDictionary

    '# ProcedureRepository初期化完了
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
    
    '# ProcListシート取得
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    On Error GoTo 0

    '# ProcList存在確認
    If ws Is Nothing Then

        MsgBox _
            "ProcListが存在しません。" & vbCrLf & _
            "先に解析実行を行ってください。", _
            vbExclamation

        Exit Sub

    End If

    '--------------------------
    '# Procedure管理辞書初期化
    '--------------------------
    Set gContext.ProcInfoDic = _
        CreateObject("Scripting.Dictionary")

    '# Procedure索引辞書初期化
    Set gContext.ProcIndexDic = _
        CreateObject("Scripting.Dictionary")

    '--------------------------
    '# ProcList読込
    '--------------------------

    '# ProcList最終行取得
    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    '# Procedure情報読込
    For r = 2 To LastRow

        '# Procedure情報生成
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

        '# Procedure検索キー生成
        Key = _
            Info.ModuleName & "|" & _
            Info.ProcName

        '# Procedure情報登録
        If Not gContext.ProcInfoDic.Exists(Key) Then

            gContext.ProcInfoDic.Add Key, Info

        Else

            ' 既存情報更新
            '# Procedure情報更新
            Set gContext.ProcInfoDic(Key) = Info

        End If

        '# Procedure索引作成
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

    '--------------------------
    '# Procedure索引取得
    '--------------------------
    If gContext.ProcIndexDic.Exists( _
            Info.ProcName) Then

        Set Col = _
            gContext.ProcIndexDic( _
                Info.ProcName)

    Else

        '# Procedure索引生成
        Set Col = New Collection

        gContext.ProcIndexDic.Add _
            Info.ProcName, _
            Col

    End If

    '--------------------------
    '# モジュール索引登録
    '--------------------------
    
    '# モジュール重複確認
    If Not ExistsModule(Col, Info.ModuleName) Then
    
        '# モジュール索引登録
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

    '# ProcedureRepository状態確認
    If gContext Is Nothing Then Exit Function

    If gContext.ProcInfoDic Is Nothing Then Exit Function

    If gContext.ProcIndexDic Is Nothing Then Exit Function

    '# Procedure名確認
    If Trim(ProcName) = "" Then Exit Function
    
    '# 呼出元優先検索キー生成
    Key = _
        CallerModule & "|" & ProcName

    '# 同一モジュールProcedure検索
    If gContext.ProcInfoDic.Exists(Key) Then

        Set ResolveProcedure = _
            gContext.ProcInfoDic(Key)

        Exit Function

    End If

    '# Procedure索引検索
    If Not gContext.ProcIndexDic.Exists( _
            ProcName) Then Exit Function

    '# Procedure候補一覧取得
    Set Col = _
        gContext.ProcIndexDic(ProcName)

    '# Procedure候補解決
    If Col.Count = 1 Then

        Key = _
            Col(1) & "|" & ProcName

        Set ResolveProcedure = _
            gContext.ProcInfoDic(Key)

        Exit Function

    End If

    #If DEBUG_MODE Then
   
        '# Procedure候補競合出力
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

    #End If

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

    '# Context存在確認
    If gContext Is Nothing Then Exit Sub
    
    '# Procedure情報辞書破棄
    Set gContext.ProcInfoDic = Nothing
    
    '# Procedure索引辞書破棄
    Set gContext.ProcIndexDic = Nothing

    '# ProcedureRepository初期化状態解除
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

    '# モジュール存在確認
    For Each V In Col

        '# モジュール一致判定
        If StrComp( _
            CStr(V), _
            ModuleName, _
            vbTextCompare) = 0 Then

            '# モジュール存在返却
            ExistsModule = True
            
            Exit Function

        End If

    Next V
    
    '# モジュール未存在
    ExistsModule = False

End Function

