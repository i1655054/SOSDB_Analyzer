Attribute VB_Name = "modHCPInfo"
Option Explicit

'===================================================================
'@JPName      JPName・Proc情報取得
'===================================================================

'--------------------------------------------------
' Procedure情報
'--------------------------------------------------
' 将来的に以下を実装予定
'   GetProcedureInfo
'   GetProcedureStatus
'   BuildHCPInfo

'===================================================================
' @JPName
' ProcedureHTMLパス取得
'
' @Category
' HCP
'
' @Input
' ProcName : Procedure名
'
' @Output
' HtmlPath
'
' @Summary
' Procedureに対応する
' HCP HTML相対パスを返却する
'
' @Remarks
' ProcListからModule名を検索する
'===================================================================
Public Function GetProcedureHtmlPath( _
                ByVal ProcName As String) _
                As String

    '# ProcList
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# Module名
    Dim ModuleName As String

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_MODULE).End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
                Trim$(ws.Cells( _
                    r, _
                    COL_PROC).Value), _
                ProcName, _
                vbTextCompare) = 0 Then

            ModuleName = _
                Trim$(ws.Cells( _
                    r, _
                    COL_MODULE).Value)

            GetProcedureHtmlPath = _
                "../" & _
                ModuleName & "/" & _
                ProcName & "." & _
                ModuleName & _
                ".hcp.html"

            Exit Function

        End If

    Next r

End Function

'===================================================================
' @JPName
' Procedure情報取得
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
' ProcName   : Procedure名
'
' @Output
' clsProcedureInfo
'
' @Summary
' ProcListからProcedure情報を取得する
'
' @Remarks
' Procedure未存在時はNothingを返却する
'===================================================================
Public Function GetProcedureInfo( _
                ByVal ModuleName As String, _
                ByVal ProcName As String) _
                As clsProcedureInfo

    '# ProcList
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# Procedure情報
    Dim Info As clsProcedureInfo

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_MODULE).End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
                Trim$(ws.Cells(r, COL_MODULE).Value), _
                ModuleName, _
                vbTextCompare) = 0 Then

            If StrComp( _
                    Trim$(ws.Cells(r, COL_PROC).Value), _
                    ProcName, _
                    vbTextCompare) = 0 Then

                Set Info = _
                    New clsProcedureInfo

                Info.ProcedureName = _
                    ProcName

                Info.JPName = _
                    Trim$(ws.Cells( _
                        r, _
                        COL_JPNAME).Value)

                Set GetProcedureInfo = _
                    Info

                Exit Function

            End If

        End If

    Next r

End Function

'===================================================================
' @JPNameF
' ProcedureJPName取得
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
' ProcName   : Procedure名
'
' @Output
' JPName
'
' @Summary
' ProcListからJPNameを取得する
'
' @Remarks
' Procedure未存在時は
' "[JPName未設定]" を返却する
'===================================================================
Public Function GetProcedureJPName( _
                ByVal ModuleName As String, _
                ByVal ProcName As String) _
                As String

    '# ProcList
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_MODULE).End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
            Trim$(ws.Cells( _
                r, _
                COL_MODULE).Value), _
            ModuleName, _
            vbTextCompare) = 0 Then

            If StrComp( _
                Trim$(ws.Cells( _
                    r, _
                    COL_PROC).Value), _
                ProcName, _
                vbTextCompare) = 0 Then

                GetProcedureJPName = _
                    Trim$(ws.Cells( _
                        r, _
                        COL_JPNAME).Value)

                Exit Function

            End If

        End If

    Next r
    
    GetProcedureJPName = _
        "[JPName未設定]"

End Function


'===================================================================
' @JPName
' Procedure一覧取得
'
' @Category
' HCP
'
' @Input
' ModuleName : モジュール名
'
' @Output
' Collection
'
' @Summary
' モジュール配下のProcedure一覧を取得する
'
' @Remarks
' HCP_TARGETSからProcedure情報を取得し
' ProcListからJPNameを補完する
'===================================================================
Public Function GetModuleProcedureList( _
                ByVal ModuleName As String) _
                As Collection

    '# 返却一覧
    Dim Result As Collection

    '# HCP_TARGETS
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# Procedure情報
    Dim Info As clsProcedureInfo

    Set Result = New Collection

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_TARGETS)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            COL_HCP_MODULE).End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
                Trim$(ws.Cells(r, _
                COL_HCP_MODULE).Value), _
                ModuleName, _
                vbTextCompare) = 0 Then

            Set Info = _
                New clsProcedureInfo

            Info.ProcedureName = _
                Trim$(ws.Cells( _
                r, _
                COL_HCP_PROC).Value)

            Info.JPName = _
                GetProcedureJPName( _
                    ModuleName, _
                    Info.ProcedureName)

            Info.FileName = _
                Info.ProcedureName & _
                "." & _
                ModuleName

            Info.Status = _
                Trim$(ws.Cells( _
                    r, _
                    COL_HCP_STATUS).Value)

            Select Case UCase$(Info.Status)

                Case "DONE"

                    Info.Status = _
                        STATUS_OK

                Case "ERROR"

                    Info.Status = _
                        STATUS_ERROR

                Case "PENDING"

                    Info.Status = _
                        STATUS_PENDING

                Case "SKIP"

                    Info.Status = _
                        STATUS_SKIP

                Case ""

                    Info.Status = _
                        STATUS_UNKNOWN

                Case Else

                    Info.Status = _
                        "[" & _
                        Info.Status & _
                        "]"

            End Select
            
            If IsDate(ws.Cells( _
                        r, _
                        COL_HCP_UPDATE_DATE).Value) Then
                
                Info.LastUpdate = _
                    Format$( _
                        ws.Cells( _
                            r, _
                            COL_HCP_UPDATE_DATE).Value, _
                        "yyyy/mm/dd HH:nn:ss")
            
            Else

                Info.LastUpdate = "-"
            End If
            
            Result.Add Info

        End If

    Next r

    Set GetModuleProcedureList = _
        Result

End Function

'===================================================================
' @JPName
' 呼出元一覧取得
'
' @Category
' HCP
'
' @Input
' ProcName : Procedure名
'
' @Output
' Collection
'
' @Summary
' 指定Procedureを呼び出す
' Procedure一覧を取得する
'
' @Remarks
' ProcListおよび解析結果を
' 利用して呼出元を収集する
'===================================================================
Public Function GetCalledByList( _
                ByVal ProcName As String) _
                As Collection

    '# 結果
    Dim Result As Collection

    '# ProcList
    Dim ws As Worksheet

    '# 最終行
    Dim LastRow As Long

    '# 行番号
    Dim r As Long

    '# 呼出元
    Dim Caller As String

    Set Result = New Collection

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_FUNCTION_DEPENDENCY)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    For r = 2 To LastRow

        If StrComp( _
                Trim$(ws.Cells(r, 2).Value), _
                ProcName, _
                vbTextCompare) = 0 Then

            Caller = _
                Trim$( _
                    ws.Cells(r, 1).Value)

            Result.Add _
                Caller

        End If

    Next r

    Set GetCalledByList = _
        Result

End Function


