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

