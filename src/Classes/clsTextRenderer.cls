VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================================================================
' @JPName
' テキストレンダラ
'
' @Category
' HCP
'
' @Summary
' HCPNodesをテキスト形式で出力する
'
' @Remarks
' RenderTextを利用して
' hcp.txt を生成する
'
' HCP本文
' Calls一覧
' Called By一覧
'
' をテキスト形式で出力する
'===================================================================

Private mFileNo As Integer

'===================================================================
' @JPName
' テキスト出力
'
' @Category
' HCP
'
' @Input
' FilePath(String)
' ProcName(String)
' HCPNodes(Collection)
'
' @Output
' hcp.txt
'
' @Summary
' HCPNodesをテキスト出力する
'
' @Remarks
' STARTノードを先頭へ出力し
' HCP本文を順次出力する
'
' 出力後に
' Calls一覧
' Called By一覧
'
' を付加出力する
'===================================================================
Public Sub Export( _
                ByVal FilePath As String, _
                ByVal ProcName As String, _
                ByVal HCPNodes As Collection)

    Dim Node As clsHCPNode
    Dim Lines() As String
    Dim i As Long

    mFileNo = FreeFile

    Open FilePath _
        For Output As #mFileNo

    '-----------------------------
    '# STARTノード出力
    '-----------------------------
    For Each Node In HCPNodes

        If Node.NodeType = NODE_START Then

            Print #mFileNo, "┬"

            If Trim$(Node.BusinessText) <> "" Then

                Print #mFileNo, _
                    Node.RenderText & _
                    vbTab & _
                    Node.BusinessText

            Else

                Print #mFileNo, _
                    Node.RenderText

            End If

            Print #mFileNo, "┴└┐"

            Exit For

        End If

    Next Node

    '-----------------------------
    '# HCP本文出力
    '-----------------------------
    For Each Node In HCPNodes

        If Node.NodeType = NODE_START Then

            GoTo ContinueLoop

        End If
        
        If Trim$(Node.RenderText) = "" Then
            
            GoTo ContinueLoop
        
        End If

        If Trim$(Node.BusinessText) = "" Then

            Print #mFileNo, _
                Node.RenderText

        Else

            Lines = Split( _
                        Node.BusinessText, _
                        vbCrLf)

            Print #mFileNo, _
                Node.RenderText & _
                vbTab & _
                Lines(0)

            For i = 1 To UBound(Lines)

                Print #mFileNo, _
                    GetBusinessPrefix( _
                        Node.RenderText) & _
                    vbTab & _
                    Lines(i)

            Next i

        End If

ContinueLoop:

    Next Node

    '-----------------------------
    '# HCP本文出力
    '-----------------------------
    '-----------------------------
    '# Calls一覧出力
    '-----------------------------
    WriteCallsText _
        mFileNo, _
        ProcName

    '-----------------------------
    '# Called By一覧出力
    '-----------------------------
    WriteCalledByText _
        mFileNo, _
        ProcName

    Close #mFileNo

End Sub



'===================================================================
' @JPName
' Calls一覧出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出先一覧を出力する
'
' @Remarks
' FunctionDependencyを参照し
' 呼出先Procedure一覧を出力する
'
' 呼出先が存在しない場合は
' 「呼出先なし」を出力する
'===================================================================
Private Sub WriteCallsText( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    '# 呼出先一覧
    Dim Calls As Collection

    '# 呼出先Procedure
    Dim Proc As Variant

    ' 呼出元Procedure情報
    Dim Info As clsProcInfo
    
    '# 呼出先一覧取得
    Set Calls = _
        GetCallList(ProcName)

    Print #FileNo, ""
    Print #FileNo, String$(40, "=")
    Print #FileNo, _
        "Calls (" & _
        Calls.Count & _
        ")"
    Print #FileNo, String$(40, "=")
    Print #FileNo, ""

    '# 呼出先なし
    If Calls.Count = 0 Then

        Print #FileNo, "呼出先なし"

        Exit Sub

    End If

    '# ヘッダ出力
    Print #FileNo, _
        "Procedure" & vbTab & _
        "JPName" & vbTab & _
        "Module"

    Print #FileNo, _
        String$(60, "-")

    '# 呼出先一覧出力
    For Each Proc In Calls

        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Proc))

        If Info Is Nothing Then

            Print #FileNo, _
                CStr(Proc)

        Else

            Print #FileNo, _
                Info.ProcName & vbTab & _
                Info.JPName & vbTab & _
                Info.ModuleName

        End If

    Next Proc

End Sub


'===================================================================
' @JPName
' Called By一覧出力
'
' @Category
' HCP
'
' @Input
' FileNo(Integer)
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 指定Procedureの
' 呼出元一覧を出力する
'
' @Remarks
' FunctionDependencyを参照し
' 呼出元Procedure一覧を出力する
'
' 呼出元が存在しない場合は
' 「呼出元なし」を出力する
'===================================================================
Private Sub WriteCalledByText( _
                ByVal FileNo As Integer, _
                ByVal ProcName As String)

    ' 呼出元一覧
    Dim Callers As Collection

    ' 呼出元Procedure
    Dim Caller As Variant

    ' 呼出元Procedure情報
    Dim Info As clsProcInfo
    
    '# 呼出元一覧取得
    Set Callers = _
        GetCalledByList( _
            ProcName)

    Print #FileNo, ""
    Print #FileNo, String$(40, "=")
    Print #FileNo, "Called By (" & _
        Callers.Count & _
        ")"
    Print #FileNo, String$(40, "=")
    Print #FileNo, ""

    '# 呼出元なし
    If Callers.Count = 0 Then

        Print #FileNo, "呼出元なし"

        Exit Sub

    End If

    '# ヘッダ出力
    Print #FileNo, _
        "Procedure" & vbTab & _
        "JPName" & vbTab & _
        "Module"

    Print #FileNo, _
        String$(60, "-")

    '# 呼出元一覧出力
    For Each Caller In Callers

        Set Info = _
            ResolveProcedure( _
                "", _
                CStr(Caller))

        If Info Is Nothing Then

            Print #FileNo, _
                CStr(Caller)

        Else

            Print #FileNo, _
                Info.ProcName & vbTab & _
                Info.JPName & vbTab & _
                Info.ModuleName

        End If

    Next Caller

End Sub

