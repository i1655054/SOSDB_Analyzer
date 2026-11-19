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
' Calls
' Called By
'
' の一覧も出力する
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

    '# HCP出力
    For Each Node In HCPNodes

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

    '# Calls出力
    WriteCallsText _
        mFileNo, _
        ProcName

    '# Called By出力
    WriteCalledByText _
        mFileNo, _
        ProcName

    Close #mFileNo

End Sub


