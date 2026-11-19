VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHtmlRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================================================================
' @JPName
' HTMLレンダラ
'
' @Category
' HCP
'
' @Summary
' HCPNodesをHTML形式で出力する
'
' @Remarks
' RenderTextを利用して
' hcp.html を生成する
'
' Procedureリンク
' Comment色分け
' Calls
' Called By
'
' を出力する
'===================================================================

Private mFileNo As Integer

'===================================================================
' @JPName
' HTML出力
'
' @Category
' HCP
'
' @Input
' FilePath(String)
' ModuleName(String)
' ProcName(String)
' HCPNodes(Collection)
'
' @Output
' hcp.html
'
' @Summary
' HCPNodesをHTML出力する
'===================================================================
Public Sub Export( _
                ByVal FilePath As String, _
                ByVal ModuleName As String, _
                ByVal ProcName As String, _
                ByVal HCPNodes As Collection)

    mFileNo = FreeFile

    Open FilePath _
        For Output As #mFileNo

    WriteHtmlHeader _
        mFileNo, _
        ProcName, _
        1.2, _
        True

    Print #mFileNo, _
        "<h1>" & _
        ProcName & _
        "</h1>"

    Print #mFileNo, _
        "<p>Module : " & _
        ModuleName & _
        "</p>"

    Print #mFileNo, _
        "<table>"

    Call OutputHCPNodes(HCPNodes)

    Print #mFileNo, _
        "</table>"

    '# Calls出力
    WriteCallsHtml _
        mFileNo, _
        ProcName

    '# Called By出力
    WriteCalledByHtml _
        mFileNo, _
        ProcName

    WriteHtmlFooter _
        mFileNo, _
        True

    Close #mFileNo

End Sub


