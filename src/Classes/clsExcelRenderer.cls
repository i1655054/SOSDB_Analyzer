VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================================================================
' @JPName
' Excelレンダラ
'
' @Category
' HCP
'
' @Summary
' HCPNodesを
' HCPChartシートへ出力する
'
' @Remarks
' RenderTextを利用して
' HCPChartシートを生成する
'
' Excel表示専用レンダラ
'===================================================================

Private mWorksheet As Worksheet
Private mRowOut As Long

'===================================================================
' @JPName
' Excel出力
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
' ModuleName(String)
' ProcName(String)
'
' @Output
' HCPChartシート
'
' @Summary
' HCPNodesを
' HCPChartシートへ出力する
'
' @Remarks
' RenderTextおよび
' BusinessTextを利用して
' HCPChartシートを生成する
'===================================================================
Public Sub Export( _
                ByVal HCPNodes As Collection, _
                ByVal ModuleName As String, _
                ByVal ProcName As String)

    ' HCPノード
    Dim Node As clsHCPNode

    ' Procedure情報
    Dim ProcInfo As clsProcInfo

    '# HCPChart取得
    Set mWorksheet = _
        gContext.ManagementBook.Worksheets( _
            SHEET_HCP_CHART)

    '# Procedure情報取得
    Set ProcInfo = _
        ResolveProcedure( _
            ModuleName, _
            ProcName)

    '# HCPChart初期化
    mWorksheet.Cells.Clear

    '# ヘッダ出力
    mWorksheet.Cells(1, 1).Value = _
        "Procedure: " & _
        ModuleName & "." & _
        ProcName

    mWorksheet.Cells(2, 1).Value = _
        "HCP"

    mWorksheet.Cells(2, 2).Value = _
        "Business"

    mRowOut = 3

    '# START記号出力
    mWorksheet.Cells(mRowOut, 1).Value = "┬"

    mRowOut = mRowOut + 1

    '# HCPNode出力
    For Each Node In HCPNodes

        Call OutputNode(Node)

    Next Node

    '# 書式設定
    mWorksheet.Columns("A:B").AutoFit

End Sub

'===================================================================
' @JPName
' ノード出力
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
'
' @Output
' HCPChart
'
' @Summary
' HCPノードを
' HCPChartへ出力する
'
' @Remarks
' RenderTextおよび
' BusinessTextを出力する
'===================================================================
Private Sub OutputNode( _
                ByVal Node As clsHCPNode)

    Dim Lines() As String
    Dim i As Long

    '# RenderText未生成
    If Trim$(Node.RenderText) = "" Then

        Exit Sub

    End If

    '# Businessなし
    If Trim$(Node.BusinessText) = "" Then

        mWorksheet.Cells(mRowOut, 1).Value = _
            Node.RenderText

        mRowOut = mRowOut + 1

        Exit Sub

    End If

    '# Business分割
    Lines = Split( _
                Node.BusinessText, _
                vbCrLf)

    '# 1行目
    mWorksheet.Cells(mRowOut, 1).Value = _
        Node.RenderText

    mWorksheet.Cells(mRowOut, 2).Value = _
        Lines(0)

    mRowOut = mRowOut + 1

    '# 2行目以降
    For i = 1 To UBound(Lines)

        mWorksheet.Cells(mRowOut, 1).Value = _
            GetBusinessPrefix( _
                Node.RenderText)

        mWorksheet.Cells(mRowOut, 2).Value = _
            Lines(i)

        mRowOut = mRowOut + 1

    Next i

End Sub


