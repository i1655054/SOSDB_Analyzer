VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHCPRenderBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===================================================================
' @JPName
' HCPレンダリング構築
'
' @Category
' HCP
'
' @Summary
' HCPNode一覧から
' RenderTextを生成する
'
' @Remarks
' TXT出力、
' HTML出力、
' Excel出力で共通利用する
'
' ノード種別に応じた
' 表示文字列(RenderText)を生成する
'===================================================================

Private mNodes As Collection

'===================================================================
' @JPName
' RenderText生成
'
' @Category
' HCP
'
' @Input
' HCPNodes(Collection)
'
' @Output
' Node.RenderText
'
' @Summary
' HCPNode一覧の
' RenderTextを生成する
'
' @Remarks
' HCP描画用文字列のみ生成する
' Excel出力は行わない
'===================================================================
Public Sub Build( _
                ByVal HCPNodes As Collection)

    Set mNodes = HCPNodes

    '# RenderText生成開始
    BuildNode _
        0, _
        "　　"

End Sub

'===================================================================
' @JPName
' RenderNode生成
'
' @Category
' HCP
'
' @Input
' ParentID(Long)
' Prefix(String)
'
' @Output
' Node.RenderText
'
' @Summary
' 指定階層のRenderTextを生成する
'
' @Remarks
' ノード種別毎に
' RenderTextを生成する
'===================================================================
Private Sub BuildNode( _
                ByVal ParentID As Long, _
                ByVal Prefix As String)

    Dim Node As clsHCPNode

    For Each Node In mNodes

        If Node.ParentID = ParentID Then

            Select Case Node.NodeType

                Case NODE_START

                    Node.RenderText = _
                        "〇│" & _
                        Node.Text

                Case NODE_END

                    Node.RenderText = _
                        Prefix & _
                        "┴" & _
                        Node.Text

                Case NODE_DECLARE

                    Node.RenderText = _
                        Prefix & _
                        "□ " & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_COMMENT

                    Node.RenderText = _
                        Prefix & _
                        "│#  " & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_IF, _
                     NODE_ELSEIF

                    BuildIfRender _
                        Node, _
                        Prefix

                Case NODE_FOR

                    Node.RenderText = _
                        Prefix & _
                        "＠┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_DO

                    Node.RenderText = _
                        Prefix & _
                        "＠┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_WITH

                    Node.RenderText = _
                        Prefix & _
                        "＠┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_WHILE

                    Node.RenderText = _
                        Prefix & _
                        "＠┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_SELECT

                    Node.RenderText = _
                        Prefix & _
                        "◆┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_COMPILE_IF

                    Node.RenderText = _
                        Prefix & _
                        "◆┐" & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

                Case NODE_FALSE

                    Node.RenderText = _
                        Prefix & _
                        "├┐ELSE"

                Case NODE_MERGE

                    ' RenderText生成しない

                Case Else

                    Node.RenderText = _
                        Prefix & _
                        GetHCPSymbol(Node) & _
                        GetDisplayText( _
                            Node, _
                            Prefix)

            End Select

        End If

    Next Node

End Sub

'===================================================================
' @JPName
' IF RenderText生成
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
' Prefix(String)
'
' @Output
' Node.RenderText
'
' @Summary
' IF系ノードの
' RenderTextを生成する
'
' @Remarks
' IFおよびELSEIFを
' 対象とする
'===================================================================
Private Sub BuildIfRender( _
                ByVal Node As clsHCPNode, _
                ByVal Prefix As String)

    If Node.NodeType = NODE_ELSEIF Then

        Node.RenderText = _
            Prefix & _
            "├┐" & _
            GetDisplayText( _
                Node, _
                Prefix)

    Else

        Node.RenderText = _
            Prefix & _
            "◆┐" & _
            GetDisplayText( _
                Node, _
                Prefix)

    End If

End Sub

