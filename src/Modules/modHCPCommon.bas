Attribute VB_Name = "modHCPCommon"
Option Explicit

'===================================================================
' @JPName
' Business継続プレフィックス取得
'
' @Category
' HTML
'
' @Input
' RenderText(String)
'
' @Output
' Prefix(String)
'
' @Summary
' Business複数行出力時の
' HCP継続行プレフィックスを取得する
'
' @Remarks
' HCP表示記号位置を基準に
' 継続行表示用の
' 「││」プレフィックスを生成する
'
'===================================================================
Public Function GetBusinessPrefix( _
                ByVal RenderText As String) _
                As String

    Dim Pos As Long

    '# CALL記号検索
    Pos = InStr( _
                RenderText, _
                "◎")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# REFERENCE記号検索
    Pos = InStr( _
                RenderText, _
                "◇")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# NORMAL記号検索
    Pos = InStr( _
                RenderText, _
                "〇")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# DECLARE記号検索
    Pos = InStr( _
                RenderText, _
                "□")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# IF記号検索
    Pos = InStr( _
                RenderText, _
                "◆")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# FOR記号検索
    Pos = InStr( _
                RenderText, _
                "＠")

    If Pos > 0 Then

        GetBusinessPrefix = _
            Left$( _
                RenderText, _
                Pos - 1) & _
            "│"

        Exit Function

    End If

    '# 継続行生成不可
    GetBusinessPrefix = "│"

End Function

' @JPName
' HCP記号取得
'
' @Category
' HCP
'
' @Input
' Node(clsHCPNode)
'
' @Output
' Symbol(String)
'
' @Summary
' ノード種別に対応する
' HCP表示記号を取得する
'
' @Remarks
' IF、CALL、FORなどの
' ノード種別に応じた記号を返却する
'
Public Function GetHCPSymbol( _
                ByVal Node As clsHCPNode) _
                As String

    '# HCP記号判定
    Select Case Node.NodeType

        Case NODE_START
            GetHCPSymbol = "〇│"

        Case NODE_DECLARE
            GetHCPSymbol = "□ "

        Case NODE_REFERENCE
            GetHCPSymbol = "◇ "

        Case NODE_CONTROL
            GetHCPSymbol = "☆ "

        Case NODE_NORMAL
            GetHCPSymbol = "〇 "

        Case NODE_CALL
            GetHCPSymbol = "◎ "

        Case NODE_IF
            GetHCPSymbol = "◆ "
        
        Case NODE_INLINE_IF
        
            If InStr(1, _
                Node.Text, _
                "Exit ", _
                vbTextCompare) > 0 Then

                GetHCPSymbol = "◆→ "

            Else

                GetHCPSymbol = "◆ "

            End If

        Case NODE_TRUE
            GetHCPSymbol = "THEN "

        Case NODE_FALSE
            GetHCPSymbol = "ELSE "

        Case NODE_FOR
            GetHCPSymbol = "＠ "
        
        Case NODE_DO
            GetHCPSymbol = "＠ "

        Case NODE_WITH
            GetHCPSymbol = "＠ "

        Case NODE_WHILE
            GetHCPSymbol = "＠ "

        Case NODE_SELECT
            GetHCPSymbol = "◆ "

        Case NODE_CASE
            GetHCPSymbol = ""

        Case NODE_MERGE
            GetHCPSymbol = ""

        Case NODE_END
            GetHCPSymbol = "┴"
        
        Case NODE_GOTO
            GetHCPSymbol = "→ "
    
        Case NODE_EXIT
            GetHCPSymbol = "→ "
        
        Case NODE_LABEL
            GetHCPSymbol = "▼ "
        
        Case NODE_RESUME
            GetHCPSymbol = "☆ "

        Case Else
            GetHCPSymbol = ""

    End Select

End Function
