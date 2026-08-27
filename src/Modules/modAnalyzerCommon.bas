Attribute VB_Name = "modAnalyzerCommon"
Option Explicit

' 共通関数

' 対象ブックの状態確認
Public Function CheckTargetBook( _
                    ByVal TargetBook As Workbook, _
                    Optional ByVal BookName As String = "") _
                    As Boolean

    If TargetBook Is Nothing Then

        If BookName = "" Then
            MsgBox _
                "対象ブックが開かれていません。", _
                vbExclamation
        Else
            MsgBox _
                BookName & _
                " が開かれていません。", _
                vbExclamation
        End If
                

        Exit Function

    End If

    CheckTargetBook = True

End Function
