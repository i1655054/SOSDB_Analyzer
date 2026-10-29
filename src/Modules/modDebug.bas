Attribute VB_Name = "modDebug"
Option Explicit

'===================================================================
'@JPName     デバッグ支援
'@Category   Debug
'@Summary    デバッグログおよびスタックトレース出力を行う
'===================================================================

'===================================================================
'@JPName     トレース出力
'@Category   Debug
'@Input      Category(String);Message(String)
'@Output     なし
'@Summary    開発モード時にデバッグログを出力する
'@Remarks    DEV_MODE=False時は出力しない
'===================================================================
Public Sub Trace( _
                ByVal Category As String, _
                ByVal Message As String)
    
    If Not DEV_MODE Then Exit Sub

    Debug.Print _
        Format$(Now, "hh:nn:ss") & _
        " [" & Category & "] " & _
        Message

End Sub

'===================================================================
'@JPName     スタックトレース出力
'@Category   Debug
'@Input      Action(String);StackValue(String);Count(Long)
'@Output     なし
'@Summary    スタック操作履歴をデバッグ出力する
'@Remarks    Push/Pop確認用ログとして利用する
'===================================================================
Public Sub TraceStack( _
                ByVal Action As String, _
                ByVal StackValue As String, _
                ByVal Count As Long)

    Trace "STACK", _
        Action & _
        " [" & StackValue & "]" & _
        " Count=" & Count

End Sub
