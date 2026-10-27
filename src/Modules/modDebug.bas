Attribute VB_Name = "modDebug"
Option Explicit

Public Sub Trace( _
                ByVal Category As String, _
                ByVal Message As String)

    If Not DEV_MODE Then Exit Sub

    Debug.Print _
        Format$(Now, "hh:nn:ss") & _
        " [" & Category & "] " & _
        Message

End Sub

Public Sub TraceStack( _
                ByVal Action As String, _
                ByVal StackValue As String, _
                ByVal Count As Long)

    Trace "STACK", _
        Action & _
        " [" & StackValue & "]" & _
        " Count=" & Count

End Sub
