Attribute VB_Name = "modHistory"
Option Explicit

Public History As Collection
Public CurrentPos As Long

Public IsHistoryMoving As Boolean

Public Sub InitHistory()

    Set History = New Collection
    CurrentPos = 0

End Sub

' —š—ð’Ç‰Á
Public Sub AddHistory(ByVal SheetName As String)

    Dim ws As Worksheet

    Set ws = Worksheets(SheetName)

    If ws.Visible <> xlSheetVisible Then Exit Sub
    
    If ws.Name = "Sheet1" Then Exit Sub

    If History Is Nothing Then
        Set History = New Collection
    End If

    If History.count > 0 Then
        If History(History.count) = SheetName Then Exit Sub
    End If

    History.Add SheetName

    CurrentPos = History.count

End Sub

Public Function GetPreviousSheet() As String

    'Debug.Print "CurrentPos=", CurrentPos

    If History Is Nothing Then
        'Debug.Print "History Is Nothing"
        Exit Function
    End If

    If CurrentPos <= 1 Then
        'Debug.Print "CurrentPos <= 1"
        Exit Function
    End If

    CurrentPos = CurrentPos - 1

    GetPreviousSheet = History(CurrentPos)

End Function

Public Function GetNextSheet() As String

    If History Is Nothing Then Exit Function
    If CurrentPos >= History.count Then Exit Function

    CurrentPos = CurrentPos + 1

    GetNextSheet = History(CurrentPos)

End Function
