VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=================================================
' プロシージャ情報
'=================================================

' 基本情報
Public ProcName As String
Public ModuleName As String
Public ProcType As String

' コメント情報
Public JPName As String
Public Category As String
Public InputText As String
Public OutputText As String
Public Summary As String
Public Remarks As String

' 解析情報
Public UsageCount As Long
Public Status As String

' 呼出関係
Public CalledBy As Collection
Public Calls As Collection

'=================================================
' 初期化
'=================================================
Private Sub Class_Initialize()

    Set CalledBy = New Collection
    Set Calls = New Collection

End Sub

'=================================================
' 呼出元追加
'=================================================
Public Sub AddCalledBy(ByVal ProcName As String)

    If Not ExistsInCollection(CalledBy, ProcName) Then
        CalledBy.Add ProcName
    End If

End Sub

'=================================================
' 呼出先追加
'=================================================
Public Sub AddCall(ByVal ProcName As String)

    If Not ExistsInCollection(Calls, ProcName) Then
        Calls.Add ProcName
    End If

End Sub

'=================================================
' 呼出元一覧
'=================================================
Public Function GetCalledByText() As String

    GetCalledByText = JoinCollection(CalledBy)

End Function

'=================================================
' 呼出先一覧
'=================================================
Public Function GetCallsText() As String

    GetCallsText = JoinCollection(Calls)

End Function

'=================================================
' Collection存在チェック
'=================================================
Private Function ExistsInCollection( _
            ByVal Col As Collection, _
            ByVal Value As String) As Boolean

    Dim V As Variant

    For Each V In Col

        If StrComp( _
                CStr(V), _
                Value, _
                vbTextCompare) = 0 Then

            ExistsInCollection = True
            Exit Function

        End If

    Next V

End Function

'=================================================
' Collection連結
'=================================================
Private Function JoinCollection( _
            ByVal Col As Collection) As String

    Dim Result As String
    Dim V As Variant

    For Each V In Col

        If Result <> "" Then
            Result = Result & ";"
        End If

        Result = Result & CStr(V)

    Next V

    JoinCollection = Result

End Function

