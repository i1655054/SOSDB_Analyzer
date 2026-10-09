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
'
' @JPName
' プロシージャ情報
'
' @Category
' Model
'
' @Summary
' プロシージャの属性情報および
' 呼出関係を保持する
'
' @Remarks
' ProcList、FunctionTrace、
' PathChartで共通利用する
'

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
' @JPName
' プロシージャ情報初期化
'
' @Category
' Model
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 呼出元および呼出先コレクションを初期化する
'
Private Sub Class_Initialize()

    Set CalledBy = New Collection
    Set Calls = New Collection

End Sub

'=================================================
' 呼出元追加
'=================================================
' @JPName
' 呼出元追加
'
' @Category
' Model
'
' @Input
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 呼出元プロシージャを登録する
'
' @Remarks
' 重複登録は行わない
'
Public Sub AddCalledBy(ByVal ProcName As String)

    If Not ExistsInCollection(CalledBy, ProcName) Then
        CalledBy.Add ProcName
    End If

End Sub

'=================================================
' 呼出先追加
'=================================================
' @JPName
' 呼出先追加
'
' @Category
' Model
'
' @Input
' ProcName(String)
'
' @Output
' なし
'
' @Summary
' 呼出先プロシージャを登録する
'
' @Remarks
' 重複登録は行わない
'
Public Sub AddCall(ByVal ProcName As String)

    If Not ExistsInCollection(Calls, ProcName) Then
        Calls.Add ProcName
    End If

End Sub

'=================================================
' 呼出元一覧
'=================================================
' @JPName
' 呼出元一覧取得
'
' @Category
' Model
'
' @Input
' なし
'
' @Output
' CalledByList(String)
'
' @Summary
' 呼出元一覧をセミコロン区切りで取得する
'
Public Function GetCalledByText() As String

    GetCalledByText = JoinCollection(CalledBy)

End Function

'=================================================
' 呼出先一覧
'=================================================
' @JPName
' 呼出先一覧取得
'
' @Category
' Model
'
' @Input
' なし
'
' @Output
' CallList(String)
'
' @Summary
' 呼出先一覧をセミコロン区切りで取得する
'
Public Function GetCallsText() As String

    GetCallsText = JoinCollection(Calls)

End Function

'=================================================
' Collection存在チェック
'=================================================
' @JPName
' Collection存在確認
'
' @Category
' Internal
'
' @Input
' Col(Collection)
' Value(String)
'
' @Output
' Exists(Boolean)
'
' @Summary
' Collection内に
' 指定値が存在するか判定する
'
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
' @JPName
' Collection連結
'
' @Category
' Internal
'
' @Input
' Col(Collection)
'
' @Output
' Text(String)
'
' @Summary
' Collection内容を
' セミコロン区切り文字列へ変換する
'
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

'=================================================
' 呼出元件数取得
'=================================================
' @JPName
' 呼出元件数取得
'
' @Category
' Model
'
' @Input
' なし
'
' @Output
' Count(Long)
'
' @Summary
' 呼出元プロシージャ件数を取得する
'
Public Property Get CalledByCount() As Long

    CalledByCount = CalledBy.Count

End Property

Public Property Get CallCount() As Long

    CallCount = Calls.Count

End Property

