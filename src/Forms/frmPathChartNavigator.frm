VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPathChartNavigator 
   Caption         =   "PathChart Navigator"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmPathChartNavigator.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPathChartNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    LoadSearchHistory cmbSearch
    
    cmbSearch.MatchEntry = fmMatchEntryComplete
    
End Sub

Private Sub UserForm_Activate()

    Me.Left = _
        Application.Left + _
        Application.Width - _
        Me.Width - 30

    Me.Top = _
        Application.Top + 80
    
    Call ClearNavigatorInfo
    
End Sub

' ×ボタン対策
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then

        'Cancel = True
        
        frmPathChart.Show vbModeless


    End If

End Sub

' 検索
Private Sub cmdFind_Click()

    gSearchText = GetProcNameFromDisplay( _
                    Trim$(cmbSearch.text))
    
    gExactMatch = chkExactMatch.Value
    
    AddSearchHistory gSearchText

    FindNode

End Sub

' 次検索
Private Sub cmdNext_Click()

    gExactMatch = chkExactMatch.Value
    
    FindNextNode

End Sub

' 親
Private Sub cmdParent_Click()

    MoveToParentNode

End Sub

' 子
Private Sub cmdChild_Click()

    MoveToChildNode

End Sub

' 兄弟
Private Sub cmdSibling_Click()

    MoveToSiblingNode

End Sub

' 解除
Private Sub cmdClear_Click()

    ClearHighlight

End Sub

Public Sub ClearNavigatorInfo()

    txtCurrentNode.text = "-"

    lblJPName.Caption = ""

    lblCategory.Caption = ""

    lblModule.Caption = ""

    txtSummary.text = ""

    txtStatus.text = ""

End Sub

Private Sub cmdCopyCurrentInfo_Click()

    Dim s As String

    s = _
        "Procedure : " & txtCurrentNode.text & vbCrLf & _
        "JPName    : " & lblJPName.Caption & vbCrLf & _
        "Category  : " & lblCategory.Caption & vbCrLf & _
        "Module    : " & lblModule.Caption & vbCrLf & _
        "Summary   : " & txtSummary.text

    Call CopyTextToClipboard(s)

    Application.StatusBar = "現在情報をコピーしました"
    
End Sub

Private Sub CopyControlValue( _
                ByVal TextValue As String, _
                Optional ByVal Title As String = "")

    If Trim$(TextValue) = "" Then Exit Sub

    Call CopyTextToClipboard(TextValue)

    If Title <> "" Then

        Application.StatusBar = _
            "コピーしました : " & _
            Title

    Else

        Application.StatusBar = _
            "コピーしました : " & _
            TextValue

    End If

End Sub

Private Sub txtCurrentNode_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            txtCurrentNode.text, _
            "現在ノード"

    End If

End Sub

Private Sub txtStatus_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            txtStatus.text, _
            "状態"

    End If

End Sub

Private Sub lblJPName_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            lblJPName.Caption, _
            "日本語名"

    End If

End Sub

Private Sub lblCategory_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            lblCategory.Caption, _
            "カテゴリ"

    End If

End Sub

Private Sub lblModule_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            lblModule.Caption, _
            "モジュール"

    End If

End Sub

Private Sub txtSummary_MouseDown( _
        ByVal Button As Integer, _
        ByVal Shift As Integer, _
        ByVal X As Single, _
        ByVal Y As Single)

    If Button = 2 Then

        CopyControlValue _
            txtSummary.text, _
            "概要"

    End If

End Sub

