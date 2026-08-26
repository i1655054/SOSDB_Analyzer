VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ComboBoxSearch_Change()
    'MsgBox "ComboBoxSearch"
End Sub

Private Sub CommandButton1_Click()
    Call modPathChart.CreatePathChart_V3
End Sub

Private Sub CommandButton2_Click()
    Call modPowerPointChart.CreatePowerPointPathChart
End Sub

Private Sub CommandButtonClearHighlight_Click()
    Call ClearHighlight
End Sub

Private Sub CommandButtonClearSearch_Click()
    Call ClearSearchHistory
End Sub

Private Sub CommandButtonFind_Click()
    Call FindNode
End Sub

Private Sub CommandButtonFindNext_Click()
    Call FindNextNode
End Sub

Private Sub CommandButtonParent_Click()
    Call MoveToParentNode
End Sub

Private Sub Worksheet_Activate()

    If ComboBoxSearch.ListCount = 0 Then

        ComboBoxSearch.AddItem "BuildDescription"
        ComboBoxSearch.AddItem "RestoreColor"
        ComboBoxSearch.AddItem "GetServerURL"

    End If

End Sub


