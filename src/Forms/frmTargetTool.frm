VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTargetTool 
   Caption         =   "関数管理対象選択"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTargetTool.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmTargetTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Me.Caption = "関数管理対象選択"
    
    lstTool.ColumnCount = 2
    lstTool.ColumnWidths = "180 pt;0 pt"

    LoadToolList

End Sub

Private Sub LoadToolList()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Set ws = _
        ThisWorkbook.Worksheets("ToolManager")

    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    lstTool.Clear

    For r = 2 To LastRow

        lstTool.AddItem _
            ws.Cells(r, "B").Value

        lstTool.List( _
            lstTool.ListCount - 1, 1) = _
            ws.Cells(r, "A").Value

    Next r

End Sub


Private Sub cmdOK_Click()

    If lstTool.ListIndex < 0 Then

        MsgBox _
            "対象ツールを選択してください。"

        Exit Sub

    End If

    gTargetToolID = _
        CLng(lstTool.List( _
            lstTool.ListIndex, _
            1))

    Unload Me

    frmFunctionManager.Show vbModeless

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub
