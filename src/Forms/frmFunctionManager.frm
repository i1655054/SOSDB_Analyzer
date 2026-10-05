VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFunctionManager 
   Caption         =   "関数管理"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3435
   OleObjectBlob   =   "frmFunctionManager.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmFunctionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    Me.Caption = "関数管理"
    
    Dim Tool As ToolInfo

    Tool = GetCurrentToolInfo()

    lblTool.Caption = _
        "対象 : " & Tool.ToolName & vbCrLf & _
        "Book : " & Tool.WorkbookName

End Sub

Private Sub cmdProcedureList_Click()

    ProcedureList

End Sub


Private Sub cmdFunctionTrace_Click()

    FunctionTrace

End Sub


Private Sub cmdFunctionDependency_Click()

    FunctionDependency

End Sub


Private Sub cmdPathChart_Click()

    frmPathChart.Show vvbModelessbModal

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub


