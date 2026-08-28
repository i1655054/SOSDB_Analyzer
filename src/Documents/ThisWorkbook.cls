VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    InitHistory

    RegisterShortcut
    
    LoadPathChartConfig
    
    ShowNavigator

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)

    'Debug.Print "SheetActivate"; Sh.Name

    If Not IsHistoryMoving Then
        'Debug.Print "AddHistory"; Sh.Name
        AddHistory Sh.Name
    Else
        'Debug.Print "HistoryMove"; Sh.Name
    End If

    UpdateNavigator

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    UnRegisterShortcut

End Sub

