VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHCPGenerator 
   Caption         =   "HCP生成"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "frmHCPGenerator.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmHCPGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    optCompact.Value = True
    
    LoadModuleList

End Sub


Private Sub LoadModuleList()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object
    Dim Keys As Variant

    Set Dic = _
        CreateObject("Scripting.Dictionary")

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells(ws.Rows.Count, COL_MODULE).End(xlUp).row

    For r = 2 To LastRow

        If Trim$(ws.Cells(r, COL_MODULE).Value) <> "" Then

            Dic(CStr(ws.Cells(r, COL_MODULE).Value)) = True

        End If

    Next r

    Keys = Dic.Keys

    SortStringArray Keys

    lstModule.Clear

    For r = LBound(Keys) To UBound(Keys)

        lstModule.AddItem Keys(r)

    Next r

End Sub

Private Sub SortStringArray( _
                ByRef Arr As Variant)

    Dim i As Long
    Dim j As Long

    Dim Temp As String

    For i = LBound(Arr) To UBound(Arr) - 1

        For j = i + 1 To UBound(Arr)

            If StrComp( _
                    Arr(i), _
                    Arr(j), _
                    vbTextCompare) > 0 Then

                Temp = Arr(i)

                Arr(i) = Arr(j)

                Arr(j) = Temp

            End If

        Next j

    Next i

End Sub

Private Sub lstModule_Click()

    LoadProcedureList _
        lstModule.Value

End Sub

Private Sub LoadProcedureList( _
                ByVal ModuleName As String)

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim Dic As Object
    Dim Keys As Variant
    Dim i As Long

    Set Dic = _
        CreateObject("Scripting.Dictionary")

    lstProcedure.Clear

    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    For r = 2 To LastRow

        If ws.Cells(r, COL_MODULE).Value = _
                ModuleName Then

            Dic( _
                CStr(ws.Cells(r, COL_PROC).Value)) = True

        End If

    Next r

    Keys = Dic.Keys

    SortStringArray Keys

    For i = LBound(Keys) To UBound(Keys)

        lstProcedure.AddItem Keys(i)

    Next i

End Sub

Private Sub lstProcedure_Click()

    LoadProcedureInfo

End Sub

Private Sub LoadProcedureInfo()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim ModuleName As String
    Dim ProcName As String
    
    Set ws = _
        gContext.ManagementBook.Worksheets( _
            SHEET_PROC_LIST)

    LastRow = _
        ws.Cells( _
            ws.Rows.Count, _
            1).End(xlUp).row

    ModuleName = lstModule.Value
    ProcName = lstProcedure.Value

    For r = 2 To LastRow

        If ws.Cells(r, COL_MODULE).Value = _
                ModuleName _
        And ws.Cells(r, COL_PROC).Value = _
                ProcName Then

            txtJPName.Text = _
                ws.Cells(r, COL_JPNAME).Value

            txtSummary.Text = _
                ws.Cells(r, COL_SUMMARY).Value

            Exit For

        End If

    Next r

End Sub

Private Sub cmdGenerate_Click()

    Dim Tool As clsToolInfo
    
    Dim ModuleName As String
    Dim ProcName As String

    If lstModule.ListIndex < 0 Then
        MsgBox "モジュールを選択してください。"
        Exit Sub
    End If

    If lstProcedure.ListIndex < 0 Then
        MsgBox "プロシージャを選択してください。"
        Exit Sub
    End If

    ModuleName = lstModule.Value
    ProcName = lstProcedure.Value

    gHCPCompactMode = optCompact.Value

    CreateHCP _
        ModuleName, _
        ProcName
   
    Application.StatusBar = _
        "HCP生成完了 : " & _
        ModuleName & "." & ProcName
    
    Trace "HCP", _
        "Complete " & _
        ModuleName & "." & ProcName
    
    gContext.ManagementBook.Activate
    
    gContext.ManagementBook.Worksheets( _
        SHEET_HCP_CHART).Activate
    
End Sub


Private Sub cmdClose_Click()

    Unload Me

End Sub


