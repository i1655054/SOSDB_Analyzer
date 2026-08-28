VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPathChartOption 
   Caption         =   "PathChart詳細設定"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "frmPathChartOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPathChartOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    If gMaxNode = 0 Then

        txtMaxNode.Text = "1000"
        txtMaxEdge.Text = "3000"

        optHorizontal.Value = True
        
    End If

    chkProperty.Value = gProperty
    chkPrivate.Value = gPrivate
    chkPublic.Value = gPublic

    chkStdModule.Value = gStdModule
    chkUserForm.Value = gUserForm
    chkClassModule.Value = gClassModule

    chkIgnoreAPI.Value = gIgnoreAPI
    chkIgnoreExcel.Value = gIgnoreExcel
    chkIgnoreSelf.Value = gIgnoreSelf
    chkSameModuleOnly.Value = gSameModuleOnly

    chkSkipCommon.Value = gSkipCommon
    chkShowCount.Value = gShowCount

    chkPowerPoint.Value = gPowerPoint
    chkPngOutput.Value = gPngOutput

    chkShowModule.Value = gShowModule
    chkColorModule.Value = gColorModule

    cmbColorTheme.Clear

    cmbColorTheme.AddItem "標準"
    cmbColorTheme.AddItem "明るい"
    cmbColorTheme.AddItem "ダーク"

    If gColorTheme = "" Then

        cmbColorTheme.ListIndex = 0

    Else

        cmbColorTheme.Value = gColorTheme

    End If


    txtMaxNode.Text = gMaxNode
    txtMaxEdge.Text = gMaxEdge

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    ' 設定値保存
    
    gProperty = chkProperty.Value
    gPrivate = chkPrivate.Value
    gPublic = chkPublic.Value

    gStdModule = chkStdModule.Value
    gUserForm = chkUserForm.Value
    gClassModule = chkClassModule.Value

    gIgnoreAPI = chkIgnoreAPI.Value
    gIgnoreExcel = chkIgnoreExcel.Value
    gIgnoreSelf = chkIgnoreSelf.Value
    gSameModuleOnly = chkSameModuleOnly.Value

    gSkipCommon = chkSkipCommon.Value
    gShowCount = chkShowCount.Value

    gPowerPoint = chkPowerPoint.Value
    gPngOutput = chkPngOutput.Value

    gShowModule = chkShowModule.Value
    gColorModule = chkColorModule.Value

    gVertical = optVertical.Value

    gColorTheme = cmbColorTheme.Value

    gMaxNode = Val(txtMaxNode.Text)
    gMaxEdge = Val(txtMaxEdge.Text)

    SavePathChartConfig
    
    Me.Hide

End Sub

Private Sub cmdDefault_Click()

    chkSkipCommon.Value = False
    chkShowCount.Value = False

    chkPowerPoint.Value = False
    chkPngOutput.Value = False

    chkShowModule.Value = False
    chkColorModule.Value = False

    optHorizontal.Value = True

    cmbColorTheme.Value = "標準"

    txtMaxNode.Text = "1000"
    txtMaxEdge.Text = "3000"

End Sub

