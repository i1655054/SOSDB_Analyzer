VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigator 
   Caption         =   "Navigator"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   OleObjectBlob   =   "frmNavigator.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mIsCollapsed As Boolean
Private mNormalWidth As Double

Private Const MENU_EXPORT As String = "VBA Export"
Private Const MENU_PATHCHART As String = "PathChart"

Private Sub UserForm_Initialize()

    mNormalWidth = Me.Width
    
    PositionControls

    RefreshSheetList
    SelectCurrentSheet

End Sub

Private Sub ExecuteMenu(ByVal MenuName As String)

    Select Case MenuName

        Case MENU_EXPORT
            
            frmExportTool.Show vbModal

        Case MENU_PATHCHART

            frmPathChart.Show vbModeless

        Case Else

            Worksheets(MenuName).Activate

    End Select

End Sub

' ×ボタン対策
Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then

        Cancel = True
        Me.Hide

    End If

End Sub

' Excel左側へ配置
Private Sub UserForm_Activate()

    Me.Left = Application.Left + 10
    Me.Top = Application.Top + 100

End Sub

' ダブルクリックで選択したメニュー実行
Private Sub lstMenu_DblClick( _
                ByVal Cancel As MSForms.ReturnBoolean)

    If lstMenu.ListIndex = -1 Then Exit Sub

    ExecuteMenu lstMenu.Value

End Sub

' 部分一致検索
Private Sub txtSearch_Change()

    Dim ws As Worksheet
    Dim SearchText As String

    SearchText = Trim$(txtSearch.Text)

    lstMenu.Clear

    For Each ws In ThisWorkbook.Worksheets

        If ws.Visible = xlSheetVisible Then

            If Not IsHiddenMenuSheet(ws.Name) Then

                If SearchText = "" _
                Or InStr(1, ws.Name, SearchText, vbTextCompare) > 0 Then

                    lstMenu.AddItem ws.Name

                End If

            End If
            
        End If

    Next ws

    If SearchText = "" _
    Or InStr(1, MENU_EXPORT, SearchText, vbTextCompare) > 0 Then

        lstMenu.AddItem MENU_EXPORT

    End If

End Sub

' 検索後に Enter キーで即実行
Private Sub txtSearch_KeyDown( _
                ByVal KeyCode As MSForms.ReturnInteger, _
                ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then

        If lstMenu.ListCount > 0 Then

            lstMenu.ListIndex = 0

            ExecuteMenu lstMenu.List(0)

        End If

    End If

End Sub

Private Sub PositionControls()

    cmdToggle.Left = Me.InsideWidth - cmdToggle.Width - 10
    cmdToggle.Top = Me.InsideHeight - cmdToggle.Height - 10

End Sub

' サイドバー開閉
Private Sub cmdToggle_Click()

    If mIsCollapsed Then

        Me.Width = mNormalWidth

        txtSearch.Visible = True
        lstMenu.Visible = True
        cmdRefresh.Visible = True
        cmdBack.Visible = True
        cmdForward.Visible = True

        cmdToggle.Caption = "<"
        
        DoEvents
        PositionControls

        mIsCollapsed = False

    Else

        Me.Width = 50

        txtSearch.Visible = False
        lstMenu.Visible = False
        cmdRefresh.Visible = False
        cmdBack.Visible = False
        cmdForward.Visible = False

        cmdToggle.Caption = ">"

        DoEvents
        PositionControls
        
        mIsCollapsed = True

    End If

    PositionControls

End Sub

' 戻るボタン
Private Sub cmdBack_Click()

    Dim SheetName As String

    SheetName = GetPreviousSheet()

    'Debug.Print "Back:", SheetName

    If Len(SheetName) > 0 Then

        IsHistoryMoving = True

        Worksheets(SheetName).Activate

        IsHistoryMoving = False

    End If

End Sub

' 進むボタン
Private Sub cmdForward_Click()

    Dim SheetName As String

    SheetName = GetNextSheet()

    'Debug.Print "Forward:", SheetName

    If Len(SheetName) > 0 Then

        IsHistoryMoving = True

        Worksheets(SheetName).Activate

        IsHistoryMoving = False

    End If

End Sub

' 更新ボタン
Private Sub cmdRefresh_Click()

    txtSearch.Text = ""
    
    RefreshSheetList
    SelectCurrentSheet

End Sub

Public Sub SelectCurrentSheet()

    Dim i As Long

    For i = 0 To lstMenu.ListCount - 1

        If lstMenu.List(i) = ActiveSheet.Name Then

            lstMenu.ListIndex = i
            Exit For

        End If

    Next i

End Sub

Public Sub RefreshSheetList()

    Dim ws As Worksheet

    lstMenu.Clear

    For Each ws In ThisWorkbook.Worksheets

        ' 非表示シートを除外する場合
        If ws.Visible = xlSheetVisible Then

            If Not IsHiddenMenuSheet(ws.Name) Then

                lstMenu.AddItem ws.Name

            End If
            
        End If

    Next ws

    lstMenu.AddItem MENU_EXPORT

End Sub

Private Function IsHiddenMenuSheet( _
                    ByVal SheetName As String) _
                    As Boolean

    Select Case SheetName

        Case "Sheet1", "ToolManager"

            IsHiddenMenuSheet = True

    End Select

End Function

