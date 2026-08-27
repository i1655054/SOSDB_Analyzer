VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportTool 
   Caption         =   "VBA Export Tool"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "frmExportTool.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmExportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fraToolList_Click()

End Sub

Private Sub UserForm_Initialize()

    SetupListView

    LoadToolList

End Sub

Private Sub SetupListView()

    Debug.Print "Start Setup"

    With lvwTools

        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True

        .ColumnHeaders.Clear

        .ColumnHeaders.Add , , "ID", 35
        .ColumnHeaders.Add , , "ツール名", 160
        .ColumnHeaders.Add , , "状態", 60
        .ColumnHeaders.Add , , "Repository", 120

    End With

End Sub

Private Sub LoadToolList()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim itm As ListItem
    Dim wb As Workbook

    Set ws = _
        ThisWorkbook.Worksheets("ToolManager")

    LastRow = _
        ws.Cells(ws.Rows.Count, "A") _
          .End(xlUp).row

    lvwTools.ListItems.Clear

    For r = 2 To LastRow

        If ws.Cells(r, "E").Value = True Then

            Set itm = _
                lvwTools.ListItems.Add( _
                    , , _
                    ws.Cells(r, "A").Value)

            itm.SubItems(1) = _
                ws.Cells(r, "B").Value

            On Error Resume Next

            Set wb = _
                Workbooks(ws.Cells(r, "C").Value)

            If wb Is Nothing Then

                itm.SubItems(2) = "Close"

            Else

                itm.SubItems(2) = "Open"

            End If

            Set wb = Nothing

            On Error GoTo 0

            itm.SubItems(3) = _
                ws.Cells(r, "D").Value

        End If

    Next r

    UpdateSelectedCount

End Sub

' 選択件数更新
Private Sub UpdateSelectedCount()

    Dim itm As ListItem
    Dim Count As Long

    For Each itm In lvwTools.ListItems

        If itm.Checked Then

            Count = Count + 1

        End If

    Next itm

    lblSelectedCount.Caption = _
        "選択: " & Count & " 件"

End Sub

' 更新 : Open/Close状態を再読込
Private Sub cmdRefresh_Click()

    LoadToolList

End Sub

' 全選択 : 全件チェック
Private Sub cmdSelectAll_Click()

    Dim itm As ListItem

    For Each itm In lvwTools.ListItems

        itm.Checked = True

    Next itm

    UpdateSelectedCount

End Sub

' 全解除 : 全件チェック解除
Private Sub cmdSelectNone_Click()

    Dim itm As ListItem

    For Each itm In lvwTools.ListItems

        itm.Checked = False

    Next itm

    UpdateSelectedCount

End Sub

' チェック変更時
Private Sub lvwTools_ItemCheck( _
    ByVal Item As MSComctlLib.ListItem)

    UpdateSelectedCount

End Sub

' Export : エクスポート実行
Private Sub cmdExport_Click()

    Dim itm As ListItem

    Dim ToolID As Long
    Dim wb As Workbook

    Dim Count As Long

    Debug.Print "Export Start"

    For Each itm In lvwTools.ListItems

        If itm.Checked Then

            ToolID = CLng(itm.Text)
            
            Debug.Print "ToolID=" & ToolID

            Set wb = _
                GetWorkbookByToolID(ToolID)

            If Not wb Is Nothing Then

                Debug.Print "Workbook=" & wb.Name
                
                ExportVBABook wb, ToolID

                Count = Count + 1

            Else

                Debug.Print "Workbook Not Found"
                
                MsgBox _
                    "対象ブックが開かれていません。" _
                    & vbCrLf _
                    & itm.SubItems(1), _
                    vbExclamation

            End If

        End If

    Next itm

    MsgBox _
        Count & " 件のエクスポートが完了しました。", _
        vbInformation

End Sub

' 閉じる : フォーム終了
Private Sub cmdClose_Click()

    Unload Me

End Sub

