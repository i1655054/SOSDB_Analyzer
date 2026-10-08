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

' @JPName
' 画面初期化
'
' @Category
' Export
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' エクスポート画面の初期表示処理を行う
'
Private Sub UserForm_Initialize()

    SetupListView

    LoadToolList

End Sub

' @JPName
' ListView初期化
'
' @Category
' Export
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ツール一覧表示用ListViewを初期化する
'
Private Sub SetupListView()

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

' @JPName
' ツール一覧読込
'
' @Category
' Export
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ToolManagerシートから
' エクスポート対象ツール一覧を読み込む
'
' @Remarks
' WorkbookのOpen/Close状態も表示する
'
Private Sub LoadToolList()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim r As Long

    Dim itm As ListItem
    Dim wb As Workbook

    Set ws = _
        ThisWorkbook.Worksheets(SHEET_TOOL_MANAGER)

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

            Set wb = Nothing
            
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

' @JPName
' 選択件数更新
'
' @Category
' Export
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' チェック済みツール件数を表示する
'
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

' @JPName
' 一覧再読込
'
' @Category
' Export
'
' @Summary
' ツール一覧を再読込し
' Open/Close状態を更新する
'
Private Sub cmdRefresh_Click()

    LoadToolList

End Sub

' @JPName
' 全選択
'
' @Category
' Export
'
' @Summary
' 一覧の全ツールを選択する
'
Private Sub cmdSelectAll_Click()

    Dim itm As ListItem

    For Each itm In lvwTools.ListItems

        itm.Checked = True

    Next itm

    UpdateSelectedCount

End Sub

' @JPName
' 全解除
'
' @Category
' Export
'
' @Summary
' 一覧の全ツールの選択を解除する
'
Private Sub cmdSelectNone_Click()

    Dim itm As ListItem

    For Each itm In lvwTools.ListItems

        itm.Checked = False

    Next itm

    UpdateSelectedCount

End Sub

' @JPName
' チェック変更
'
' @Category
' Export
'
' @Input
' Item(ListItem)
'
' @Output
' なし
'
' @Summary
' チェック状態変更時に
' 選択件数を更新する
'
Private Sub lvwTools_ItemCheck( _
    ByVal Item As MSComctlLib.ListItem)

    UpdateSelectedCount

End Sub

' @JPName
' VBAエクスポート実行
'
' @Category
' Export
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択ツールのVBAソースを
' Git管理フォルダへ出力する
'
' @Remarks
' エクスポート対象が1件の場合は
' PowerShellを自動起動する
'
Private Sub cmdExport_Click()

    Dim itm As ListItem

    Dim ToolID As Long
    Dim wb As Workbook

    Dim Count As Long
    
    Dim ExportCount As Long
    Dim LastRootPath As String

    For Each itm In lvwTools.ListItems

        If itm.Checked Then

            ToolID = CLng(itm.Text)
            
            Set wb = _
                GetWorkbookByToolID(ToolID)

            If Not wb Is Nothing Then

                ExportCount = ExportCount + 1
                
                LastRootPath = _
                    GetGitRootPathByToolID(ToolID)

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

    If Count = 0 Then

        MsgBox _
            "選択されたツールはありません。", _
            vbInformation

        Exit Sub

    End If

    MsgBox _
        Count & " 件のエクスポートが完了しました。", _
        vbInformation
        
    If ExportCount = 1 Then
    
        OpenGitPowerShell LastRootPath
        
    End If

End Sub

' @JPName
' 画面終了
'
' @Category
' Export
'
' @Summary
' エクスポート画面を閉じる
'
Private Sub cmdClose_Click()

    Unload Me
    
    frmMainMenu.Show vbModeless

End Sub

