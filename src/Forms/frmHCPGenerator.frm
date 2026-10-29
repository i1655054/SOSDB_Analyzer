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

' @JPName
' 画面初期化
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' HCP生成画面の初期化を行う
'
' @Remarks
' Compactモードを初期設定し
' モジュール一覧を読み込む
'
Private Sub UserForm_Initialize()

    optCompact.Value = True
    
    LoadModuleList

End Sub

' @JPName
' モジュール一覧読込
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' ProcListからモジュール一覧を取得する
'
' @Remarks
' 重複を除外したモジュール名を
' ソートして一覧へ表示する
'
Private Sub LoadModuleList()

    ' ProcListシート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' モジュール名管理Dictionary
    Dim Dic As Object
    
    ' ソート対象キー配列
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

' @JPName
' 文字列配列ソート
'
' @Category
' HCPGenerator
'
' @Input
' Arr(Variant)
'
' @Output
' なし
'
' @Summary
' 文字列配列を昇順に並び替える
'
' @Remarks
' 大文字小文字を区別せずに
' 配列要素をソートする
'
Private Sub SortStringArray( _
                ByRef Arr As Variant)

    ' 比較用Index
    Dim i As Long
    
    ' 比較用Index
    Dim j As Long

    ' 交換用退避文字列
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

' @JPName
' モジュール選択
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択モジュールのプロシージャ一覧を表示する
'
' @Remarks
' モジュール変更時に
' プロシージャ一覧を再読込する
'
Private Sub lstModule_Click()

    LoadProcedureList _
        lstModule.Value

End Sub

' @JPName
' プロシージャ一覧読込
'
' @Category
' HCPGenerator
'
' @Input
' ModuleName(String)
'
' @Output
' なし
'
' @Summary
' 指定モジュールのプロシージャ一覧を取得する
'
' @Remarks
' 重複を除外して
' ソート後に一覧へ表示する
'
Private Sub LoadProcedureList( _
                ByVal ModuleName As String)

    ' ProcListシート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' プロシージャ名管理Dictionary
    Dim Dic As Object
    
    ' ソート対象キー配列
    Dim Keys As Variant
    
    ' 配列Index
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

' @JPName
' プロシージャ選択
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択プロシージャの情報を表示する
'
' @Remarks
' 日本語名および概要を
' 画面へ表示する
'
Private Sub lstProcedure_Click()

    LoadProcedureInfo

End Sub

' @JPName
' プロシージャ情報読込
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' 選択プロシージャの情報を取得する
'
' @Remarks
' ProcListから日本語名および
' Summaryを取得して表示する
'
Private Sub LoadProcedureInfo()

    ' ProcListシート
    Dim ws As Worksheet
    
    ' 最終行
    Dim LastRow As Long
    
    ' 行Index
    Dim r As Long

    ' 選択モジュール名
    Dim ModuleName As String
    
    ' 選択プロシージャ名
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

' @JPName
' HCP生成
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' HCPチャートを生成する
'
' @Remarks
' 選択プロシージャを解析し
' HCPシートを生成して表示する
'
Private Sub cmdGenerate_Click()

    ' 対象モジュール名
    Dim ModuleName As String
    
    ' 対象プロシージャ名
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

' @JPName
' 画面終了
'
' @Category
' HCPGenerator
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' HCP生成画面を閉じる
'
' @Remarks
' フォームをアンロードして終了する
'
Private Sub cmdClose_Click()

    Unload Me

End Sub


