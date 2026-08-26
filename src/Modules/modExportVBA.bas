Attribute VB_Name = "modExportVBA"
Option Explicit

Private Const GIT_REPOSITORY As String = "SOSDB_VBA"

'**************************************
'====  VBA ソース自動エクスポート  ====
'
' Excelブック内のVBAプロジェクトを、
' ボタン一発で .bas .cls .frm に出力する
'**************************************

Public Sub ExportVBA()

    Dim TargetBook As Workbook
    
    Set TargetBook = GetTargetBook()

    If Not CheckTargetBook(TargetBook) Then Exit Sub
    
    Dim VBComp As Object
    Dim RootPath As String
    Dim ModulesPath As String
    Dim ClassesPath As String
    Dim DocumentsPath As String
    Dim FormsPath As String
    Dim ExportFile As String

    RootPath = GetGitRootPath()

    ModulesPath = RootPath & "\src\Modules\"
    ClassesPath = RootPath & "\src\Classes\"
    FormsPath = RootPath & "\src\Forms\"
    DocumentsPath = RootPath & "\src\Documents\"

    For Each VBComp In TargetBook.VBProject.VBComponents

        Select Case VBComp.Type

            '標準モジュール
            Case 1
                ExportFile = ModulesPath & VBComp.Name & ".bas"
             
            'クラス
            Case 2
                ExportFile = ClassesPath & VBComp.Name & ".cls"

            'UserForm
            Case 3
                ExportFile = FormsPath & VBComp.Name & ".frm"

            'ThisWorkbook・Sheet
            Case 100
                ExportFile = DocumentsPath & VBComp.Name & ".cls"
                
            Case Else
                ExportFile = ""

        End Select
        
        If ExportFile <> "" Then
            
            If Dir(ExportFile) <> "" Then Kill ExportFile
            
            VBComp.Export ExportFile
        
        End If

    Next
    
    MsgBox "VBAソースのエクスポートが完了しました。", vbInformation

End Sub

