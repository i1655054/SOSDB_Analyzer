Attribute VB_Name = "modExportVBA"
Option Explicit

' VBAエクスポート

Public Sub ShowExportTool()

    frmExportTool.Show vbModal

End Sub

'**************************************
'====  VBA ソース自動エクスポート  ====
'
' Excelブック内のVBAプロジェクトを、
' ボタン一発で .bas .cls .frm に出力する
'**************************************
Public Sub ExportVBABook( _
                ByVal TargetBook As Workbook, _
                ByVal ToolID As Long)
    
    Dim VBComp As Object

    Dim RootPath As String
    Dim ModulesPath As String
    Dim ClassesPath As String
    Dim FormsPath As String
    Dim DocumentsPath As String
    Dim FSO As Object

    Dim ExportFile As String

    RootPath = GetGitRootPathByToolID(ToolID)
    
    ModulesPath = RootPath & "\src\Modules\"
    ClassesPath = RootPath & "\src\Classes\"
    FormsPath = RootPath & "\src\Forms\"
    DocumentsPath = RootPath & "\src\Documents\"

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FolderExists(ModulesPath) Then
        FSO.CreateFolder ModulesPath
    End If

    If Not FSO.FolderExists(ClassesPath) Then
        FSO.CreateFolder ClassesPath
    End If

    If Not FSO.FolderExists(FormsPath) Then
        FSO.CreateFolder FormsPath
    End If

    If Not FSO.FolderExists(DocumentsPath) Then
        FSO.CreateFolder DocumentsPath
    End If

    #If DEBUG_MODE Then
        Debug.Print "ModulesPath=" & ModulesPath
        Debug.Print "ClassesPath=" & ClassesPath
        Debug.Print "FormsPath=" & FormsPath
        Debug.Print "DocumentsPath=" & DocumentsPath
    #End If

    For Each VBComp In TargetBook.VBProject.VBComponents

        Select Case VBComp.Type

            Case 1
                ExportFile = ModulesPath & VBComp.Name & ".bas"

            Case 2
                ExportFile = ClassesPath & VBComp.Name & ".cls"

            Case 3
                ExportFile = FormsPath & VBComp.Name & ".frm"

            Case 100
                ExportFile = DocumentsPath & VBComp.Name & ".cls"

            Case Else
                ExportFile = ""

        End Select

        If ExportFile <> "" Then

            If Dir(ExportFile) <> "" Then

                Kill ExportFile

            End If

            #If DEBUG_MODE Then
                Debug.Print "--------------------------------"
                Debug.Print "VBComp.Name=" & VBComp.Name
                Debug.Print "ExportFile=" & ExportFile
                Debug.Print "Exists=" & (Dir$(Left$(ExportFile, InStrRev(ExportFile, "\") - 1), vbDirectory) <> "")
            #End If
            
            VBComp.Export ExportFile

        End If

    Next VBComp

End Sub

Public Function GetGitRootPathByToolID( _
            ByVal ToolID As Long) As String

    Dim Tool As ToolInfo

    Tool = GetToolInfo(ToolID)

    GetGitRootPathByToolID = _
        Environ$("OneDriveCommercial") & _
        "\ドキュメント\GitHub\" & _
        Tool.GitRepository

End Function

Public Sub OpenGitPowerShell( _
                    ByVal RootPath As String)

    Dim Cmd As String

    Cmd = _
        "powershell.exe -NoExit -Command " & _
        """Set-Location '" & RootPath & "'"""

    Shell Cmd, vbNormalFocus

End Sub
