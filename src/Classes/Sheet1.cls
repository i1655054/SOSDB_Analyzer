VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub CommandButton1_Click()

End Sub

' ‡@ ExportVBA
Private Sub GetTargetBook_Click()
    Call modExportVBA.ExportVBA
End Sub

' ‡A ŠÖ”ˆê——ì¬
Private Sub ExtractProcedures_WithAnalysis_Click()
    Call modAnalyze.ExtractProcedures_WithAnalysis
End Sub

' ‡B ŠÖ”ƒgƒŒ[ƒXì¬
Private Sub CreateFunctionDependency_Click()
    Call modFunctionDependency.CreateFunctionDependency
End Sub

' ‡C ŠÖ”ˆË‘¶ŠÖŒWì¬
Private Sub CreateFunctionTrace_Click()
    Call modFunctionTrac.CreateFunctionTrace
End Sub

' ‡D ƒpƒXƒ`ƒƒ[ƒgì¬
Private Sub CreatePathChart_Click()
    Call modPathChart.CreatePathChart
End Sub

