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

' ① ExportVBA
Private Sub cmdExportVBA_Click()
    Call modExportVBA.ShowExportTool
End Sub

' ② 関数一覧作成
Private Sub ExtractProcedures_WithAnalysis_Click()
    Call modAnalyze.ExtractProcedures_WithAnalysis
End Sub

' ③ 関数トレース作成
Private Sub CreateFunctionDependency_Click()
    Call modFunctionDependency.CreateFunctionDependency
End Sub

' ④ 関数依存関係作成
Private Sub CreateFunctionTrace_Click()
    Call modFunctionTrac.CreateFunctionTrace
End Sub

' ⑤ パスチャート作成
Private Sub CreatePathChart_Click()
    frmPathChart.Show vbModeless
End Sub

