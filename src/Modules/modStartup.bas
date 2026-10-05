Attribute VB_Name = "modStartup"
Option Explicit

Public Sub InitializeAnalyzer()

    gTargetToolID = 1
    
    InitHistory

    RegisterShortcut

    LoadPathChartConfig

    frmMainMenu.Show vbModeless

End Sub
