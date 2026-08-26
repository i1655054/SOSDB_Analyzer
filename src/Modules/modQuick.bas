Attribute VB_Name = "modQuick"
Option Explicit

Public Sub TestNavigator()

    Set History = Nothing
    CurrentPos = 0

    Unload frmNavigator

    frmNavigator.Show vbModeless

End Sub
