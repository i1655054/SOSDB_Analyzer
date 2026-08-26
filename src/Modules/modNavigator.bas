Attribute VB_Name = "modNavigator"
Option Explicit

Public Sub ShowNavigator()

    If Not frmNavigator.Visible Then
        frmNavigator.Show vbModeless
    End If

End Sub

Public Sub UpdateNavigator()

    If frmNavigator.Visible Then

        frmNavigator.SelectCurrentSheet

    End If

End Sub
