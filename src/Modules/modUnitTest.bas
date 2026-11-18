Attribute VB_Name = "modUnitTest"
Option Explicit

Public Sub RefreshHCPIndexes()

    Trace "INDEX", _
        "Manual Refresh"

    RebuildAllIndexes

    MsgBox _
        "INDEX更新完了", _
        vbInformation

End Sub

'==================================================
' テスト支援
'==================================================

Public Sub TestStartup()

    InitializeAnalyzer

End Sub

' @JPName
' MainMenuテスト起動
'
' @Category
' Test
'
' @Input
' なし
'
' @Output
' なし
'
' @Summary
' MainMenu画面を初期状態で表示する
'
' @Remarks
' 履歴情報を初期化してから
' MainMenuを再表示する
'
Public Sub TestNavigator()

    Dim History As Workbook
    Dim CurrentPos As Long
    
    Set History = Nothing
    CurrentPos = 0

    Unload frmMainMenu

    frmMainMenu.Show vbModeless

End Sub

'==================================================
' CountUsage回帰テスト
'==================================================
' @JPName
' CountUsageテスト
'
' @Category
' Test
'
' @Summary
' CountUsageおよびIsFunctionCallの
' 回帰テストを実行する
'
' @Remarks
' 期待値
' Test      = 2
' TestSub   = 0
' FuncA     = 1
'
Public Sub Test_CountUsage()

    Debug.Print CountUsage("Test")
    Debug.Print CountUsage("TestSub")
    Debug.Print CountUsage("FuncA")

End Sub


'==================================================
' 呼出判定回帰テスト
'==================================================
' @JPName
' 関数呼出判定テスト
'
' @Category
' Test
'
' @Summary
' IsFunctionCall
' 回帰テストを実行する
'
' @Remarks
' Debug.Assertですべてのテストが
' 成功した場合のみOKを出力する
'
'
Public Sub Test_IsFunctionCall()

    Debug.Assert _
        IsFunctionCall( _
            "Call Test", _
            "Test") = True
    
    Debug.Assert _
        IsFunctionCall( _
            "Call test", _
            "Test") = True

    Debug.Assert _
        IsFunctionCall( _
        "Call TEST", _
        "Test") = True

    Debug.Assert _
        IsFunctionCall( _
            "Call TestSub", _
            "Test") = False

    Debug.Assert _
        IsFunctionCall( _
            "MsgBox ""Test""", _
            "Test") = False

    Debug.Assert _
        IsFunctionCall( _
            "Calc = 1", _
            "Calc") = False

    Debug.Print "OK"

End Sub


Public Sub Test()

End Sub

Public Sub TestSub()

    Test

End Sub

Public Sub Sample()

    MsgBox "Test"

    Dim TestFlg As Boolean

    TestFlg = True

End Sub

Public Sub Test2()

    'Test()

End Sub

Public Sub Test3()

    Call Test

End Sub

Public Function FuncA() As Long

End Function

Public Sub Test4()

    Dim x As Long

    x = FuncA()

End Sub

Public Sub Test_For()

    '# FOR開始コメント

    For i = 1 To 3

        '# FOR内コメント

        Debug.Print i

    Next i

End Sub

Public Sub Test_If()

    '# IF前コメント

    If A Then

        '# TRUEコメント

        Call Proc1

    Else

        '# FALSEコメント

        Call Proc2

    End If

End Sub

Public Sub Test_Do()

    Do While A

        '# DOコメント

        Call Proc1

    Loop

End Sub

Public Sub Test_With()

    With Obj

        '# WITHコメント

        .Name = "ABC"

    End With

End Sub

Public Sub Test_While()

    While A

        '# WHILEコメント

        Call Proc1

    Wend

End Sub

Public Sub Test_Select()

    Select Case x

        Case 1

            '# CASE1

            Call Proc1

        Case 2

            '# CASE2

            Call Proc2

    End Select

End Sub

Public Sub Test_Compile()

#If VBA7 Then

    '# COMPILEコメント

    Call Proc1

#End If

End Sub




