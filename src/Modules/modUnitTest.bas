Attribute VB_Name = "modUnitTest"
Option Explicit

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

    Dim X As Long

    X = FuncA()

End Sub

