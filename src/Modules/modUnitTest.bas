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
'起動確認
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

' 回帰確認
Public Sub Test_All()

    Test_IsFunctionCall

    Test_Assert_CountUsage

    Debug.Print _
        "=========="

    Debug.Print _
        "UnitTest OK"

End Sub

'==================================================
' 回帰テスト
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

    Debug.Assert _
        IsFunctionCall( _
            "If Not FP_InitAboutCalendar() Then", _
            "FP_InitAboutCalendar")

    Debug.Assert _
        IsFunctionCall( _
            "RegEx.Test(LineText)", _
            "Test") = False

    Debug.Assert _
        IsFunctionCall( _
            "Dictionary.Exists(Key)", _
            "Exists") = False
        
    Debug.Print "OK"

End Sub

Public Sub Test_Assert_CountUsage()

    Debug.Print "Test_Assert_CountUsage()"
    
    Debug.Assert _
        CountUsage("Test") = 0

    Debug.Assert _
        CountUsage("TestSub") = 0

    Debug.Assert _
        CountUsage("FuncA") = 5

    Debug.Print "OK"

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
' Test      = 0
' TestSub   = 0
' FuncA     = 5
'

Public Sub Test_CountUsage()

    Debug.Print "Test_CountUsage()"
    
    Debug.Print "Test=" & _
        CountUsage("Test")
    
    Debug.Print "TestSub=" & _
        CountUsage("TestSub")
    
    Debug.Print "FuncA=" & _
        CountUsage("FuncA")

End Sub


'==================================================
' HCPレンダリング確認
'==================================================


Public Sub Test_IfElseIf()

    Dim Value As Long

    If Value = 1 Then

        Call ProcA

    ElseIf Value = 2 Then

        Call ProcB

    Else

        Call ProcC

    End If

End Sub

Private Function ProcA() As Boolean

    ProcA = True

End Function

Private Function ProcB() As Boolean

    ProcB = True

End Function

Private Function ProcC() As Boolean

    ProcC = True

End Function


Public Sub Test_Select()

     Dim x As Integer
    Select Case x

        Case 1

            '# CASE1

            Call FuncA

        Case 2

            '# CASE2

            Call FuncA

    End Select

End Sub


Public Sub Test_Compile()

#If VBA7 Then

    '# COMPILEコメント

    Call FuncA
    
#Else

    Call FuncA

#End If

End Sub

Private Function FuncA() As Boolean

    FuncA = True

End Function

' 特殊ケース確認

Public Sub Test_ObjectMethod()

    Dim RegEx As Object

    Set RegEx = _
        CreateObject("VBScript.RegExp")

    If RegEx.Test("ABC") Then

        Debug.Print "OK"

    End If

End Sub

' If条件内関数呼出確認
Public Sub Test_InternalCall()

    If Not TestBoolFunction() Then

        Exit Sub

    End If

End Sub

Private Function TestBoolFunction() As Boolean

    TestBoolFunction = True

End Function

'Module.Procedure
Public Sub Test_ModuleCall()

    modDebug.Trace _
        "TEST", _
        "HELLO"

End Sub

