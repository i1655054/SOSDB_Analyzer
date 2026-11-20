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

    '# ④ EXIT系
    If FuncA Then

        Exit Sub

    End If

End Sub

Public Sub TestSub()

'# Public Sub Test_Goto()

Retry:

    If FuncA Then

        GoTo Retry

    End If

End Sub

Public Sub Sample()

    MsgBox "Test"

    Dim TestFlg As Boolean

    TestFlg = True

End Sub

Public Sub Test2()

     Dim i As Integer
     Dim x As Integer

    '# ③ FOR → SELECT → IF
    For i = 1 To 3

        Select Case x

            Case 1

                If FuncA Then

                    Call FuncA

                End If

        End Select

    Next i

End Sub

Public Sub Test3()

     Dim x As Integer
    
    '# ② CASEの中にIF

    Select Case x

        Case 1

            If FuncA Then

                Call FuncA

            End If

        Case 2

            Call FuncA

    End Select

End Sub

Public Function FuncA() As Long

End Function

Public Sub Test4()

    '# ① FOR × IF × ELSE（最優先）
    '# Public Sub Test_ForIf()
    
    Dim i As Integer
    Dim A As String
    
    For i = 1 To 3

        If A Then

            Call FuncA

        Else

            Call FuncA

        End If

    Next i

End Sub

Public Sub Test_For()

     Dim i As Integer
    
    '# FOR開始コメント

    For i = 1 To 3

        '# FOR内コメント

        Debug.Print i

    Next i

End Sub

Public Sub Test_If()

    Dim A As String
    '# IF前コメント

    If A Then

        '# TRUEコメント

        Call FuncA

    Else

        '# FALSEコメント

        Call FuncA
    End If

End Sub

Public Sub Test_Do()

    Dim A As String
    Do While A

        '# DOコメント

        Call FuncA

    Loop

End Sub

Public Sub Test_With()

    Dim Obj
    With Obj

        '# WITHコメント

        .Name = "ABC"

    End With

End Sub

Public Sub Test_While()

    Dim A As String
    While A

        '# WHILEコメント

        Call FuncA

    Wend

End Sub

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

#End If

End Sub

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

Public Sub Test_ForIf()

    Dim i As Long

    For i = 1 To 10

        If i Mod 2 = 0 Then

            Call ProcA

        Else

            Call ProcB

        End If

    Next i

End Sub

Public Sub Test_SelectIf()

    Dim Value As Long

    Select Case Value

        Case 1

            If Value > 0 Then

                Call ProcA

            End If

        Case 2

            Call ProcB

        Case Else

            Call ProcC

    End Select

End Sub

Public Sub Test_Do2()

    Dim Value As Long

    Do While Value < 10

        Call ProcA

        Value = Value + 1

    Loop

End Sub

Public Sub Test_With2()

    With ActiveSheet

        .Name = "TEST"

        Call ProcA

    End With

End Sub


Public Sub Test_Compile2()

#If VBA7 Then

    Call ProcA

#Else

    Call ProcB

#End If

End Sub




