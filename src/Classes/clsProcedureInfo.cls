VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcedureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Procedure名
Public ProcedureName As String

' JPName
Public JPName As String

' HCPファイル名
Public FileName As String

' 更新日
Public LastUpdate As String

' 状態
Public Status As String
