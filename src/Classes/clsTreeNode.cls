VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================
' TreeNode情報
'====================================================
'
' @JPName
' ツリーノード
'
' @Category
' Model
'
' @Summary
' PathChartTree出力用の
' ツリー情報を保持する
'
' @Remarks
' Procedure情報および
' Tree表示情報を保持する
'

Private mLevel As Long

Private mProcName As String
Private mParentProcName As String

Private mJPName As String
Private mCategory As String
Private mModuleName As String
Private mSummary As String

Private mTreeText As String

Private mJPLine As String

'====================================================
' 日本語Tree行
'====================================================
Public Property Get JPLine() As String

    JPLine = mJPLine

End Property

Public Property Let JPLine( _
                ByVal Value As String)

    mJPLine = Value

End Property

'====================================================
' 階層レベル
'====================================================
Public Property Get Level() As Long

    Level = mLevel

End Property

Public Property Let Level( _
                ByVal Value As Long)

    mLevel = Value

End Property

'====================================================
' プロシージャ名
'====================================================
Public Property Get ProcName() As String

    ProcName = mProcName

End Property

Public Property Let ProcName( _
                ByVal Value As String)

    mProcName = Value

End Property

'====================================================
' 親プロシージャ名
'====================================================
Public Property Get ParentProcName() As String

    ParentProcName = mParentProcName

End Property

Public Property Let ParentProcName( _
                ByVal Value As String)

    mParentProcName = Value

End Property

'====================================================
' 日本語名
'====================================================
Public Property Get JPName() As String

    JPName = mJPName

End Property

Public Property Let JPName( _
                ByVal Value As String)

    mJPName = Value

End Property

'====================================================
' カテゴリ
'====================================================
Public Property Get Category() As String

    Category = mCategory

End Property

Public Property Let Category( _
                ByVal Value As String)

    mCategory = Value

End Property

'====================================================
' モジュール名
'====================================================
Public Property Get ModuleName() As String

    ModuleName = mModuleName

End Property

Public Property Let ModuleName( _
                ByVal Value As String)

    mModuleName = Value

End Property

'====================================================
' 概要
'====================================================
Public Property Get Summary() As String

    Summary = mSummary

End Property

Public Property Let Summary( _
                ByVal Value As String)

    mSummary = Value

End Property

'====================================================
' Tree表示文字列
'====================================================
Public Property Get TreeText() As String

    TreeText = mTreeText

End Property

Public Property Let TreeText( _
                ByVal Value As String)

    mTreeText = Value

End Property


