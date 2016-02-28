VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DCシナリオが使うコンテナクラス
'
'目的:
'   各アクション間のパラメータのやり取りに使用する
'
'作成者:
'   0145206097
'
Option Explicit

Dim mDcTest As IDcTest
Dim mWaitTimer As CStopWatch
Dim mTempResult As CResultManager

Private Sub Class_Initialize()
    Me.Initialize
End Sub
'#V21-Release (mTempResult Collection型→CResultManager型)
Public Sub Initialize()
    Set mWaitTimer = New CStopWatch
    Set mTempResult = New CResultManager
End Sub

Public Property Let DcTest(ByRef dcPerformer As IDcTest)
    Set mDcTest = dcPerformer
End Property

Public Property Get DcTest() As IDcTest
    Set DcTest = mDcTest
End Property

Public Property Let WaitTimer(ByVal watch As CStopWatch)
    Set mWaitTimer = watch
End Property

Public Property Get WaitTimer() As CStopWatch
    Set WaitTimer = mWaitTimer
End Property
'#V21-Release
Public Sub AddTempResult(ByVal Keyword As String, ByRef tempResult() As Double, Optional ByVal site As Long = ALL_SITE)
    If site = ALL_SITE Then
      If mTempResult.IsExist(Keyword) = True Then
      mTempResult.Delete (Keyword)
      End If
    mTempResult.Add Keyword, tempResult()
    Else 'Site指定時
     If mTempResult.IsExist(Keyword) = True Then
     Call mTempResult.SiteJudgeDelete(Keyword, site)
     End If
    mTempResult.AddBySite Keyword, tempResult(), site
    End If
    
End Sub
'#V21-Release
Public Property Let TempResults(ByVal paramTemp As CResultManager)
    Set mTempResult = paramTemp
End Property
'#V21-Release
Public Property Get TempResults() As CResultManager
    Set TempResults = mTempResult
End Property

Private Function contains(ByVal paramTemporary As Collection, ByVal Keyword As String) As Boolean
    On Error GoTo IsContain
    paramTemporary (Keyword)
    contains = True
    Exit Function
IsContain:
    contains = False
End Function
