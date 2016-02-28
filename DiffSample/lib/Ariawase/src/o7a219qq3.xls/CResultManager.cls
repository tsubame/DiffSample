VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CResultManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   測定結果保存用クラス
'
'目的:
'   測定結果の保存
'
'作成者:
'   0145184004
'
Option Explicit

Private m_Results As Collection
Private m_AcceptOverwrite As Boolean
Private m_siteResult As Collection

Private Sub Class_Initialize()
    m_AcceptOverwrite = False
    Call Clear
End Sub

Public Sub Add(ByVal pName As String, ByRef pResult() As Double)
'内容:
'   データの追加
'
'[pName]       IN/OUT String型:     結果につける名前
'[pResult()]   IN/OUT Double型:     保存するデータ
'
'備考:
'   IsOverwriteModeがFalseの時は上書き禁止。
'   上書きするとエラーを発生。
'

    If IsExist(pName) Then
        If m_AcceptOverwrite = True Then
            Call Delete(pName)
            Call Add(pName, pResult)
        Else
            Call Err.Raise(999, "CResultManager.Add()", """" & pName & """ has already been registered.")
        End If
    Else
        Dim tmp As CResult
        Set tmp = New CResult
        tmp.Name = pName
        tmp.result = pResult
        Call m_Results.Add(tmp, pName)
    End If
    
End Sub
'#V21-Release
Public Sub AddBySite(ByVal pName As String, ByRef pResult() As Double, ByVal pSite As Long)
'内容:
'   データの追加,Site指定用
'
'[pName]       IN/OUT String型:     結果につける名前
'[pResult()]   IN/OUT Double型:     保存するデータ
'
'備考:
'   IsOverwriteModeがFalseの時は上書き禁止。
'   上書きするとエラーを発生。
'
    
    Dim j As Integer
    j = CountExistSite
    Dim SiteData() As Double
    ReDim SiteData(j)
    Dim tmp As CResult
    Dim sitetmp As CResult  '状態判別用

    If IsExist(pName) Then
        
        Set sitetmp = m_siteResult.Item(pName)
        
        '上書き禁止+同Site測定時のみエラー出力
        If m_AcceptOverwrite = False Then
           If sitetmp.result(pSite) = 1# Then
            Call Err.Raise(999, "CResultManager.Add()", """" & pName & """ has already been registered.")
           End If
        End If
        
        'マージ処理
        Set tmp = m_Results.Item(pName)
'        tmp.Result(pSite) = pResult(pSite)
         SiteData = tmp.result
         SiteData(pSite) = pResult(pSite)
         tmp.result = SiteData
'        sitetmp.Result(pSite) = 1#
         SiteData = sitetmp.result
         SiteData(pSite) = 1#
         sitetmp.result = SiteData
         
    Else
        Set tmp = New CResult
        tmp.Name = pName
        tmp.result = pResult
        Call m_Results.Add(tmp, pName)
                
        Dim i As Integer
        For i = 0 To j
        SiteData(i) = 0#
        Next i
        SiteData(pSite) = 1#
        
        Set sitetmp = New CResult
        sitetmp.Name = pName
        sitetmp.result = SiteData
        Call m_siteResult.Add(sitetmp, pName)
        
    End If
        
End Sub
Public Sub Delete(ByVal pName As String)
'内容:
'   登録したデータを削除
'
'[pName]       IN/OUT String型:     削除するデータの名前
'
'備考:
'   指定した名前が存在しない場合エラーを発生。
'   意図的に上書きする場合はまず削除してからすること。
'

    If IsExist(pName) Then
        Call m_Results.Remove(pName)
    Else
        Call Err.Raise(999, "CResultManager.Delete()", """" & pName & """ doesn't exist.")
    End If
    
End Sub
'#V21-Release
Public Sub SiteJudgeDelete(ByVal pName As String, ByVal pSite As Long)

    If Not IsExist(pName) Then
        Dim sitetmp As CResult '状態判別用
        Set sitetmp = m_siteResult.Item(pName)
        If sitetmp.result(pSite) = 1# Then
           Dim SiteData() As Double
           ReDim SiteData(CountExistSite)
           SiteData = sitetmp.result
           SiteData(pSite) = 0#
           sitetmp.result = SiteData
        End If
    Else
        Call Err.Raise(999, "CResultManager.SiteJudgeDelete()", """" & pName & """ doesn't exist.")
    End If

End Sub
Public Function IsExist(ByVal pName As String) As Boolean
'内容:
'   指定した名前が存在するかどうか確認
'
'[pName]       IN/OUT String型:     確認する名前
'
'備考:
'   存在する場合True、しない場合Falseを返す
'

    Dim tmp As CResult
    
On Error GoTo NOT_FOUND
    Set tmp = m_Results.Item(pName)
    IsExist = True
    
    Exit Function
    
NOT_FOUND:
    IsExist = False
    
End Function

Public Sub GetResult(ByVal pName As String, ByRef pResult() As Double)
'内容:
'   登録したデータの取得
'
'[pName]       IN/OUT String型:     取得するデータの名前
'[pResult()]   IN/OUT Double型:     データ保存用配列
'
'備考:
'   pResultは動的配列を使うこと
'   データが存在しない場合エラーを返す
'

    If IsExist(pName) Then
        Dim tmp As CResult
        
        Set tmp = m_Results.Item(pName)
        pResult = tmp.result
    Else
        Call Err.Raise(999, "CResultManager.GetResult()", """" & pName & """ doesn't exist.")
    End If
    
End Sub

Public Property Let IsOverwriteMode(ByVal RHS As Boolean)

    If m_AcceptOverwrite = False And RHS = True Then
        If MsgBox( _
            "Do you accept overwrite?", vbYesNo Or vbCritical) _
        = vbNo Then
            Exit Property
        End If
    End If
    
    m_AcceptOverwrite = RHS
    
End Property

Public Property Get IsOverwriteMode() As Boolean
    IsOverwriteMode = m_AcceptOverwrite
End Property
'#V21-Release
Public Sub Clear()
    Set m_Results = New Collection
    Set m_siteResult = New Collection
End Sub

Public Property Get Count() As Long
    Count = m_Results.Count
End Property

Public Property Get List() As String

    Dim i As Long
    
    If Count > 0 Then
        List = m_Results.Item(1).Name
        
        For i = 2 To m_Results.Count
            List = List & "," & m_Results.Item(i).Name
        Next i
    Else
        List = ""
    End If
    
End Property

