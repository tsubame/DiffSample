VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_ApmuUBChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   コンディションチェッカ　APMU UBから情報取得モジュール
'
'目的:
'   そのときの条件設定をAPMUUBから取得する
'
'注意事項:
'   IGXLのバージョンによってだめになるリスクを負っている
'
'作成者:
'   丸山
'   有川 編集 Ver1.1 2013/02/25
'
Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用

Private Const THIS_CLASS_NAME As String = "CEeeAuto_ApmuUBChecker"

Private m_IsEnable As Boolean

Private Const MAX_BIT_CH = 255         ' UBCHの最大値
Private Const INDEX_CELL_TOP = "G3" '"INDEX_CELL_TOP"

Private m_aryBefore(MAX_BIT_CH) As Long
Private m_aryAfter(MAX_BIT_CH) As Long

Implements IEeeAuto_TestConditionChecker

Private Sub Class_Initialize()
    m_IsEnable = ChecktApmuUbModify
    
#If CUB_UB_USE <> 0 Then 'CUBがあったら殺しておく
    m_IsEnable = False
#End If

End Sub

Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()
    
    If Not m_IsEnable Then Exit Sub

    Dim lBolardNum As Long
    lBolardNum = TheUB.AsAPMU.BoardNumber
    
    If Not GetApmuUb(m_aryBefore) Then
        TheError.Raise 9999, THIS_CLASS_NAME & " " & "CheckBeforeCondition", "GetUB FAIL"
    End If
    
End Sub

Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean

    If Not m_IsEnable Then
        IEeeAuto_TestConditionChecker_CheckAfterCondition = True
        Exit Function
    End If
    
    Dim lBolardNum As Long
    lBolardNum = TheUB.AsAPMU.BoardNumber
    
    If Not GetApmuUb(m_aryAfter) Then
         TheError.Raise 9999, THIS_CLASS_NAME & " " & "CheckAfterCondition", "GetUB FAIL"
   End If

    '条件が一緒だった場合 InValidとする
    Dim i As Long
    Dim IsValid As Boolean
    IsValid = False
    For i = 0 To MAX_BIT_CH
        If (m_aryBefore(i) <> m_aryAfter(i)) Then
            IsValid = True
            Exit For
        End If
    Next i
    
    IEeeAuto_TestConditionChecker_CheckAfterCondition = IsValid

End Function


Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    'モードはこのクラスにはない
    
End Sub

Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)

    '最終位置はこのクラスにはない

End Sub

Private Function GetApmuUb(ByRef aryUB() As Long) As Boolean

    Dim bitOnoff As Long
    Dim lBolardNum As Long
    lBolardNum = TheUB.AsAPMU.BoardNumber
    
    Dim i As Long
    For i = 0 To MAX_BIT_CH
        If (tl_APMU_GetUtilBitOne(lBolardNum, i, aryUB(i)) <> TL_SUCCESS) Then
            GetApmuUb = False
            Exit Function
        End If
    Next
    GetApmuUb = True
    
End Function

Private Function ChecktApmuUbModify() As Boolean

    On Error GoTo ErrorHandler
    
    'WorkSheetを取得
    Dim mySht As Object
    Set mySht = ThisWorkbook.Worksheets(TheUB.AsAPMU.UBSetShtName)
    
    '一番最初のApmuUBSetを取得
    Dim strCondition As String
    strCondition = mySht.Range(INDEX_CELL_TOP).Value
        
    'シートから設定値を取得
    Dim myCol As Collection
    Set myCol = mySht.UBCondition.Item(strCondition)
    
    '一番最初の条件設定を実行
    TheUB.AsAPMU.SetUBCondition strCondition
    
    'APMUUBの取得
    Call GetApmuUb(m_aryBefore)
    
    '比較
    Dim IsSame As Boolean
    IsSame = True
    Dim myItem As Variant
    For Each myItem In myCol
        If (m_aryBefore(myItem(0)) <> myItem(1)) Then
             IsSame = False
            Exit For
        End If
    Next
    
    ChecktApmuUbModify = IsSame
    
    If Not IsSame Then
        MsgBox THIS_CLASS_NAME & " is Disable"
    End If
    
    '最後はこの処理がなかったことにしておく
    TheUB.AsAPMU.ResetUtilityBits
    
    Exit Function
    
ErrorHandler:

End Function
