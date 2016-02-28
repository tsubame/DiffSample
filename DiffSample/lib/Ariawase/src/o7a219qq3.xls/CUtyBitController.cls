VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUtyBitController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ユーティリティBitの設定機能の提供
'
'目的:
'   APMU/CUBユーティリティBit設定機能の提供
'
'作成者:
'   SLSI今手
'
'注意事項:
'   APMU_UB、CUB_UBに対応します。
'   CUB UBの設定機能利用時には､条件付コンパイル引数に
'   CUB_UB_USE = 1の記述が必要です。
'
'   初期化後、使用開始前に、APMU、CUBのユーティリティBit設定に
'   使用する､Utility Bit Settingワークシートの指定が必要です
'
'Code Checked
'
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用

Private m_ApmuUbSet As CApmuUBSet

#If CUB_UB_USE <> 0 Then
Private m_CubUbSet As CCubUBSet
#End If

'#Pass-Ok
Public Function AsAPMU() As CApmuUBSet
'内容:
'   APMUのUB設定機能OBJを選択
'
'パラメータ:
'
'戻り値:
'   APMUのUB設定機能Object
'
'注意事項:
'
    
    Set AsAPMU = m_ApmuUbSet

End Function

'#Pass-Ok
#If CUB_UB_USE <> 0 Then
Public Function AsCUB() As CCubUBSet
'内容:
'   CUBのUB設定機能OBJを選択
'
'パラメータ:
'
'戻り値:
'   CUBのUB設定機能Object
'
'注意事項:
'   利用時には、条件付コンパイル引数に
'   CUB_UB_USE = 1の記述が必要です。
'

    Set AsCUB = m_CubUbSet

End Function
#End If

'#Pass-Ok
Private Sub Class_Initialize()
    
    Me.Initialize

End Sub

'#Pass-Ok
Public Sub Initialize()
'内容:
'   UBコントローラの初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Call InitApmuUb
    Call InitCubUb

End Sub

'#Pass-Ok
Private Sub InitApmuUb()
'APMU UB設定Objectの初期化
    
    Set m_ApmuUbSet = Nothing
    Set m_ApmuUbSet = New CApmuUBSet

End Sub

'#Pass-Ok
Private Sub InitCubUb()
'CUB UB設定Objectの初期化

    #If CUB_UB_USE <> 0 Then
    Set m_CubUbSet = Nothing
    Set m_CubUbSet = New CCubUBSet
    Set m_CubUbSet.CubUBObj = CUBUtilBit
    #End If

End Sub
