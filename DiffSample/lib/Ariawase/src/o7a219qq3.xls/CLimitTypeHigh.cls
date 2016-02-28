VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLimitTypeHigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   スライスレベルのInclude処理をする
'
'目的:
'   Include処理のため、引数を+1して返す
'
'作成者:
'   0145184306
'
Option Explicit

Implements ILimitTypeOperator

Private Const ERR_NUMBER As Long = 9999                  ' Error番号を保持する
Private Const CLASS_NAME As String = "CLimitTypeHigh"       ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty      ' Default値"Empty"を保持する

Public Function AsIlimitTypeOperator() As ILimitTypeOperator
    Set AsIlimitTypeOperator = Me
End Function

Public Function IlimitTypeOperator_Include(ByVal pLevel As Long) As Long
'内容:
'   Include処理をする
'
'戻り値:
'   pLevel + 1
'
'備考:
'
    IlimitTypeOperator_Include = pLevel + 1
End Function


