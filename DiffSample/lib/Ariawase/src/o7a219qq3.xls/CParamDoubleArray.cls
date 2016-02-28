VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamDoubleArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   Double型のデータを配列Double型で保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  Double型の情報を配列Double型で保持する機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Private Const ERR_NUMBER As Long = 9999                  ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamDoubleArray" ' Class名称を保持する

Private mDouble() As Double ' 配列で数値データを保持する

Public Sub SetArrayAsDouble(ByRef dblArray() As Double)
'内容:
'   配列Double型のデータを配列Double型のデータで格納する
'
'パラメータ:
'   [dblArray]  In  配列Double型データを保持する
'
'戻り値:
'
'注意事項:
'

    '#####  データを格納する  #####
    ReDim mDouble(UBound(dblArray))
    mDouble = dblArray

End Sub

Public Sub SetAsDouble(ByVal dblValue As Double)
'内容:
'   Double型のデータを配列Double型のデータで格納する
'
'パラメータ:
'   [dblValue]  In  Double型データを保持する
'
'戻り値:
'
'注意事項:
'

    '#####  Site数を求めて、配列Double型に展開＆格納する  #####
    Dim i As Integer
    ReDim mDouble(TheExec.sites.ExistingCount - 1)
    For i = 0 To (TheExec.sites.ExistingCount - 1) Step 1
        mDouble(i) = dblValue
    Next i

End Sub

Public Property Get ReadAsDouble() As Double()
'内容:
'   保持している配列Double型のデータを戻す
'
'パラメータ:
'
'戻り値:
'   保持している配列Double型のデータ
'
'注意事項:
'

    '#####  保持している配列Double型のデータを戻す  #####
    ReadAsDouble = mDouble

End Property

