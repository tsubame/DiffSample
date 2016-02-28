VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgLUTInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   LUTの情報保持用
'目的:
'   登録したLUTの情報の問い合わせ用。
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'注意事項:
'   LUTの情報はループで細かく設定する(３万件程度)ことがあるため、
'   このクラスがメモリを大量に消費する可能性があります。
'   IMX083の定義例で確認した場合、メモリ消費は１Ｍ以下でした。
'参考
'   IG-XLのLUT登録仕様
'   ・InputValue1がInputValue2より大きいとエラー。
'   ・同一アドレスへの再定義は最後の定義が有効
'   ・同じ名前でBit指定の変更は問題なし。変更した場合、最後のBit指定が有効。
'　　　(コード指定の場合。LUTシートの場合は不可能。Eee-JOBはエラーとする)
'   ・OutputValue 換算値の小数点以下は切り捨て。
'   　(※Fix関数相当。Int関数ではない。)
'   ・同一アドレス（InputValue1とInputValue2が同じ）に違う出力値を設定した場合、OutputValue1が有効。

Option Explicit

Private Const INPUT1 As Integer = 0
Private Const INPUT2 As Integer = 1
Private Const OUTPUT1 As String = 2
Private Const OUTPUT2 As String = 3

Private mName As String                 'このLUT定義の名前
Private mValidBit As Integer            'このLUT定義の有効bit数
Private mOutputValue() As Long          'LUT変換値の保持用。配列番号がInputValueに相当。
                                        '配列数はCreateIndexLUTで最後に指定したbit数で決定する。対応数からあふれた情報は無視。(IG-XLに合わせた)
Private mLutDefinitions As Collection   '定義情報のCollection

Public Sub Initialize(ByVal pName As String, pBit As Integer)
'内容:
'   LUTの情報を登録します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月1日
'パラメータ:
'   [pName]  In  1):LUT定義名
'   [pBit]   In  2):登録されたbit数
'戻り値:
'
'注意事項:

    mName = pName
    mValidBit = pBit

End Sub

Public Sub Add(pName As String, pInputVal1 As Long, pInputVal2 As Long, pOutputVal1 As Long, pOutputVal2 As Long)
'内容:
'   LUT定義情報をこのクラスのCollectionに追加します。
'作成者:
'  tomoyoshi.takase
'作成日: 2011年2月3日
'パラメータ:
'   [pName]         In  1):LUT定義名。チェックするため。
'   [pInputVal1]    In  2):
'   [pInputVal2]    In  3):
'   [pOutputVal1]   In  4):
'   [pOutputVal2]   In  5):
'戻り値:
'
'注意事項:
'

    Dim pArrayVal(3) As Long    'LUT定義格納用。ユーザー定義変数はCollectionできないので配列使ってる。
    
    pArrayVal(INPUT1) = pInputVal1
    pArrayVal(INPUT2) = pInputVal2
    pArrayVal(OUTPUT1) = pOutputVal1
    pArrayVal(OUTPUT2) = pOutputVal2

    Call mLutDefinitions.Add(pArrayVal)

End Sub

Public Property Get OutputValue(ByVal pInputValue As Long) As Long
'内容:
'   InputValueに対応するOutputValueを返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月1日
'パラメータ:
'   [pInputValue]    In/Out  1):
'戻り値:
'   Long
'
'注意事項:
'

    If mLutDefinitions.Count <> 0 Then
        Call ExpandData
    End If
    
    On Error GoTo NOT_DEFINITION
    OutputValue = mOutputValue(pInputValue)
    Exit Property
NOT_DEFINITION:
    Call TheError.Raise(9999, "", "This InputValue is not found at LUT." & vbCrLf & "LUT Name:" & mName & vbCrLf & "Error InputValue:" & pInputValue)

End Property

Public Function AllValue() As Long()
'内容:
'   全変換情報を動的配列で返します。
'   配列番号が入力値、対応値が出力値に相当します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月1日
'パラメータ:
'
'戻り値:
'   Long
'
'注意事項:
'
    If mLutDefinitions.Count <> 0 Then
        Call ExpandData
    End If
    
    AllValue = mOutputValue

End Function

Public Property Get Bit() As Integer
'内容:
'   このLUTの有効Bit数を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年12月1日
'パラメータ:
'   なし
'戻り値:
'   Integer
'
'注意事項:

    Bit = mValidBit
End Property

Private Sub ExpandData()
'内容:
'   登録したLUT定義データを、配列情報に展開します。

Dim x1 As Long, y1 As Long      '２点の座標の始点
Dim x2 As Long, y2 As Long      '２点の座標の終点
Dim A As Double                 '２点を通る直線の傾き
Dim B As Double                 '２点を通る直線の切片
Dim pTmp  As Double             '変換結果の格納用
Dim pValue  As Variant          'Value群の一時保持用

Dim pLimitData As Long          'LUTデータの範囲

Dim i As Long

    '#### ユーザー定義最後のbit指定からデータ範囲を決定 ####
    pLimitData = 2 ^ mValidBit / 2                    'プラスマイナス範囲なので÷２
    If UBound(mOutputValue) = 0 Then                    'このメソッドが動いたことがない。
        ReDim mOutputValue(-pLimitData To pLimitData - 1)
    End If
    '#### ユーザー定義データを配列に展開 ####
    For Each pValue In mLutDefinitions
        '始点座標
        x1 = pValue(INPUT1)
        y1 = pValue(OUTPUT1)
        '終点座標
        x2 = pValue(INPUT2)
        y2 = pValue(OUTPUT2)
        
        '#### ２点を通る直線の算出 ####
        
        If x2 - x1 = 0 Then                 'ゼロ割対策
            mOutputValue(x1) = y1           'y1とy2が違った場合、y1が有効
            
        ElseIf x2 - x1 = 1 Then             'For文1Step刻みで定義することが多そうなので、計算無しで対応
            mOutputValue(x1) = y1
            mOutputValue(x2) = y2
        Else
            '２点を通る方程式(y=ax+b)を算出して、Outputデータを配列に展開
            A = (y1 - y2) / (x1 - x2)               '傾き
            B = (x1 * y2 - x2 * y1) / (x1 - x2)     '切片
            
            For i = x1 To x2
                pTmp = A * i + B
                '小数点以下切捨て(fix関数相当。Int関数とは違う。IG-XL仕様より。)
                pTmp = Fix(pTmp)
                
                mOutputValue(i) = pTmp
            Next i
        End If
        
    Next
    
    Set mLutDefinitions = Nothing
    Set mLutDefinitions = New Collection

End Sub

Private Sub Class_Initialize()
'内容:
'   コンストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    ReDim mOutputValue(0)
    Set mLutDefinitions = New Collection

End Sub

Private Sub Class_Terminate()
'内容:
'   デストラクタ
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月30日
'パラメータ:
'   なし
'戻り値:
'
'注意事項:
'

    Set mLutDefinitions = Nothing

End Sub
