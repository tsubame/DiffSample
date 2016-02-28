VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamLsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートのFixedLSB/DynamicLSB情報を保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  FixedLSB/DynamicLSB情報を保持する機能を実装した
'       2010/11/30  Clone/Initialize機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter用I/Fを宣言する

Private Const ERR_NUMBER As Long = 9999             ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamLsb"    ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty ' Default値"Empty"を保持する

Private Const PARAM_BASIC_LSB As String = "Basic Lsb@LSB Config"        ' パラメータ参照キーを保持する
Private Const PARAM_DYNAMIC_LSB As String = "Procedure Name@LSB Config" ' パラメータ参照キーを保持する
Private Const PARAM_NAME As String = "LSB"                              ' パラメータ名称を保持する

Private Const MAIN_UNIT As String = "V" ' 単位名称を保持する

Private mLsbSnapper As ILSBSnapper ' LSB取得用インターフェースを保持する

Private mMainUnit As String             ' 単位情報を保持する
Private mSubUnit As String              ' サブ単位情報を保持する
Private mValue As String                ' パラメータ値を保持する
Private mFlagDynamic As Boolean         ' パラメータ種別を保持する
Private mDataArray As CParamDoubleArray ' 配列型データを保持する

Public Sub Initialize(ByVal strData As String, Optional ByRef clsPrmDblAry As CParamDoubleArray = Nothing, Optional ByVal DynamicMode As Boolean = False)
'内容:
'   パラメータ情報を格納する
'
'パラメータ:
'   [strValue]      In  パラメータ値情報を保持する
'   [clsPrmDblAry]  In  配列Double型データを保持する
'   [DynamicMode]   In  パラメータ種別情報を保持する
'
'戻り値:
'
'注意事項:
'

    '#####  パラメータが単位付き数値の場合は、単位と数値に分解＆格納する  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strData, mMainUnit, mSubUnit, mValue)
    mFlagDynamic = DynamicMode
    If Not (clsPrmDblAry Is Nothing) Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetArrayAsDouble clsPrmDblAry.ReadAsDouble
    Else
        Set mDataArray = clsPrmDblAry
    End If
End Sub

Public Function Clone() As CParamLsb
'内容:
'   Cloneを作成する
'
'パラメータ:
'
'戻り値:
'   CParamLsb(Clone)
'
'注意事項:
'

    '#####  自身のCloneを作成して戻す  #####
    Dim cloneobj As CParamLsb
    Set cloneobj = New CParamLsb
    cloneobj.Initialize Me.Parameter, mDataArray, mFlagDynamic
    Set Clone = cloneobj
    Set cloneobj = Nothing

End Function

Public Sub ReadParameter(ByRef paramReader As IParameterReader, ByRef strParamName As String)
'内容:
'   パラメータ情報を取得＆格納する
'
'パラメータ:
'   [paramReader]   In  DeviceConfigurationsシートの情報を保持する
'   [strParamName]  In  取得するパラメータ参照キーの情報を保持する
'
'戻り値:
'
'注意事項:
'

    '#####  パラメータを単位と数値に分解して、取得する  #####
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strParamName, MAIN_UNIT, mMainUnit, mSubUnit, mValue)

    '#####  LSB種別によって、LSB値を取得するクラスを切り替える  #####
    If mValue <> "" Then
        Select Case strParamName
            Case PARAM_BASIC_LSB ' 固定値か変数名称の場合
                Dim clsFixedLSB As CFixedLsb
                Set clsFixedLSB = New CFixedLsb
                clsFixedLSB.Initialize mValue
                Set mLsbSnapper = clsFixedLSB
                Set clsFixedLSB = Nothing
                mFlagDynamic = False
            Case PARAM_DYNAMIC_LSB ' マクロ実行してLSB値を取得する場合
                Dim clsDynamicLsb As CDynamicLsb
                Set clsDynamicLsb = New CDynamicLsb
                clsDynamicLsb.Initialize mValue
                Set mLsbSnapper = clsDynamicLsb
                Set clsDynamicLsb = Nothing
                mFlagDynamic = True
        End Select
    End If

End Sub

Public Property Get Parameter() As String
'内容:
'   取得したパラメータ情報を戻す
'
'パラメータ:
'
'戻り値:
'   取得したパラメータ情報
'
'注意事項:
'

    '#####  数値の場合は、サブ単位を利用して、単位換算を実施する  #####
    Parameter = mValue
    If IsNumeric(mValue) = True Then
        Parameter = CStr(CDbl(mValue) / SubUnitToValue(mSubUnit)) & mSubUnit & mMainUnit
    End If

End Property

Public Sub CalcLsb(Optional ByVal dblGainRatio As Variant)
'内容:
'   LSB値を算出する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    '#####  数値データの有無を確認する  #####
    If mLsbSnapper Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CalcLsb", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If

    '#####  Staticの場合は、Gain倍率で割り戻す  #####
    '#####  Dynamicの場合は、何もしない         #####
    Dim dblArray() As Double
    If mFlagDynamic = False Then
        Dim i As Long
        Dim dblLsb() As Double
        ReDim dblArray(UBound(dblGainRatio))
        dblLsb = mLsbSnapper.Lsb.ReadAsDouble
        For i = 0 To UBound(dblGainRatio) Step 1
            If dblGainRatio(i) <> 0 Then
                dblArray(i) = dblLsb(i) / dblGainRatio(i)
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".CalcLsb", "Couldn't Calculated LSB.[GainRatio = 0]")
            End If
        Next i
    Else
        dblArray = mLsbSnapper.Lsb.ReadAsDouble
    End If
    
    '#####  データを格納する  #####
    Set mDataArray = New CParamDoubleArray
    mDataArray.SetArrayAsDouble dblArray

End Sub

Private Function calcUnit(ByVal dblValue As Double) As String

    '#####  サブ単位が必要か確認する  #####
    If dblValue = 0 Then
        calcUnit = ""
        Exit Function
    ElseIf dblValue <= 1 * SubUnitToValue("p") Then
        calcUnit = "p"
        Exit Function
    End If

    '#####  最適なサブ単位を求める  #####
    Dim strUnit As Variant
    Dim dblCalc As Double
    For Each strUnit In Array("m", "u", "n", "p")
        dblCalc = dblValue / SubUnitToValue(strUnit)
        If dblCalc > 1 Then
            calcUnit = strUnit
            Exit Function
        End If
    Next

End Function

Private Sub Class_Initialize()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    mFlagDynamic = False
    Set mDataArray = Nothing
    Set mLsbSnapper = Nothing
End Sub

Private Sub Class_Terminate()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    mFlagDynamic = False
    Set mDataArray = Nothing
    Set mLsbSnapper = Nothing
End Sub

Private Function IDeviceParameter_AsAccTimeH() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsAccTimeH", "UnSupported Data Type.[AccTimeH]")
End Function

Private Function IDeviceParameter_AsAccTimeV() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsAccTimeV", "UnSupported Data Type.[AccTimeV]")
End Function

Private Function IDeviceParameter_AsDecibel() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDecibel", "UnSupported Data Type.[Decibel]")
End Function

Private Function IDeviceParameter_AsDouble() As Double()

    '#####  数値データの有無を確認する  #####
    If mDataArray Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDouble", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If
    IDeviceParameter_AsDouble = mDataArray.ReadAsDouble

End Function

Private Function IDeviceParameter_AsRatio() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsRatio", "UnSupported Data Type.[Ratio]")
End Function

Private Function IDeviceParameter_AsString() As String()
    
    '#####  数値データの有無を確認する  #####
    If mDataArray Is Nothing Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsString", "Parameter is Nothing.[" & PARAM_NAME & "]")
    End If

    '#####  数値データを取得する  #####
    Dim dblValue() As Double
    dblValue = mDataArray.ReadAsDouble
    
    '#####  サブ単位を利用して、単位換算した数値を算出して、単位付き数値に変換する  #####
    Dim i As Integer
    Dim strValue() As String
    Dim strSubUnit As String
    ReDim strValue(UBound(dblValue))
    strSubUnit = mSubUnit ' サブ単位を格納する
    For i = 0 To UBound(dblValue) Step 1
        '#####  初期値が、変数名称やマクロ名称の場合は、サブ単位を求める  #####
        If IsAlphabet(mValue) = True Then
            strSubUnit = calcUnit(dblValue(i))
        End If
        '#####  サブ単位を利用して、単位換算した数値を算出して、単位付き数値に変換する  #####
        strValue(i) = CStr(Format((dblValue(i) / SubUnitToValue(strSubUnit)), "0.0###")) & strSubUnit & MAIN_UNIT
    Next i
    IDeviceParameter_AsString = strValue

End Function

Private Function IDeviceParameter_AsStringAccTimeH() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsStringAccTimeH", "UnSupported Data Type.[StringAccTimeH]")
End Function

Private Function IDeviceParameter_AsStringAccTimeV() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsStringAccTimeV", "UnSupported Data Type.[StringAccTimeV]")
End Function

Public Function AsIDeviceParameter() As IDeviceParameter
    Set AsIDeviceParameter = Me
End Function
