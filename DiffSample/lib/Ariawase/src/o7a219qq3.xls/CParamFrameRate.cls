VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamFrameRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートのFrameRate情報を保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  FrameRate情報を保持する機能を実装した
'       2010/11/30  Clone/Initialize機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter用I/Fを宣言する

Private Const ERR_NUMBER As Long = 9999                ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamFrameRate" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty    ' Default値"Empty"を保持する

Private Const MAIN_UNIT As String = "fps"         ' 単位名称を保持する
Private Const PARAM_NAME As String = "Frame Rate" ' パラメータ名称を保持する

Private mMainUnit As String             ' 単位情報を保持する
Private mSubUnit As String              ' サブ単位情報を保持する
Private mValue As String                ' パラメータ値を保持する
Private mDataArray As CParamDoubleArray ' 配列型データを保持する

Public Sub Initialize(ByVal strData As String, Optional ByRef clsPrmDblAry As CParamDoubleArray = Nothing)
'内容:
'   パラメータ情報を格納する
'
'パラメータ:
'   [strValue]      In  パラメータ値情報を保持する
'   [clsPrmDblAry]  In  配列Double型データを保持する
'
'戻り値:
'
'注意事項:
'

    '#####  パラメータが単位付き数値の場合は、単位と数値に分解＆格納する  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strData, mMainUnit, mSubUnit, mValue)
    If Not (clsPrmDblAry Is Nothing) Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetArrayAsDouble clsPrmDblAry.ReadAsDouble
    Else
        Set mDataArray = clsPrmDblAry
    End If
End Sub

Public Function Clone() As CParamFrameRate
'内容:
'   Cloneを作成する
'
'パラメータ:
'
'戻り値:
'   CParamFrameRate(Clone)
'
'注意事項:
'

    '#####  自身のCloneを作成して戻す  #####
    Dim cloneobj As CParamFrameRate
    Set cloneobj = New CParamFrameRate
    cloneobj.Initialize Me.Parameter, mDataArray
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

    '#####  数値をSite展開して、格納する  #####
    If mValue <> "" Then
        Set mDataArray = New CParamDoubleArray
        mDataArray.SetAsDouble mValue
    End If

End Sub

Public Property Get Parameter() As String
'内容:
'   保持しているパラメータ情報を戻す
'
'パラメータ:
'
'戻り値:
'   保持しているパラメータ情報
'
'注意事項:
'

    '#####  サブ単位を利用して、単位換算を実施する  #####
    Parameter = mValue
    If IsNumeric(mValue) = True Then
        Parameter = CStr(CDbl(mValue) / SubUnitToValue(mSubUnit)) & mSubUnit & mMainUnit
    End If

End Property

Private Sub Class_Initialize()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    Set mDataArray = Nothing
End Sub

Private Sub Class_Terminate()
    mMainUnit = INITIAL_EMPTY_VALUE
    mSubUnit = INITIAL_EMPTY_VALUE
    mValue = INITIAL_EMPTY_VALUE
    Set mDataArray = Nothing
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
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        strValue(i) = CStr((dblValue(i) / SubUnitToValue(mSubUnit))) & mSubUnit & mMainUnit
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
