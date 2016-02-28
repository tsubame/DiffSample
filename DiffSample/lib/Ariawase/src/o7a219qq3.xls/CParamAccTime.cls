VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParamAccTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceConfigurationsシートのAccTime情報を保持するクラス
'
'   Revision History:
'       Data        Description
'       2010/11/19  AccTime情報を保持する機能を実装した
'       2010/11/30  Clone/Initialize機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Implements IDeviceParameter ' DeviceParameter用I/Fを宣言する

Private Const ERR_NUMBER As Long = 9999              ' Error番号を保持する
Private Const CLASS_NAME As String = "CParamAccTime" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty  ' Default値"Empty"を保持する

Private Const MAIN_UNIT_V As String = "V"               ' 単位名称を保持する
Private Const MAIN_UNIT_H As String = "H"               ' 単位名称を保持する
Private Const PARAM_NAME_ACCTIME As String = "Acc Time" ' パラメータ名称を保持する
Private Const PARAM_NAME_XHSpV As String = "XHS/V"      ' パラメータ名称を保持する

Private mMainUnit_XHSpV As String                ' 単位情報を保持する
Private mSubUnit_XHSpV As String                 ' サブ単位情報を保持する
Private mValue_XHSpV As String                   ' パラメータ値を保持する
Private mMainUnit_AccTime As String              ' 単位情報を保持する
Private mSubUnit_AccTime As String               ' サブ単位情報を保持する
Private mValue_AccTime As String                 ' パラメータ値を保持する
Private mDataArray_AccTimeH As CParamDoubleArray ' 配列型データを保持する
Private mDataArray_AccTimeV As CParamDoubleArray ' 配列型データを保持する
Private mArrange_Unit As String                  ' 計算直前に取得したAccTimeの単位
Private mArrange_Value() As Double               ' 計算直前に取得したAccTimeの値

Public Sub Initialize(ByVal strXHSpV As String, ByVal strAccTime As String, ByRef dblParam() As Double, ByVal strPrmUnit As String)
'内容:
'   パラメータ情報を格納する
'
'パラメータ:
'   [strXHSpV]    In  XHS/Vパラメータ値情報を保持する
'   [strAccTime]  In  AccTimeパラメータ値情報を保持する
'   [dblParam]    In  算出前パラメータの配列Double型データを保持する
'   [strPrmUnit]  In  算出前パラメータの単位名称を保持する
'
'戻り値:
'
'注意事項:
'

    '#####  パラメータが単位付き数値の場合は、単位と数値に分解＆格納する  #####
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strXHSpV, mMainUnit_XHSpV, mSubUnit_XHSpV, mValue_XHSpV)
    Call XLibTheDeviceProfilerUtility.DivideDeviceParameter(strAccTime, mMainUnit_AccTime, mSubUnit_AccTime, mValue_AccTime)
    mArrange_Value = dblParam
    mArrange_Unit = strPrmUnit

End Sub

Public Function Clone() As CParamAccTime
'内容:
'   Cloneを作成する
'
'パラメータ:
'
'戻り値:
'   CParamAccTime(Clone)
'
'注意事項:
'

    '#####  自身のCloneを作成して戻す  #####
    Dim cloneobj As CParamAccTime
    Set cloneobj = New CParamAccTime
    cloneobj.Initialize Me.ParameterXHSpV, Me.ParameterAccTime, mArrange_Value, mArrange_Unit
    Set Clone = cloneobj
    Set cloneobj = Nothing

End Function

Public Sub ReadParameter(ByRef paramReader As IParameterReader, ByRef strXHSpVName As String, ByRef strAccTimeName As String)
'内容:
'   パラメータ情報を取得＆格納する
'
'パラメータ:
'   [paramReader]     In  DeviceConfigurationsシートの情報を保持する
'   [strXHSpVName]    In  取得するパラメータ参照キーの情報を保持する
'   [strAccTimeName]  In  取得するパラメータ参照キーの情報を保持する
'
'戻り値:
'
'注意事項:
'

    '#####  パラメータを単位と数値に分解して、取得する  #####
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strXHSpVName, MAIN_UNIT_H, mMainUnit_XHSpV, mSubUnit_XHSpV, mValue_XHSpV)
    Call XLibTheDeviceProfilerUtility.GetDeviceParameter(paramReader, strAccTimeName, "", mMainUnit_AccTime, mSubUnit_AccTime, mValue_AccTime)

    '#####  単位名称(V/H)を確認する  #####
    If mValue_AccTime <> "" Then
        If IsNumeric(mValue_AccTime) = True Then
            If (mMainUnit_AccTime <> MAIN_UNIT_V) And (mMainUnit_AccTime <> MAIN_UNIT_H) Then
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                                strAccTimeName & " : [" & paramReader.ReadAsString(strAccTimeName) & "]  - This Is Not Appropriate Main Unit !")
            End If
        End If
    End If

    '#####  サブ単位設定を許可しないので、エラーチェックを実施する  #####
    If mSubUnit_XHSpV <> "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                        strXHSpVName & " : [" & paramReader.ReadAsString(strXHSpVName) & "]  - This Is Not Appropriate Sub Unit !")
    ElseIf mSubUnit_AccTime <> "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetupParameter", _
                        strAccTimeName & " : [" & paramReader.ReadAsString(strAccTimeName) & "]  - This Is Not Appropriate Sub Unit !")
    End If

End Sub

Public Property Get ParameterXHSpV() As String
'内容:
'   保持していパラメータ情報(XHS/V)を戻す
'
'パラメータ:
'
'戻り値:
'   保持していパラメータ情報
'
'注意事項:
'

    '#####  単位を付加する  #####
    ParameterXHSpV = mValue_XHSpV & mMainUnit_XHSpV

End Property

Public Property Get ParameterAccTime() As String
'内容:
'   保持しているパラメータ情報(AccTime)を戻す
'
'パラメータ:
'
'戻り値:
'   保持しているパラメータ情報
'
'注意事項:
'

    '#####  単位を付加する  #####
    ParameterAccTime = mValue_AccTime & mMainUnit_AccTime

End Property

Private Sub paramAccTimeIsEmpty()
    If mValue_AccTime = "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramAccTimeIsEmpty", "Parameter is Nothing.[" & PARAM_NAME_ACCTIME & "]")
    End If
End Sub

Private Sub paramXHSpVIsEmpty()
    If mValue_XHSpV = "" Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramXHSpVIsEmpty", "Parameter is Nothing.[" & PARAM_NAME_XHSpV & "]")
    End If
End Sub

Private Sub paramXHSpVIsNumeric()
    If IsNumeric(mValue_XHSpV) = False Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".paramXHSpVIsNumeric", "Unsupported Variable.[" & mValue_XHSpV & "]")
    End If
End Sub

Public Sub ArrangeParameter()
'内容:
'   AccTime算出するために、変数名称から数値＆単位を取得する
'   数値の場合は、代入のみ実施する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    '#####  変数名称の場合は、数値＆単位を取得する  #####
    If IsAlphabet(mValue_AccTime) = True Then
        Call XLibTheDeviceProfilerUtility.GetAccTimeParam(mValue_AccTime, mArrange_Unit, mArrange_Value) ' 変数値と単位を取得する
    ElseIf IsNumeric(mValue_AccTime) = True Then
        Dim i As Integer
        ReDim mArrange_Value(TheExec.sites.ExistingCount - 1)
        For i = 0 To (TheExec.sites.ExistingCount - 1) Step 1
            mArrange_Value(i) = CDbl(mValue_AccTime)
        Next i
        mArrange_Unit = mMainUnit_AccTime
    End If

End Sub

Private Sub calcAccTimeByUnit(ByRef strCalcType As String)
'内容:
'   取得したパラメータを利用して、指定した単位でのAccTimeを算出する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    '#####  データが設定されていないため、エラーとする  #####
    Call paramAccTimeIsEmpty

    '#####  単位が異なる場合は、換算が必要となる              #####
    '#####  変数名称の場合は、格納先から値を取得して換算する  #####
    Dim i As Long
    Dim dblCalcValue() As Double
    ReDim dblCalcValue(UBound(mArrange_Value))
    If (strCalcType = MAIN_UNIT_H) And (mArrange_Unit = MAIN_UNIT_V) Then ' "H"基準で算出＆単位が"V"の場合に算出する
        Call paramXHSpVIsEmpty
        Call paramXHSpVIsNumeric
        For i = 0 To UBound(mArrange_Value) Step 1
            dblCalcValue(i) = mArrange_Value(i) * CDbl(mValue_XHSpV)
        Next i
    ElseIf (strCalcType = MAIN_UNIT_V) And (mArrange_Unit = MAIN_UNIT_H) Then ' "V"基準で算出＆単位が"H"の場合に算出する
        Call paramXHSpVIsEmpty
        Call paramXHSpVIsNumeric
        For i = 0 To UBound(mArrange_Value) Step 1
            If CDbl(mValue_XHSpV) <> 0 Then
                dblCalcValue(i) = mArrange_Value(i) / CDbl(mValue_XHSpV)
            Else
                Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".calcAccTimeByUnit", "Couldn't Calculated AccTime.[XHS/V = 0]")
            End If
        Next i
    Else
        dblCalcValue = mArrange_Value ' 単位が一致している場合は、何もしない
    End If
    
    
    '#####  算出したデータを格納する  #####
    Dim clsDataArray As CParamDoubleArray
    Set clsDataArray = New CParamDoubleArray
    clsDataArray.SetArrayAsDouble dblCalcValue

    '#####  算出したデータをモジュール変数に格納する  #####
    Select Case strCalcType
        Case MAIN_UNIT_H
            Set mDataArray_AccTimeH = clsDataArray
        Case MAIN_UNIT_V
            Set mDataArray_AccTimeV = clsDataArray
    End Select
    Set clsDataArray = Nothing

End Sub

Private Sub Class_Initialize()
    mMainUnit_XHSpV = INITIAL_EMPTY_VALUE
    mSubUnit_XHSpV = INITIAL_EMPTY_VALUE
    mValue_XHSpV = INITIAL_EMPTY_VALUE
    mMainUnit_AccTime = INITIAL_EMPTY_VALUE
    mSubUnit_AccTime = INITIAL_EMPTY_VALUE
    mValue_AccTime = INITIAL_EMPTY_VALUE
    Set mDataArray_AccTimeH = Nothing
    Set mDataArray_AccTimeV = Nothing
    mArrange_Unit = INITIAL_EMPTY_VALUE
End Sub

Private Sub Class_Terminate()
    mMainUnit_XHSpV = INITIAL_EMPTY_VALUE
    mSubUnit_XHSpV = INITIAL_EMPTY_VALUE
    mValue_XHSpV = INITIAL_EMPTY_VALUE
    mMainUnit_AccTime = INITIAL_EMPTY_VALUE
    mSubUnit_AccTime = INITIAL_EMPTY_VALUE
    mValue_AccTime = INITIAL_EMPTY_VALUE
    Set mDataArray_AccTimeH = Nothing
    Set mDataArray_AccTimeV = Nothing
    mArrange_Unit = INITIAL_EMPTY_VALUE
End Sub

Private Function IDeviceParameter_AsAccTimeH() As Double()

    '#####  数値データの有無を確認する  #####
    If mDataArray_AccTimeH Is Nothing Then
        Call calcAccTimeByUnit("H")
    End If
    IDeviceParameter_AsAccTimeH = mDataArray_AccTimeH.ReadAsDouble

End Function

Private Function IDeviceParameter_AsAccTimeV() As Double()

    '#####  数値データの有無を確認する  #####
    If mDataArray_AccTimeV Is Nothing Then
        Call calcAccTimeByUnit("V")
    End If
    IDeviceParameter_AsAccTimeV = mDataArray_AccTimeV.ReadAsDouble

End Function

Private Function IDeviceParameter_AsDecibel() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDecibel", "UnSupported Data Type.[Decibel]")
End Function

Private Function IDeviceParameter_AsDouble() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsDouble", "UnSupported Data Type.[Double]")
End Function

Private Function IDeviceParameter_AsRatio() As Double()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsRatio", "UnSupported Data Type.[Ratio]")
End Function

Private Function IDeviceParameter_AsString() As String()
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".AsString", "UnSupported Data Type.[String]")
End Function

Private Function IDeviceParameter_AsStringAccTimeH() As String()

    '#####  数値データの有無を確認する  #####
    If mDataArray_AccTimeH Is Nothing Then
        Call calcAccTimeByUnit("H")
    End If

    '#####  数値データを取得する  #####
    Dim dblValue() As Double
    dblValue = mDataArray_AccTimeH.ReadAsDouble

    '#####  サブ単位を利用して、単位換算した数値を算出して、単位付き数値に変換する  #####
    Dim i As Integer
    Dim strValue() As String
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        '#####  小数点以下の桁が存在する数値は、小数点以下4桁とする  #####
        If dblValue(i) Like "*.*" Then
            strValue(i) = CStr(Format(dblValue(i), "0.0###")) & MAIN_UNIT_H
        Else
            strValue(i) = CStr(dblValue(i)) & MAIN_UNIT_H
        End If
    Next i
    IDeviceParameter_AsStringAccTimeH = strValue

End Function

Private Function IDeviceParameter_AsStringAccTimeV() As String()

    '#####  数値データの有無を確認する  #####
    If mDataArray_AccTimeV Is Nothing Then
        Call calcAccTimeByUnit("V")
    End If

    '#####  数値データを取得する  #####
    Dim dblValue() As Double
    dblValue = mDataArray_AccTimeV.ReadAsDouble

    '#####  サブ単位を利用して、単位換算した数値を算出して、単位付き数値に変換する  #####
    Dim i As Integer
    Dim strValue() As String
    ReDim strValue(UBound(dblValue))
    For i = 0 To UBound(dblValue) Step 1
        '#####  小数点以下の桁が存在する数値は、小数点以下4桁とする  #####
        If dblValue(i) Like "*.*" Then
            strValue(i) = CStr(Format(dblValue(i), "0.0###")) & MAIN_UNIT_V
        Else
            strValue(i) = CStr(dblValue(i)) & MAIN_UNIT_V
        End If
    Next i
    IDeviceParameter_AsStringAccTimeV = strValue

End Function

Public Function AsIDeviceParameter() As IDeviceParameter
    Set AsIDeviceParameter = Me
End Function
