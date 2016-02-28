Attribute VB_Name = "XLibTheDeviceProfilerUtility"
'概要:
'   TheDeviceProfilerのユーティリティ
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceConfigのUtility機能を実装した
'       2010/11/30  DivideDeviceParameterを追加した
'       2010/12/07  不要コードを削除した
'
'作成者:
'   0145184346
'

Option Explicit

Public TheDeviceProfiler As CDeviceProfiler ' DeviceProfilerを宣言する

Private Const ERR_NUMBER = 9999                           ' Error番号を保持する
Private Const CLASS_NAME = "XLibTheDeviceProfilerUtility" ' Class名称を保持する
Private Const INITIAL_EMPTY_VALUE As String = Empty       ' Default値"Empty"を保持する

Private mLSBParam As Collection         ' LSB変数を保持する
Private mAccTimeParam As Collection     ' AccTime変数を保持する
Private mAccTimeParamUnit As Collection ' AccTime変数に対応する単位を保持する

Public Sub CreateTheDeviceProfilerIfNothing()
'内容:
'   TheDeviceProfilerを初期化する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheDeviceProfiler Is Nothing Then
        Set TheDeviceProfiler = New CDeviceProfiler
        Call TheDeviceProfiler.Initialize
    End If
    Exit Sub
ErrHandler:
    Set TheDeviceProfiler = Nothing
    Call TheError.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

Public Sub InitializeTheDeviceProfiler()
'内容:
'   TheDeviceProfilerを初期化する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Call TheDeviceProfiler.Initialize
End Sub

Public Sub DestroyTheDeviceProfiler()
'内容:
'   TheDeviceProfilerを破棄する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set TheDeviceProfiler = Nothing
    Set mLSBParam = Nothing
    Set mAccTimeParam = Nothing
    Set mAccTimeParamUnit = Nothing
End Sub

Public Function RunAtJobEnd() As Long
End Function

Public Sub SetLSBParam(ByVal strName As String, ByRef dblValue() As Double)
'内容:
'   変数化されたLSB値を格納する
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'

    '#####  変数値を格納する  #####
    If mLSBParam Is Nothing Then
        Set mLSBParam = New Collection
    End If
    On Error GoTo ErrHandler
    mLSBParam.Remove strName
    mLSBParam.Add dblValue, strName
    Exit Sub
ErrHandler:
    mLSBParam.Add dblValue, strName
    Exit Sub
End Sub

Public Function GetLSBParam(ByRef strName As String) As Double()
'内容:
'   変数化されたLSB値を取得する
'
'パラメータ:
'
'戻り値:
'   配列Double型(LSB値)
'
'注意事項:
'

    '#####  変数値を取得する  #####
    On Error GoTo ErrHandler
    GetLSBParam = mLSBParam.Item(strName)
    Exit Function
ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetLSBParam", "Nothing Parameter Name.[" & strName & "]")
    Exit Function
End Function

Public Sub SetAccTimeParam(ByVal strName As String, ByVal pramValue As Variant, Optional ByVal strUnit As String = "H")
'内容:
'   変数化されたAccTime値を格納する
'
'パラメータ:
'   [strName]    In  変数名称を保持する
'   [pramValue]  In  変数の数値を保持する
'   [strUnit]    In  変数値の単位名称を保持する
'
'戻り値:
'
'注意事項:
'

    '#####　単位名称を確認する  #####
    If (strUnit <> "H") And (strUnit <> "V") Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".SetAccTimeParam", "Illegal Unit. [""H""or""V""]")
    End If

    '#####  変数値と単位名称を格納する  #####
    If (mAccTimeParam Is Nothing) Or (mAccTimeParamUnit Is Nothing) Then
        Set mAccTimeParam = New Collection
        Set mAccTimeParamUnit = New Collection
    End If
    On Error GoTo ErrHandler
    mAccTimeParam.Remove strName
    mAccTimeParam.Add pramValue, strName
    mAccTimeParamUnit.Remove strName
    mAccTimeParamUnit.Add strUnit, strName
    Exit Sub
ErrHandler:
    mAccTimeParam.Add pramValue, strName
    mAccTimeParamUnit.Add strUnit, strName
    Exit Sub
End Sub

Public Sub GetAccTimeParam(ByRef strName As String, ByRef strUnit As String, ByRef dblValue() As Double)
'内容:
'   変数化されたAccTime値を取得する
'
'パラメータ:
'   [strName]   In  変数名称を保持する
'   [strUnit]   In  変数値の単位名称を保持する
'   [dblValue]  In  変数名称の数値を保持する
'
'戻り値:
'   CDeviceParamArray(クラス)
'
'注意事項:
'

    '#####  変数値と単位名称を取得する  #####
    On Error GoTo ErrHandler
    strUnit = mAccTimeParamUnit.Item(strName)
    dblValue = mAccTimeParam.Item(strName)
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetAccTimeParam", "Nothing Parameter Name.[" & strName & "]")
    Exit Sub

End Sub

Public Sub GetDeviceParameter(ByRef paramReader As IParameterReader, ByRef strParamName As String, ByRef strFixedMainUnit As String, _
                            ByRef retMainUnit As String, ByRef retSubUnit As String, ByRef retValue As String)
'内容:
'   DeviceConfigurationsシートのパラメータ値を分解＆格納する
'
'パラメータ:
'   [paramReader]       In  DeviceConfigurationsシートの情報を保持する
'   [strParamName]      In  取得するパラメータ名称を保持する
'   [strFixedMainUnit]  In  取得するパラメータの単位名称を保持する
'
'戻り値:
'   [retMainUnit]  Out  取得したパラメータの単位名称を戻す
'   [retSubUnit]   Out  取得したパラメータのサブ単位名称を戻す
'   [retValue]     Out  取得したパラメータの値を戻す
'
'注意事項:
'

    '#####  パラメータを取得して、文字列を確認する  #####
    Dim strData As String
    strData = paramReader.ReadAsString(strParamName)
    Call CheckAsString(strData)

    '#####  パラメータが単位付き数値の場合は、単位と数値に分解する  #####
    If IsNumeric(strData) = True Then
        Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetDeviceParameter", _
                        strParamName & " : [" & strData & "]  - This Is Not Appropriate Main Unit !")
    End If
    Call DivideDeviceParameter(strData, retMainUnit, retSubUnit, retValue)
    If (IsAlphabet(strData) = False) And (IsNumeric(strData) = False) Then
        '#####  想定している単位と設定した単位が不一致ならば、エラーとする  #####
        If (strFixedMainUnit <> retMainUnit) And (strFixedMainUnit <> "") Then
            Call TheError.Raise(ERR_NUMBER, CLASS_NAME & ".GetDeviceParameter", _
                            strParamName & " : [" & strData & "]  - This Is Not Appropriate Main Unit !")
        End If
    End If

End Sub

Public Sub DivideDeviceParameter(ByRef strData As String, ByRef retMainUnit As String, ByRef retSubUnit As String, ByRef retValue As String)
'内容:
'   DeviceConfigurationsシートのパラメータ値を分解＆格納する
'
'パラメータ:
'   [strData]  In  取得するパラメータ名称を保持する
'
'戻り値:
'   [retMainUnit]  Out  取得したパラメータの単位名称を戻す
'   [retSubUnit]   Out  取得したパラメータのサブ単位名称を戻す
'   [retValue]     Out  取得したパラメータの値を戻す
'
'注意事項:
'

    '#####  パラメータが単位付き数値の場合は、単位と数値に分解＆格納する  #####
    Dim strMainUnit As String
    Dim strSubUnit As String
    Dim dblSubValue As Double
    
    If (IsNumeric(strData) = True) Or (IsAlphabet(strData) = True) Or (strData = "") Then
        retMainUnit = INITIAL_EMPTY_VALUE
        retSubUnit = INITIAL_EMPTY_VALUE
        retValue = strData
    Else
        Call SplitUnitValue(CStr(strData), strMainUnit, strSubUnit, dblSubValue)
        retMainUnit = strMainUnit
        retSubUnit = strSubUnit
        retValue = CStr(dblSubValue * SubUnitToValue(strSubUnit))
    End If

End Sub
