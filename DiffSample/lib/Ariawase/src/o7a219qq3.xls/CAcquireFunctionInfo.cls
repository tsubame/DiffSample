VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAcquireFunctionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   Image ACQTBL シートの情報格納クラス
'目的:
'   Image Acquire Table シートにおける情報を格納し、Actionクラスに関連付け、ユーザーマクロからアクセス可能にする。
'   ユーザーマクロにこのインスタンスを渡すことで、ユーザーはProperty Getでシートの情報を取得可能。
'   イニシャル時のみ書き込み可能。その他は不可（シートの情報との不整合避けるため）
'作成者:
'  tomoyoshi.takase
'作成日: 2010年3月11日
'注意事項:
'   ImageACQTBL シートの Instance Name は必ずユニークでなければならない。他のInstance Nameとの重複はできません。
'
Option Explicit

Private Const ERR_NUMBER = 9999                         'エラー時に渡すエラー番号
Private Const CLASS_NAME = "CAcquireFunctionInfo"       'このクラスの名前

Private Const m_MAX_ARG_NO As Long = 79                 'ImageACQTBLシートの最大Arg数0~79

Private m_GroupName As String                           'ImageACQTBLシートのGroupName
Private m_MacroName As String                           'ImageACQTBLシートのMacroName
Private m_AutoAcquire As String                         'ImageACQTBLシートのAutoAcquire
Private m_InstanceName As String                        'ImageACQTBLシートのInstanceName
Private m_Arg(m_MAX_ARG_NO) As String                   'ImageACQTBLシートのArg群
Private m_ArgParameterMaxIndex As Integer               'ImageACQTBLシートのArgの最後の数
Private m_Comment As String                             'ImageACQTBLシートのComment

Property Get GroupName() As String
'内容:
'
'パラメータ:
'
'戻り値:
'   Group Name に対応するラベル名
'
'注意事項:
'
    GroupName = m_GroupName
End Property
Property Get macroName() As String
    macroName = m_MacroName
End Property
Property Get AutoAcquire() As String
    AutoAcquire = m_AutoAcquire
End Property
Property Get InstanceName() As String
    InstanceName = m_InstanceName
End Property
Property Get Arg(ByVal index As Long) As Variant
'内容:
'   ImageACQTBLシートのArg0-79の情報を返します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年3月15日
'パラメータ:
'   [index] In  1):Argに対応する数
'戻り値:
'   Variant
'
'注意事項:
'
    If index < 0 Or index > m_ArgParameterMaxIndex Then
        Arg = Empty
    Else
        Arg = m_Arg(index)
    End If
End Property
Property Get Comment() As String
    Comment = m_Comment
End Property

Property Get ArgParameterCount() As Long
    ArgParameterCount = m_ArgParameterMaxIndex + 1
End Property


Property Get ToCSVString() As String
'内容:
'   このインスタンスが持つ全ARG情報をCSVで返します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   なし
'戻り値:
'   String
'注意事項:
'
    Dim str As String
    Dim i  As Integer

    For i = 0 To m_ArgParameterMaxIndex
        str = str & m_Arg(i) & ","
    Next i
    
    ToCSVString = str

End Property

Public Function Clone() As CAcquireFunctionInfo
'内容:
'
'作成者:
'  tomoyoshi.takase
'作成日: 2010年11月25日
'パラメータ:
'   なし
'戻り値:
'   CAcquireFunctionInfo
'
'注意事項:
'

    Dim pClone As CAcquireFunctionInfo
    Set pClone = New CAcquireFunctionInfo
    Call pClone.CopyObj(Me)
    Set Clone = pClone

End Function

Public Sub CopyObj(OrgObj As CAcquireFunctionInfo)
'内容:
'   OrgObj の内容をこのインスタンスの内容に反映します。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [OrgObj]    In  1):内容を取得したい CAcquireFunctionInfo
'戻り値:
'
'注意事項:
'
    Dim i As Integer
    
    If m_GroupName = "" And m_MacroName = "" And m_AutoAcquire = "" Then
        With OrgObj
            m_GroupName = .GroupName                              'OrgObjのGroupName
            m_MacroName = .macroName                              'OrgObjのMacroName
            m_AutoAcquire = .AutoAcquire                          'OrgObjのAutoAcquire
            m_InstanceName = .InstanceName                        'OrgObjのInstanceName
            For i = 0 To m_MAX_ARG_NO
                If .Arg(i) <> "" Then
                    m_Arg(i) = .Arg(i)                             'OrgObjのArg群
                    m_ArgParameterMaxIndex = i
                End If
            Next i
            m_Comment = .Comment                                   'OrgObjのComment
        End With
    Else
        Call TheError.Raise(9999, TypeName(Me), "CAcquireFunctionInfo can write at once.")
    End If

End Sub

Public Sub Initialize(reader As IParameterReader)
'内容:
'   ImageACQTBLシートの情報を格納します。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年3月15日
'パラメータ:
'   [Reader]    In/Out  1):ImageACQTBLのパラメータリーダーインターフェース
'戻り値:
'
'注意事項:
'

    Dim i As Integer
    
    With reader
        m_GroupName = .ReadAsString("Group Name")                              'ImageACQTBLシートのGroupName
        m_MacroName = .ReadAsString("Macro Name")                              'ImageACQTBLシートのMacroName
        m_AutoAcquire = .ReadAsString("Auto Acquire")                          'ImageACQTBLシートのAutoAcquire
        m_InstanceName = .ReadAsString("Instance Name")                        'ImageACQTBLシートのInstanceName
        For i = 0 To m_MAX_ARG_NO
            If .ReadAsString("Arg" & i & "@Parameters") <> "" Then
                m_Arg(i) = .ReadAsString("Arg" & i & "@Parameters")                          'ImageACQTBLシートのArg群
                m_ArgParameterMaxIndex = i
            End If
        Next i
        m_Comment = .ReadAsString("Comment")                                 'ImageACQTBLシートのComment
    End With

End Sub

Private Sub Class_Initialize()
'内容:
'　このクラスのオブジェクト生成時の処理。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   なし
'戻り値:
'
'注意事項:

m_GroupName = ""                            'ImageACQTBLシートのGroupName
m_MacroName = ""                            'ImageACQTBLシートのMacroName
m_AutoAcquire = ""                          'ImageACQTBLシートのAutoAcquire
m_InstanceName = ""                         'ImageACQTBLシートのInstanceName
m_Comment = ""                              'ImageACQTBLシートのComment
Erase m_Arg                                 'ImageACQTBLシートのArg群
m_ArgParameterMaxIndex = 0                  'ImageACQTBLシートのArgの最後の数

End Sub

Private Sub Class_Terminate()
'内容:
'　このクラスのオブジェクト破棄時の処理。
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   なし
'戻り値:
'
'注意事項:

m_GroupName = ""                            'ImageACQTBLシートのGroupName
m_MacroName = ""                            'ImageACQTBLシートのMacroName
m_AutoAcquire = ""                          'ImageACQTBLシートのAutoAcquire
m_InstanceName = ""                         'ImageACQTBLシートのInstanceName
m_Comment = ""                              'ImageACQTBLシートのComment
Erase m_Arg                                 'ImageACQTBLシートのArg群
m_ArgParameterMaxIndex = 0                  'ImageACQTBLシートのArgの最後の数

End Sub


