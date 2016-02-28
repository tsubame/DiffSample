VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   ワークシート情報
'
'目的:
'   汎用ワークシートReaderのシート依存の情報をこのオブジェクトに集約
'   汎用ワークシートReaderに持たせるSheetInfoを切り替えることでシート情報を切り替える
'
'作成者:
'   a_oshima
'
'前提条件:
'   グローバルオブジェクト TheError:CErrManager が存在すること


Option Explicit

Private mKeys As Collection        '列見出しの名称
Private mAddresses As Collection  '列見出しのアドレス
Private mPrimaryKey As String      '主キー

Property Let SetKeyAddress(ByVal pName As String, ByVal pIndexAddress As String)
'内容:
'   見出しの名称とアドレスを登録
'
'パラメータ:
'   [pName]          In  String型:  列見出しの名称
'   [pIndexAddress]  In  String型:  列見出しのアドレス
'
    mKeys.Add pName, pName
    mAddresses.Add pIndexAddress, pName
End Property

Property Get GetAddress(ByVal pName As String) As String
'内容:
'   列見出しのアドレスを返します
'
'パラメータ:
'   [pName]          In  String型:  列見出しの名称
'
'戻り値:
'    列見出しのアドレス
'
'例外:
'   指定した列見出しが見つからないときは例外9999を投げます
'
    If IsExist(pName) Then
        GetAddress = mAddresses(pName)
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetAddress", "Not found: " & pName
    End If
End Property
Property Let SetPrimaryKey(ByVal pName As String)
'内容:
'   主キーを登録
'
'パラメータ:
'   [pName]          In  String型:  列見出しの名称
    mPrimaryKey = pName
End Property
Function IsExist(ByVal pName As String) As Boolean
'内容:
'   見出しが存在するかどうかを返します
'
'パラメータ:
'   [pName]          In  String型:  列見出しの名称
'
'戻り値:
'    True:存在する、False:存在しない
'
    Dim flg_found As Boolean
    Dim i As Long
    For i = 1 To mKeys.Count
        If mKeys.Item(i) = pName Then
            flg_found = True
            Exit For
        End If
    Next i
    IsExist = flg_found
End Function
Property Get Keys() As Collection
'内容:
'   列見出しの一覧を返します
'
'パラメータ:
'   なし
'
'戻り値:
'    列見出し名称のCollection(Collection型)
    Set Keys = mKeys
End Property
Property Get GetPrimaryKey() As String
'内容:
'   主キーの名称を返します
'
'パラメータ:
'   なし
'
'戻り値:
'   主キーの名称(String型)
    GetPrimaryKey = mPrimaryKey
End Property

Private Sub Class_Initialize()
    Set mKeys = New Collection
    Set mAddresses = New Collection
End Sub

Private Sub Class_Terminate()
    Set mKeys = Nothing
    Set mAddresses = Nothing
End Sub

