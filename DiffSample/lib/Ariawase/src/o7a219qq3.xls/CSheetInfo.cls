VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���[�N�V�[�g���
'
'�ړI:
'   �ėp���[�N�V�[�gReader�̃V�[�g�ˑ��̏������̃I�u�W�F�N�g�ɏW��
'   �ėp���[�N�V�[�gReader�Ɏ�������SheetInfo��؂�ւ��邱�ƂŃV�[�g����؂�ւ���
'
'�쐬��:
'   a_oshima
'
'�O�����:
'   �O���[�o���I�u�W�F�N�g TheError:CErrManager �����݂��邱��


Option Explicit

Private mKeys As Collection        '�񌩏o���̖���
Private mAddresses As Collection  '�񌩏o���̃A�h���X
Private mPrimaryKey As String      '��L�[

Property Let SetKeyAddress(ByVal pName As String, ByVal pIndexAddress As String)
'���e:
'   ���o���̖��̂ƃA�h���X��o�^
'
'�p�����[�^:
'   [pName]          In  String�^:  �񌩏o���̖���
'   [pIndexAddress]  In  String�^:  �񌩏o���̃A�h���X
'
    mKeys.Add pName, pName
    mAddresses.Add pIndexAddress, pName
End Property

Property Get GetAddress(ByVal pName As String) As String
'���e:
'   �񌩏o���̃A�h���X��Ԃ��܂�
'
'�p�����[�^:
'   [pName]          In  String�^:  �񌩏o���̖���
'
'�߂�l:
'    �񌩏o���̃A�h���X
'
'��O:
'   �w�肵���񌩏o����������Ȃ��Ƃ��͗�O9999�𓊂��܂�
'
    If IsExist(pName) Then
        GetAddress = mAddresses(pName)
    Else
        TheError.Raise 9999, TypeName(Me) & ".GetAddress", "Not found: " & pName
    End If
End Property
Property Let SetPrimaryKey(ByVal pName As String)
'���e:
'   ��L�[��o�^
'
'�p�����[�^:
'   [pName]          In  String�^:  �񌩏o���̖���
    mPrimaryKey = pName
End Property
Function IsExist(ByVal pName As String) As Boolean
'���e:
'   ���o�������݂��邩�ǂ�����Ԃ��܂�
'
'�p�����[�^:
'   [pName]          In  String�^:  �񌩏o���̖���
'
'�߂�l:
'    True:���݂���AFalse:���݂��Ȃ�
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
'���e:
'   �񌩏o���̈ꗗ��Ԃ��܂�
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'    �񌩏o�����̂�Collection(Collection�^)
    Set Keys = mKeys
End Property
Property Get GetPrimaryKey() As String
'���e:
'   ��L�[�̖��̂�Ԃ��܂�
'
'�p�����[�^:
'   �Ȃ�
'
'�߂�l:
'   ��L�[�̖���(String�^)
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
