VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProcParamReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �e���[�_�[����������C���^�[�t�F�C�X
'
'�ړI:
'   �e�ړI���ƂɕK�v�ƂȂ郊�[�_�[����������
'
'�쐬��:
'   0145184306
'
Option Explicit

Public Function ReadAsBoolean(ByVal dataName As String) As Boolean
End Function

Public Function ReadAsDouble(ByVal dataName As String) As Double
End Function

Public Function ReadAsInteger(ByVal dataName As String) As Integer
End Function

Public Function ReadAsLong(ByVal dataName As String) As Long
End Function

'Public Function ReadAsPixel(ByVal pParamName As String) As CImgPixel
'End Function

Public Function ReadAsPlane(ByVal pParamName As String, ByRef pImgPlane As CImgPlane)
End Function

'Public Function ReadAsSite(ByVal pParamName As String) As CParamSite
'End Function

Public Function ReadAsString(ByVal dataName As String) As String
End Function

Public Property Get MyName() As String
End Property

Public Function ReadAsBlockEnd(ByVal pParamName As String) As IProcParamReader
End Function

Public Function ReadAsBlockStart(ByVal pParamName As String) As IProcParamReader
End Function