VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgKernelInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   カーネルの情報
'
'目的:
'   カーネルの情報の取りまとめ
'
'作成者:
'   0145184004
'
Option Explicit

Private m_Name As String
Private m_Type As IdpKernelType
Private m_Width As Long
Private m_Height As Long
Private m_ShiftR As Long
Private m_Value As String
Private m_Comment As String


Friend Sub Create(ByVal pName As String, ByVal pType As IdpKernelType, ByVal pWidth As Long, ByVal pHeight As Long, ByVal pShiftR As Long, ByVal pValue As String, Optional ByVal pComment As String = "")

    m_Name = pName
    m_Type = pType
    m_Width = pWidth
    m_Height = pHeight
    m_ShiftR = pShiftR
    m_Value = pValue
    m_Comment = pComment

End Sub

Public Property Get Name() As String
'内容:
'   カーネルの名前を取得
'
'備考:
'
    Name = m_Name
End Property

Public Property Get KernelType() As IdpKernelType
'内容:
'   カーネルのタイプを取得
'
'備考:
'
    KernelType = m_Type
End Property

Public Property Get width() As Long
'内容:
'   カーネルの幅を取得
'
'備考:
'
    width = m_Width
End Property

Public Property Get height() As Long
'内容:
'   カーネルの高さを取得
'
'備考:
'
    height = m_Height
End Property

Public Property Get ShiftR() As Long
'内容:
'   カーネルのシフト量を取得
'
'備考:
'
    ShiftR = m_ShiftR
End Property

Public Property Get Value() As String
'内容:
'   カーネルの値を取得
'
'備考:
'
    Value = m_Value
End Property

Public Property Get Comment() As String
'内容:
'   カーネルのコメントを取得
'
'備考:
'
    Comment = m_Comment
End Property

