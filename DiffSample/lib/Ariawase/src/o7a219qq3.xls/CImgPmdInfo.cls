VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImgPmdInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   PMDの情報を纏めるクラス。
'
'目的:
'   PMDの情報を保存。
'
'作成者:
'   0145184004
'
Option Explicit

Dim m_Name As String
Dim m_X As Long
Dim m_Y As Long
Dim m_Width As Long
Dim m_Height As Long


Friend Sub Create(ByVal pName As String, ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)
'内容:
'   データ作成
'
'[pName]       IN   String型:       PMD名
'[pX]          IN   Long型:         基点Xアドレス
'[pY]          IN   Long型:         基点Yアドレス
'[pWidth]      IN   Long型:         幅
'[pHeight]     IN   Long型:         高さ
'
'備考:
'
    m_Name = pName
    m_X = pX
    m_Y = pY
    m_Width = pWidth
    m_Height = pHeight
End Sub

Public Function Clone() As CImgPmdInfo
'内容:
'   自分自身と同じデータのオブジェクトを返す
'
'備考:
'   内部的な使用にしたい。
'
    Set Clone = New CImgPmdInfo
    Call Clone.Create(m_Name, m_X, m_Y, m_Width, m_Height)
End Function

Public Property Get Name() As String
'内容:
'   名前の取得
'
'備考:
'
    Name = m_Name
End Property

Public Property Get XAdr() As Long
'内容:
'   X原点の取得
'
'備考:
'
    XAdr = m_X
End Property

Public Property Let XAdr(ByVal RHS As Long)
'内容:
'   X原点の設定
'
'備考:
'
    m_X = RHS
End Property

Public Property Get YAdr() As Long
'内容:
'   Y原点の取得
'
'備考:
'
    YAdr = m_Y
End Property

Public Property Let YAdr(ByVal RHS As Long)
'内容:
'   Y原点の設定
'
'備考:
'
    m_Y = RHS
End Property

Public Property Get width() As Long
'内容:
'   幅の取得
'
'備考:
'
    width = m_Width
End Property

Public Property Let width(ByVal RHS As Long)
'内容:
'   幅の設定
'
'備考:
'
    m_Width = RHS
End Property

Public Property Get height() As Long
'内容:
'   高さの取得
'
'備考:
'
    height = m_Height
End Property

Public Property Let height(ByVal RHS As Long)
'内容:
'   高さの設定
'
'備考:
'
    m_Height = RHS
End Property

Public Property Get Left() As Long
'内容:
'   PMDの左アドレス
'
'備考:
'
    Left = m_X
End Property

Public Property Let Left(ByVal RHS As Long)
    m_Width = Right - RHS + 1
    m_X = RHS
End Property

Public Property Get Right() As Long
'内容:
'   PMDの右アドレス
'
'備考:
'
    Right = m_X + m_Width - 1
End Property

Public Property Let Right(ByVal RHS As Long)
    m_Width = RHS - Left + 1
End Property

Public Property Get Top() As Long
'内容:
'   PMDの上アドレス
'
'備考:
'
    Top = m_Y
End Property

Public Property Let Top(ByVal RHS As Long)
    m_Height = Bottom - RHS + 1
    m_Y = RHS
End Property

Public Property Get Bottom() As Long
'内容:
'   PMDの下アドレス
'
'備考:
'
    Bottom = m_Y + m_Height - 1
End Property

Public Property Let Bottom(ByVal RHS As Long)
    m_Height = RHS - Top + 1
End Property

Public Property Get Size() As Long
'内容:
'   PMDのサイズ(Read Only)
'
'備考:
'
    Size = m_Width * m_Height
End Property
