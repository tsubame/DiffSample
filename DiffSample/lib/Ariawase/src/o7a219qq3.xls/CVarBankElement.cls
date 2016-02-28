VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVarBankElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CVarBank にCollection.Addされるクラス
'   値やその他のパラメータを保持する
'   変数１つにつきこのインスタンス１つを割り当てる
'目的:
'   Public 変数の使用を回避する｡
'   TOPT はプロシージャ内で別のプロシージャに制御が移るため、Public変数に意図しない変更が発生する可能性が出てくる。
'   基本、読み取り専用にしておくことで不用意な書き込みを回避する。
'作成者:
'  tomoyoshi.takase
'作成日: 2010年4月22日

Option Explicit

Private m_Val As Variant         '値の格納用
Private m_ReadOnly As Boolean    '値の固定状態 格納用
Private m_Name As String         'キー格納用
Private m_Comment As String      'コメント格納用

Public Property Let val(pVal As Variant)
    m_ReadOnly = True
    m_Val = pVal
End Property

Public Property Get val() As Variant
    val = m_Val
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
End Property

Public Property Let ReadOnly(status As Boolean)
    m_ReadOnly = status
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Comment As String)
    m_Comment = Comment
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Sub Initialize(Name As String, val As Variant, Comment As String)
    m_Name = Name
    m_Comment = Comment
    m_Val = val
End Sub

Private Sub Class_Initialize()
    m_Name = ""
    m_ReadOnly = True
    m_Val = ""
End Sub

Private Sub Class_Terminate()
'dummy
End Sub
