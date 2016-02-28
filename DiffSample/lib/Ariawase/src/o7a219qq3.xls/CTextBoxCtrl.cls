VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextBoxCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/* テキストボックスオブジェクトのラッパークラス */
'/* これを利用すると複数あるテキストボックスをまとめて扱うことができて */
'/* ボックスのイベントをダラダラと記述することもなくなる */

Option Explicit

'/* なぜかTextBox型で受けると取れるイベントが少なくなる… */
Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Dim mBoxId As Long
Dim mInitialValue As String
Dim mIsSystemEdit As Boolean

Private Const ORIGINAL_COLOR = &H80000005
Private Const ALERT_COLOR = &H8080FF

Public Sub Initialize(ByVal pTextBox As MSForms.TextBox, ByVal pId As Long)
    Set mTextBox = pTextBox
    mBoxId = pId
    resetProperty
    With mTextBox
        If mBoxId = 1 Then .SetFocus
        .TabIndex = mBoxId - 1
        .EnterFieldBehavior = fmEnterFieldBehaviorSelectAll
        '/* IMEモードはOFF */
        .IMEMode = fmIMEModeDisable
    End With
    mIsSystemEdit = False
End Sub

Private Sub Class_Terminate()
    resetProperty
    Set mTextBox = Nothing
    mIsSystemEdit = False
    mBoxId = 0
End Sub

Property Let Value(ByVal pItem As String)
    mIsSystemEdit = True
    mTextBox.Text = pItem
    mInitialValue = mTextBox.Text
    mIsSystemEdit = False
End Property

Property Get Value() As String
    Value = mTextBox.Text
End Property

Property Let enabled(ByVal pStatus As Boolean)
    With mTextBox
        .enabled = pStatus
        If Not (pStatus) Then
            .BackColor = ALERT_COLOR
        Else
            .BackColor = ORIGINAL_COLOR
        End If
    End With
End Property

Property Get enabled() As Boolean
    enabled = mTextBox.enabled
End Property

Property Get id() As Long
    id = mBoxId
End Property

Private Sub resetProperty()
    With mTextBox
        .BackColor = ORIGINAL_COLOR
        .enabled = True
        Me.Value = ""
    End With
End Sub

Private Sub mTextBox_Change()
    If mIsSystemEdit Then Exit Sub
    If mInitialValue <> mTextBox.Text Then
        mTextBox.BackColor = ALERT_COLOR
    Else
        mTextBox.BackColor = ORIGINAL_COLOR
    End If
    
End Sub

Private Sub mTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    With mTextBox
        .SetFocus
        .SelStart = 0
        .SelLength = Len(mTextBox)
    End With
End Sub

