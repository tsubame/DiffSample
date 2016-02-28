VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPowerSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   CPowerSquenceWait�N���X
'
'�ړI:
'   IPowerSequenceItem�̊Ǘ�����
'   �V�[�P���X�̎��s���s��
'
'�쐬��:
'   2011/12/05 Ver0.1 D.Maruyama
'   2012/11/02 Ver0.2 H.Arikawa
'                       �EExcecuteForUS�ǉ��B
'
Option Explicit

Private m_Name As String
Private m_Collection As Collection

Public Sub InitializeThisClass(ByVal strPowerSequenceName As String)
    
    Set m_Collection = Nothing
    Set m_Collection = New Collection
    
    m_Name = strPowerSequenceName
        
End Sub

Public Sub Add(ByRef seq As IPowerSequenceItem)
        
    Call m_Collection.Add(seq)
    
End Sub

Property Get Name()

    Name = m_Name
 
End Property

Public Sub ExcecuteForUS(ByVal strPowerConditionName As String)

    Dim seq As IPowerSequenceItem
    
    For Each seq In m_Collection
        seq.ExcecuteForUS (strPowerConditionName)
    Next seq
    
End Sub

Public Sub Execute(ByVal strPowerConditionName As String)

    Dim seq As IPowerSequenceItem
    
    For Each seq In m_Collection
        seq.Excecute (strPowerConditionName)
    Next seq
    
End Sub
Private Sub Class_Initialize()

    Set m_Collection = Nothing
    
End Sub

Private Sub Class_Terminate()

    Set m_Collection = Nothing

End Sub