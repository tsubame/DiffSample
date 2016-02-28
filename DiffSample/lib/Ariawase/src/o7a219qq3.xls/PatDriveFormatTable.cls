VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormatTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NoHeader = "_No_"
Private Const EventHeader = "_event_"

Private Name_ As String
Private Columns_() As PatDriveFormatColumn
Private ColumnGroups_() As PatDriveFormatColumnGroup
Private Events_() As PatDriveFormatEvent
Private EventGroups_() As PatDriveFormatEventGroup
Private Values_() As String

Private Sub Class_Initialize()
    Name_ = ""
    ReDim Columns_(0)
    ReDim ColumnGroups_(0)
    ReDim Events_(0)
    ReDim EventGroups_(0)
End Sub

Private Sub Class_Terminate()
    Erase Columns_
    Erase ColumnGroups_
    Erase Events_
    Erase EventGroups_
End Sub

' ������
' @param values ��͑ΏۂƂȂ�Z���͈�
' @param pin_names Chans�V�[�g��PinName���X�g
Public Sub Initialize( _
    ByRef values As Range, _
    pin_names() As String)

    Dim pi As Long
    Dim ri As Long
    Dim ci As Long
    Dim ni As Long ' No Index
    Dim ei As Long ' Event Index
    Dim hri As Long ' Header Row Index
    Dim str_tmp As String
    Dim chans_pin_name As String
    
On Error GoTo ReleaseObjects
    
    If values Is Nothing Then
        Exit Sub
    End If
    
    ' ����
    Name_ = Trim(values(1, 1))
            
    ' ���o���̈ʒu���擾
    ' 2��ڂ�"_"�Ŏn�܂�s�̎擾
    For ri = 1 To values.Rows.Count
        If InStr(1, Trim(values.Cells(ri, 2)), "_", vbBinaryCompare) = 1 Then
            Exit For
        End If
    Next
    
    ' Column�쐬����
    For ci = 2 To values.Columns.Count
    
        ' Event�쐬�p�̃C���f�b�N�X���擾
        str_tmp = Trim(values.Cells(ri, ci))
        If StrComp(str_tmp, NoHeader, vbBinaryCompare) = 0 Then
            ni = ci
        ElseIf StrComp(str_tmp, EventHeader, vbBinaryCompare) = 0 Then
            ei = ci
        Else
            
On Error GoTo ReleaseColumn

            Dim pos As Integer
            Dim pin_name As String
            pin_name = values.Cells(ri - 1, ci)
            
            ' Column�̒[�q���ƈ�v����Chans��PinName���擾
            For pi = 0 To UBound(pin_names)
                ' '_'�ȍ~�̕�������擾
                chans_pin_name = pin_names(pi)
                pos = InStr(1, chans_pin_name, "_", vbTextCompare)
                If pos > 0 Then
                    If Mid(chans_pin_name, pos + 1) = pin_name Then
                        pin_name = chans_pin_name
                        Exit For
                    End If
                End If
            Next
    
            ' Column�̍쐬
            ' ���͎̂��ʎq�̈��̃Z��
            ' �C���f�b�N�X��0�n�܂�ɂ��邽��-2�����Ă���
            Dim pdf_column As PatDriveFormatColumn
            Set pdf_column = New PatDriveFormatColumn
            Call pdf_column.Initialize( _
                values.Cells(ri, ci), _
                pin_name, _
                ci)
            
            If Not Columns_(0) Is Nothing Then
                ReDim Preserve Columns_(UBound(Columns_) + 1)
            End If
            Set Columns_(UBound(Columns_)) = pdf_column
                
            ' �O���[�v�֒ǉ�
            Dim pdf_group As PatDriveFormatColumnGroup
            Dim group_name As String
            group_name = GetColumnGroupName(pdf_column.GetIdentifier())
            If group_name <> "" Then
                ' �O���[�v���Ώۂł���ΑΏۃO���[�v�̎擾
                Set pdf_group = GetColumnGroup(group_name)
                If pdf_group Is Nothing Then
                    ' �O���[�v�������ꍇ�͍쐬
                    Set pdf_group = New PatDriveFormatColumnGroup
                    Call pdf_group.Initialize(group_name)
                    
                    If Not ColumnGroups_(0) Is Nothing Then
                        ReDim Preserve ColumnGroups_(UBound(ColumnGroups_) + 1)
                    End If
                    Set ColumnGroups_(UBound(ColumnGroups_)) = pdf_group
                End If
                
                ' �O���[�v�֒ǉ�
                Call pdf_group.AddColumn(pdf_column)
                            
            End If
            
ReleaseColumn:
            Set pdf_group = Nothing
            Set pdf_column = Nothing
            
        End If
    Next
    
    ' Event�쐬����
    For ri = ri + 1 To values.Rows.Count
    
On Error GoTo ReleaseEvent

        ' Event�̍쐬
        Dim pdf_event As PatDriveFormatEvent
        Set pdf_event = New PatDriveFormatEvent
        Call pdf_event.Initialize( _
            CLng(values.Cells(ri, ni)), values.Cells(ri, ei), ri)
            
        If Not Events_(0) Is Nothing Then
            ReDim Preserve Events_(UBound(Events_) + 1)
        End If
        Set Events_(UBound(Events_)) = pdf_event
            
        ' EventGroup�̎擾
        Dim group As PatDriveFormatEventGroup
        Set group = GetEventGroup(pdf_event.GetName())
        If group Is Nothing Then
            ' ���݂��Ȃ��ꍇ�͐V�K�쐬
            Set group = New PatDriveFormatEventGroup
            Call group.Initialize(pdf_event.GetName())
            
            ' �z��ɒǉ�
            If Not EventGroups_(0) Is Nothing Then
                ReDim Preserve EventGroups_(UBound(EventGroups_) + 1)
            End If
            Set EventGroups_(UBound(EventGroups_)) = group
        End If
        
        ' EventGroup��Event��ǉ�
        Call group.AddEvent(pdf_event)
                
ReleaseEvent:
        Set pdf_event = Nothing
        Set group = Nothing
            
    Next
    
    ' �l�̎擾
    ReDim Values_(UBound(Events_), UBound(Columns_))
    For ri = 0 To UBound(Events_)
        For ci = 0 To UBound(Columns_)
            Values_(ri, ci) = values.Cells( _
                Events_(ri).GetRangeIndex(), _
                Columns_(ci).GetRangeIndex())
        Next
    Next
        
ReleaseObjects:
    ' Nothing
    
End Sub

' EventGroup�̎擾
' @param group_name �O���[�v��(�C�x���g���Ɠ���)
' @return !Nothing PatDriveFormatEventGroup
' @return Nothing �O���[�v��������Ȃ�����
Public Function GetEventGroup(group_name As String) As PatDriveFormatEventGroup

    ' �擪��Nothing�ł����Group�͑��݂��Ȃ�
    If EventGroups_(0) Is Nothing Then
        Set GetEventGroup = Nothing
        Exit Function
    End If

    Dim ii As Integer
    For ii = 0 To UBound(EventGroups_)
        If EventGroups_(ii).GetName() = group_name Then
            Set GetEventGroup = EventGroups_(ii)
            Exit Function
        End If
    Next
    
    Set GetEventGroup = Nothing
    
End Function

' �SEventGroup�̎擾
' @return �SEventGroup
Public Function GetEventGroups() As PatDriveFormatEventGroup()
    GetEventGroups = EventGroups_
End Function

' ���̎擾
' @return ����
Public Function GetName() As String
    GetName = Name_
End Function

' ColumnGroup�̎擾
' @param group_name �O���[�v��
' @return !Nothing PatDriveFormatColumnGroup
' @return Nothing �O���[�v��������Ȃ�����
Public Function GetColumnGroup(group_name As String) As PatDriveFormatColumnGroup

    If ColumnGroups_(0) Is Nothing Then
        Set GetColumnGroup = Nothing
        Exit Function
    End If
    
    Dim ii As Long
    For ii = 0 To UBound(ColumnGroups_)
        If group_name = ColumnGroups_(ii).GetName() Then
            Set GetColumnGroup = ColumnGroups_(ii)
            Exit Function
        End If
    Next
    
    Set GetColumnGroup = Nothing
    
End Function

' �SColumnGroup�̎擾
' @return �SColumnGroup
Public Function GetColumnGroups() As PatDriveFormatColumnGroup()
    GetColumnGroups = ColumnGroups_
End Function

' �O���[�v���̎擾
' ���ʎq����O���[�v�����擾����
' @param column_identifier ���ʎq
' @return �O���[�v��(�󔒎��̓O���[�v�Ώۖ���)
Private Function GetColumnGroupName(column_identifier As String) As String
    
    ' ���ʎq������̃p�^�[���ł��邩�`�F�b�N
    If Not column_identifier Like "_*[0-9]_" Then
        GetColumnGroupName = ""
        Exit Function
    End If
    
    ' ���2��������������̂��O���[�v���Ƃ���
    GetColumnGroupName = Mid(column_identifier, 1, Len(column_identifier) - 2)
    
End Function

' �l�̎擾
' @param event_name �C�x���g��
' @param identifier ���ʎq
' @param event_index �C�x���g�̃C���f�b�N�X
' @return �l
Public Function GetValue( _
    event_name As String, _
    identifier As String, _
    Optional event_index As Integer = 0)
    
    Dim ri As Long
    Dim ci As Long
    Dim no As Long
    Dim ii As Long
    
    no = -1

    ' event_name�������݂̂ł���΃C�x���g�ԍ��Ƃ��ĉ��߂���
    For ii = 1 To Len(event_name)
        If Not Mid(event_name, ii, 1) Like "[0-9]" Then
            Exit For
        End If
    Next
    If ii > Len(event_name) Then
        no = CLng(event_name)
    End If
    
    ' ��v����C�x���g�̌���
    If no = -1 Then
        Dim ei As Integer
        ei = 0
        ' �C�x���g�����L�[�Ƃ���ꍇ�͓����C�x���g���̃C���f�b�N�X���l��
        For ri = 0 To UBound(Events_)
            If event_name = Events_(ri).GetName() Then
                If event_index <= ei Then
                    Exit For
                End If
                ei = ei + 1
            End If
        Next
    Else
        For ri = 0 To UBound(Events_)
            If event_name = Events_(ri).GetNo() Then
                Exit For
            End If
        Next
    End If
    
    If ri > UBound(Events_) Then
        GetValue = ""
        Exit Function
    End If
    
    ' ��v�����̌���
    For ci = 0 To UBound(Columns_)
        If identifier = Columns_(ci).GetIdentifier() Then
            Exit For
        End If
    Next
    
    If ci > UBound(Columns_) Then
        GetValue = ""
        Exit Function
    End If
    
    GetValue = Values_(ri, ci)
    
End Function

' �[�q���̎擾
' @param identifier ���ʎq
' @param pin_number �s���ԍ�
' @return �[�q��
Public Function GetPinName( _
    identifier As String, _
    pin_number As Long) As String
    
    Dim id As String
    Dim ci As Long
    
    id = "_" + identifier + CStr(pin_number) + "_"
    
    ' ��v�����̌���
    For ci = 0 To UBound(Columns_)
        If id = Columns_(ci).GetIdentifier() Then
            GetPinName = Columns_(ci).GetName()
            Exit Function
        End If
    Next
    
    GetPinName = ""
    
End Function