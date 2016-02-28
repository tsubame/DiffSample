VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PatDriveFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �ΏۃV�[�g��
Private Const SheetName = "PatDriveFormat"

' ��͊J�n�ʒu
Private Const BeginCol = 2
Private Const BeginRow = 4

' �ŏI�s����
Private Const SpaceRowMax = 3

' �e�[�u���̔z��
Private Tables_() As PatDriveFormatTable

Private Sub Class_Initialize()
    ReDim Tables_(0)
End Sub

Private Sub Class_Terminate()
    Erase Tables_
End Sub

Private Function GetChansPinNames(ByRef wbook As Workbook) As String()
    
    Dim wsheet As Worksheet
    Dim names() As String
    Dim Name As String
    Dim ii As Long
    
On Error GoTo ReleaseObjects:

    ' Chans�V�[�g��PinName���擾
    For ii = 1 To wbook.Sheets.Count
        If "Chans" = wbook.Worksheets(ii).Name Then
            Set wsheet = wbook.Worksheets(ii)
            Exit For
        End If
    Next
    
    If wsheet Is Nothing Then
        GetChansPinNames = names
        Exit Function
    End If
    
    ' �s���������߂�
    ' �Z��B7����PinName�����͂���Ă�����̂Ƃ���
    For ii = 7 To wsheet.Rows.Count
        Name = Trim(wsheet.Cells(ii, 2))
        If Name = "" Then
            Exit For
        End If
    Next
    
    ReDim names(ii - 8)
    For ii = 0 To UBound(names)
        names(ii) = wsheet.Cells(ii + 7, 2)
    Next
    
    GetChansPinNames = names
    
ReleaseObjects:
    Set wsheet = Nothing
    
End Function

Public Sub Initialize(ByRef wbook As Workbook)

On Error GoTo ReleaseObjects

    Dim wsheet As Worksheet
    
    Dim ii As Long
    Dim ci As Long ' Column Index
    Dim ri As Long ' Row Index
    Dim cirs As Long ' Column Index Range Start
    Dim rirs As Long ' Row Index Range Start
    Dim cire As Long ' Column Index Range End
    Dim rire As Long ' Row Index Range End
    Dim space_count As Integer
    Dim table_name As String
    Dim str_tmp As String
    Dim pin_names() As String
    
    ' Chans�V�[�g��PinName�ꗗ���擾
    ' �������݂��Ȃ��ꍇ�̓A�N�Z�X���ɗ�O�ƂȂ�
    pin_names = GetChansPinNames(wbook)
    
    ' �ΏۃV�[�g�̎擾
    For ii = 1 To wbook.Sheets.Count
        If SheetName = wbook.Worksheets(ii).Name Then
            Set wsheet = wbook.Worksheets(ii)
            Exit For
        End If
    Next
    
    ' �ΏۃV�[�g���Ȃ���Ώ������I��
    If wsheet Is Nothing Then
        Exit Sub
    End If
    
    ' Table�̍쐬
    ci = BeginCol
    cirs = -1
    cire = -1
    rirs = -1
    rire = -1
    space_count = 0
    table_name = ""
    For ri = BeginRow To wsheet.Rows.Count
    
        ' �V�[�g��(�e�[�u����)�擾
        str_tmp = Trim(wsheet.Cells(ri, ci))
        
        ' ��s�`�F�b�N
        If Len(str_tmp) > 0 Then
        
            ' �V�[�g���̓��͗L��
            space_count = 0
            
            ' Range�J�n�Z���̐ݒ�
            If table_name = "" Then
                        
                table_name = str_tmp
                cirs = ci
                rirs = ri
                
                ' ������'_'�Ŏn�܂�s�̌���
                Dim ri_ As Long
                For ri_ = ri To wsheet.Rows.Count
                    If InStr(1, Trim(wsheet.Cells(ri_, ci + 1)), "_") = 1 Then
                        Exit For
                    End If
                Next
                                
                ' �ŏI��̌���
                For cire = ci + 1 To wsheet.Columns.Count
                    ' �󔒃Z���̒��O���ŏI��Ƃ���
                    If Trim(wsheet.Cells(ri_, cire)) = "" Then
                        cire = cire - 1
                        Exit For
                    End If
                Next
                
            ' �V�[�g�������̍s�ƈقȂ�ꍇ�̓e�[�u���̍쐬
            ElseIf table_name <> str_tmp Then
            
                ' �e�[�u���̍ŏI�s�����ݒl��-1�Ƃ���
                rire = ri - 1
                
                ' �e�[�u���̒ǉ�
                Call AddTable(wsheet.Range( _
                    wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)), _
                    pin_names)
            
            End If
                        
        Else
        
            ' ��s
            
            ' ���O�܂ł̃f�[�^�Ńe�[�u���̍쐬
            If Not table_name = "" Then
                rire = ri - 1
            End If
            
            space_count = space_count + 1
            
        End If
        
        ' rire���ݒ肳��Ă���΃e�[�u�����쐬
        If rire > 0 Then
        
            Call AddTable(wsheet.Range( _
                wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)), _
                pin_names)
            
            ' �ϐ�������
            table_name = ""
            rirs = -1
            cirs = -1
            cire = -1
            rire = -1
            
        End If
        
        ' ��s���w��s�����A�������ꍇ�̓��[�v�𔲂���
        If space_count >= SpaceRowMax Then
            Exit For
        End If
    
    Next
        
ReleaseObjects:

    Set wsheet = Nothing

End Sub

Private Sub AddTable( _
    values As Range, _
    pin_names() As String)

On Error GoTo ReleaseObjects

    ' Table�쐬
    Dim table As PatDriveFormatTable
    Set table = New PatDriveFormatTable
    Call table.Initialize(values, pin_names)

    ' �쐬����Table���i�[
    ' �����l�͐擪��Nothing�̂��߁ANothing�̏ꍇ��Redim���s��Ȃ�
    If Not Tables_(0) Is Nothing Then
        ReDim Preserve Tables_(UBound(Tables_) + 1)
    End If
        
    Set Tables_(UBound(Tables_)) = table
            
ReleaseObjects:
    Set table = Nothing
    
End Sub

' �l�̎擾
' @param sheet_name �V�[�g��(�e�[�u����)
' @param event_name �C�x���g��
' @param identifier ���ʎq
' @param event_index �C�x���g�C���f�b�N�X
' @return �l
Public Function GetValue( _
    sheet_name As String, _
    event_name As String, _
    identifier As String, _
    Optional event_index As Integer = 0) As String
    
    Dim ti As Integer
    For ti = 0 To UBound(Tables_)
        If sheet_name = Tables_(ti).GetName() Then
            GetValue = Tables_(ti).GetValue(event_name, identifier, event_index)
            Exit Function
        End If
    Next
        
    GetValue = ""

End Function

' Table�̎擾
' @param sheet_name �V�[�g��(�e�[�u����)
' @return !Nothing PatDriveFormatTable
' @return Nothing ��v����Table����
Public Function GetTable(sheet_name As String) As PatDriveFormatTable

    Dim ii As Long
    For ii = 0 To UBound(Tables_)
        If Tables_(ii).GetName() = sheet_name Then
            Set GetTable = Tables_(ii)
            Exit Function
        End If
    Next
    
    Set GetTable = Nothing
    
End Function

' �STable�̎擾
' @return �STable
Public Function GetTables() As PatDriveFormatTable()
    GetTables = Tables_
End Function

' �[�q��(Output�̂�)�̎擾
' @param sheet_name �V�[�g��(�e�[�u����)
' @param identifier ���ʎq
' @param pin_number �s���ԍ�
' @return �[�q��
Public Function GetPinName( _
    sheet_name As String, _
    identifier As String, _
    pin_number As Long) As String
    
    Dim table As PatDriveFormatTable
    Set table = GetTable(sheet_name)
    If table Is Nothing Then
        GetPinName = ""
    Else
        GetPinName = table.GetPinName(identifier, pin_number)
    End If
    
    Set table = Nothing
    
End Function