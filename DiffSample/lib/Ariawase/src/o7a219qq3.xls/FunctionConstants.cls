VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �ΏۃV�[�g��
Private Const SheetName = "FunctionConstants"

' ��͊J�n�ʒu
Private Const BeginCol = 2
Private Const BeginRow = 4

' �ŏI�s����
Private Const SpaceRowMax = 3

' �e�[�u���̔z��
Private Tables_() As FunctionConstantsTable

Private Sub Class_Initialize()
    ReDim Tables_(0)
End Sub

Private Sub Class_Terminate()
    Erase Tables_
End Sub

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
                    wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)))
            
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
                wsheet.Cells(rirs, cirs), wsheet.Cells(rire, cire)))
            
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

Private Sub AddTable(values As Range)

On Error GoTo ReleaseObjects

    ' Table�쐬
    Dim table As FunctionConstantsTable
    Set table = New FunctionConstantsTable
    Call table.Initialize(values)

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
' @param sheet_name �V�[�g��
' @param constant_name �萔��
' @return �萔 ������Ȃ��ꍇ��Empty
Public Function GetValue( _
    sheet_name As String, _
    constant_name As String) As Variant
        
On Error GoTo ReleaseObjects

    Dim table As FunctionConstantsTable
    Set table = GetTable(sheet_name)
    If Not table Is Nothing Then
        GetValue = table.GetValue(constant_name)
    End If
    
ReleaseObjects:
    Set table = Nothing
    
End Function

' Table�̎擾
' @param sheet_name �V�[�g��
Public Function GetTable(sheet_name As String) As FunctionConstantsTable
    
    Dim ii As Integer
    
    For ii = 0 To UBound(Tables_)
    
        If sheet_name = Tables_(ii).GetName() Then
        
            Set GetTable = Tables_(ii)
            Exit Function
            
        End If
    Next
    
    Set GetTable = Nothing
    
End Function

' �STable�̎擾
Public Function GetTables() As FunctionConstantsTable
    GetTables = Tables_
End Function