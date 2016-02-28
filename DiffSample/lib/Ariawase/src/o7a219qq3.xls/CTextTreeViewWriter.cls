VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextTreeViewWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   ���[�N�V�[�g�i�r�Q�[�V�����c���[�r���[�̃e�L�X�g�p���C�^�[�N���X
'
'�ړI:
'   �c���[�r���[�f�[�^���e�L�X�g�֏o�͂���
'
'   Revision History:
'   Data        Description
'   2009/06/15  �쐬
'
'�쐬��:
'   0145206097
'
Option Explicit

Implements IParameterWriter
Implements IFileStream

Dim mFileNo As Integer
Dim mFileName As String

Private Const FILE_NAME = "WorksheetTreeViewer.csv"

Public Sub OpenFile(ByVal filePath As String)
    On Error GoTo ErrorHandler
    mFileNo = FreeFile
    mFileName = filePath & "\" & getToday & "_" & FILE_NAME
    Open mFileName For Output As mFileNo
    Exit Sub
ErrorHandler:
    Close mFileNo
    Err.Raise 9999, "CTextTreeViewWriter.OpenFile", "Opening File Is Failed"
End Sub

Public Sub CloseFile()
    Close mFileNo
End Sub

Private Sub IParameterWriter_WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
    '### ���̃C���^�[�t�F�C�X�@�\�̓���𖳎����� ###################
End Sub

Private Sub IParameterWriter_WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
    '### ���̃��C�^�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsDouble", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
    '### ���̃��C�^�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsInteger", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
    '### ���̃��C�^�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CTextTreeViewWriter.AsIParameterWriter.WriteAsLong", "Does Not Support This Method !"
End Sub

Private Sub IParameterWriter_WriteAsString(ByVal dataName As String, ByVal strData As String)
    '### �t�@�C���փf�[�^�o�� #######################################
    Select Case dataName:
        Case SUPPLIER_NAME:
            Print #mFileNo, strData;
        Case TOOL_NAME:
            Print #mFileNo, "," & strData;
        Case sheet_name:
            Print #mFileNo, ",," & strData;
    End Select
End Sub

Private Sub IFileStream_MoveNext()
    '### �f�[�^�̉��s ###############################################
    Print #mFileNo,
End Sub

Private Function IFileStream_IsEOR() As Boolean
    '### ���̃C���^�[�t�F�C�X�@�\�̓���𖳎����� ###################
End Function

Private Sub IFileStream_SetLocation(ByVal dataName As String)
    '### ���̃��C�^�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CTextTreeViewWriter.AsIFileStream.SetLocation", "Does Not Support This Method !"
End Sub

Private Sub IFileStream_Rewind()
    '### ���̃��C�^�[�͂��̃C���^�[�t�F�C�X�@�\��ۏႵ�Ȃ� #########
    Err.Raise 999, "CTextTreeViewWriter.AsIFileStream.Rewind", "Does Not Support This Method !"
End Sub

Public Function AsIParameterWriter() As IParameterWriter
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IParameterWriter�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIParameterWriter = Me
End Function

Public Function AsIFileStream() As IFileStream
'���e:
'   �������g���������Ă���C���^�[�t�F�C�X�I�u�W�F�N�g�^�ւ̃L���X�g
'
'�߂�l�F
'   IFileStreamr�^�I�u�W�F�N�g
'
'���ӎ���:
'
    Set AsIFileStream = Me
End Function

Private Function getToday() As String
    getToday = Format(DateTime.Now, "yymmdd") & "_" & Format(DateTime.Now, "hhnnss")
End Function