VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �f�[�^�G�N�X�v���[���[�����[�U�[�ɒ񋟂���ׂ��@�\���`����C���^�[�t�F�C�X
'
'�ړI:
'   �G�N�X�v���[���[�̎��Ԃ����̃C���^�[�t�F�C�X���������񋟋@�\��ۏႷ��
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬
'   2008/11/11�@�S�R�[�h���s�E����OK
'   2008/12/11�@�f�[�^�c���[�̕��ނ���R�ꂽ�c�[���J�e�S���̎擾�@�\�ǉ�
'   2008/12/15�@�f�[�^�c���[�̊O���o�͋@�\�ǉ�
'   2008/12/17�@12/11�ǉ��̋@�\���폜
'
'�쐬��:
'   0145206097
'
Option Explicit

Public Sub ConstructDataTree(ByVal dataList As Collection)
'���e:
'   �f�[�^�c���[�\�z�̎��s
'
'�p�����[�^:
'[dataList]   In  CDataSheet�^�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'   �\�ߓo�^���ꂽ�f�[�^�c���[�ɏ]���f�[�^���X�g�𕪗ނ���
'
End Sub

Public Property Get DataTree() As Collection
'���e:
'   �f�[�^�c���[�̎擾
'
'�߂�l�F
'   CToolSupplier�^�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'
End Property

Public Sub WriteDataTree(ByVal writer As IParameterWriter)
'���e:
'   �c���[�f�[�^�̏o�͂����s
'
'�p�����[�^:
'[writer]   In  �c���[�f�[�^�̏o�͐�ƂȂ郉�C�^�[
'
'���ӎ���:
'
'
End Sub