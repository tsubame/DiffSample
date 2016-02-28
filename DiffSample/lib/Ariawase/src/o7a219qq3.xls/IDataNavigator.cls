VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDataNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�T�v:
'   �i�r�Q�[�V���������[�U�[�ɒ񋟂���ׂ��@�\���`����C���^�[�t�F�C�X
'
'�ړI:
'   �i�r�Q�[�V�����̎��Ԃ����̃C���^�[�t�F�C�X���������񋟋@�\��ۏႷ��
'
'   Revision History:
'   Data        Description
'   2008/11/10  �쐬
'   2008/11/11�@�S�R�[�h���s�E����OK
'   2008/12/11�@�C���f�b�N�X�w��ł̃f�[�^�V�[�g�擾�@�\�ǉ�
'   2009/02/20�@���d�l�ύX
'               �@���\�b�h�E�v���p�e�B���̕ύX
'               ���@�\�ǉ�
'               �@HistoryIndex�v���p�e�B�ǉ��i�q�X�g���̃J�����g�C���f�b�N�X�擾�@�\�j
'
'�쐬��:
'   0145206097
'
Option Explicit

Public Sub Add(ByVal dataSht As CDataSheet)
'���e:
'   �i�r�Q�[�V�����q�X�g���ւ̒ǉ�
'
'�p�����[�^:
'[dataSht]   In  CDataSheet�^�I�u�W�F�N�g
'
'���ӎ���:
'   �q�X�g�������ő�q�X�g�����𒴂���ꍇ�͌Â��q�X�g�����玩���I�ɍ폜�����
'
End Sub

Public Property Get History() As Collection
'���e:
'   �i�r�Q�[�V�����q�X�g���̎擾
'
'�߂�l�F
'   CDataSheet�^�I�u�W�F�N�g�̃R���N�V����
'
'���ӎ���:
'   �q�X�g������0�̏ꍇ��Nothing��Ԃ�
'
End Property

Public Function DataSheet(ByVal hIndex As Long) As CDataSheet
'���e:
'   �C���f�b�N�X�w��ɂ��q�X�g�����̃f�[�^�V�[�g�擾
'
'�p�����[�^:
'[hIndex]   In  �f�[�^�V�[�g�̃C���f�b�N�X�ԍ�
'
'�߂�l�F
'   CDataSheet�^�I�u�W�F�N�g
'
'���ӎ���:
'   �C���f�b�N�X���ő�q�X�g�����𒴂���A�܂���1�����̏ꍇ��Nothing��Ԃ�
'   �w�肵���C���f�b�N�X�̃f�[�^�V�[�g���폜����Ă���ꍇ��Nothing��Ԃ�
'
End Function

Public Property Get ForwardDataSheet() As CDataSheet
'���e:
'   �q�X�g�����̌��݂̃C���f�b�N�X�̈�O�̃f�[�^�V�[�g�擾
'
'�߂�l�F
'   CDataSheet�^�I�u�W�F�N�g
'
'���ӎ���:
'   �C���f�b�N�X���ő�q�X�g�����𒴂����ꍇ��Nothing��Ԃ�
'   �J�����g�̃C���f�b�N�X�̃f�[�^�V�[�g���폜����Ă���ꍇ�A
'   �q�X�g���O���̍폜����Ă��Ȃ��f�[�^�V�[�g�܂�
'   �C���f�b�N�X���C���N�������g���ăI�u�W�F�N�g��Ԃ�
'
End Property

Public Property Get BackwardDataSheet() As CDataSheet
'���e:
'   �q�X�g�����̌��݂̃C���f�b�N�X�̈���̃f�[�^�V�[�g�擾
'
'�߂�l�F
'   CDataSheet�^�I�u�W�F�N�g
'
'���ӎ���:
'   �C���f�b�N�X��1�ȉ��̏ꍇ��Nothing��Ԃ�
'   �J�����g�̃C���f�b�N�X�̃f�[�^�V�[�g���폜����Ă���ꍇ�A
'   �q�X�g������̍폜����Ă��Ȃ��f�[�^�V�[�g�܂�
'   �C���f�b�N�X���f�N�������g���ăI�u�W�F�N�g��Ԃ�
'
End Property

Public Property Get IsEnableForward() As Boolean
'���e:
'   �q�X�g�����̃J�����g�C���f�b�N�X���O���̃f�[�^�V�[�g�擾���\���ǂ����̔���
'
'�߂�l�F
'   �O���̃f�[�^�擾���\�ł����True�A�s�ł����False
'
'���ӎ���:
'   �C���f�b�N�X�͈̔͂��L���ł����Ă��A�q�X�g�����̃J�����g���O����
'   �f�[�^�V�[�g���S�č폜����Ă���ꍇ��False��Ԃ�
'
End Property

Public Property Get IsEnableBackward() As Boolean
'���e:
'   �q�X�g�����̃J�����g�C���f�b�N�X������̃f�[�^�V�[�g�擾���\���ǂ����̔���
'
'�߂�l�F
'   ����̃f�[�^�擾���\�ł����True�A�s�ł����False
'
'���ӎ���:
'   �C���f�b�N�X�͈̔͂��L���ł����Ă��A�q�X�g�����̃J�����g�������
'   �f�[�^�V�[�g���S�č폜����Ă���ꍇ��False��Ԃ�
'
End Property

Public Property Get HistoryIndex() As Long
'���e:
'   �q�X�g���̃J�����g�C���f�b�N�X�擾
'
'�߂�l�F
'   �J�����g�C���f�b�N�X�l
'
'���ӎ���:
'
End Property

Public Sub Clear()
'���e:
'   �q�X�g���̏�����
'
'���ӎ���:
'
End Sub
