Attribute VB_Name = "XLibImageMod"
'�T�v:
'   �摜�������C�u����
'
'�ړI:
'   ��ʓI�ɂ悭�g�������̓Z��
'
'�쐬��:
'   0145184004
'
Option Explicit

Private Const TMP_SIZE = 11

'2009/09/15 D.Maruyama ColorAll�����ǉ� ��������
Public Type SiteValues
    SiteValue() As Double
End Type

Public Type ColorAllResult
    color(TMP_SIZE) As SiteValues
End Type
'2009/09/15 D.Maruyama ColorAll�����ǉ� �����܂�

Public Sub Average( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   ���ϒl���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Average(retResult, srcColor, pFlgName)

End Sub

Public Sub sum( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   ���v�l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.sum(retResult, srcColor, pFlgName)

End Sub

Public Sub StdDev( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �W���΍����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.StdDev(retResult, srcColor, pFlgName)

End Sub

Public Sub GetPixelCount( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �Ώۂ̉�f�����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Num(retResult, srcColor, pFlgName)

End Sub

Public Sub Min( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Min(retResult, srcColor, pFlgName)

End Sub

Public Sub max( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ő�l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.max(retResult, srcColor, pFlgName)

End Sub

Public Sub MinMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retMin() As Double, ByRef retMax() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�A�ő�l����x�Ɏ擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retMin()]    OUT  Double�^:       �ŏ��l�i�[�p�z��(�T�C�g��)
'[retMax()]    OUT  Double�^:       �ő�l�i�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)

'=======2009/05/11 �ύX Maruyama ��������==============
'    Call srcPlane.Min(retMin, srcColor, pFlgName)
'    Call srcPlane.Max(retMax, srcColor, pFlgName)
    Call srcPlane.MinMax(retMin, retMax, srcColor, pFlgName)
'=======2009/05/11 �ύX Maruyama �����܂�==============

End Sub

Public Sub DiffMinMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�ƍő�l�̍����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.DiffMinMax(retResult, srcColor, pFlgName)

End Sub

Public Sub AbsMax( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult() As Double, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�ƍő�l�̓���Βl�̑傫�������擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[retResult()] OUT  Double�^:       ���ʊi�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AbsMax(retResult, srcColor, pFlgName)

End Sub

Public Sub Count( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, ByRef retResult() As Double, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "" _
)
'���e:
'   �����ɊY������_�̌����擾����B
'
'[srcPlane]         IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]          IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]         IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[countType]        IN   IdpCountType�^: �J�E���g�����w��
'[loLim]            IN   Variant�^:      �����l
'[hiLim]            IN   Variant�^:      ����l
'[limitType]        IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[retResult()]      OUT  Double�^:       ���ʊi�[�p�z��(���I�z��)
'[pFlgName]         IN   String�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Count(retResult, countType, loLim, hiLim, limitType, srcColor, pFlgName, pInputFlgName)

End Sub

'=======2009/05/19 Add Maruyama ��������==============
Public Sub CountForFlgBitImgPlane( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, ByRef retResult() As Double, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pInputFlgName As String = "" _
)
'���e:
'   �����ɊY������_�̌����擾����B(�t���O�r�b�g���C���[�W�v�����ɗ��Ă�)
'
'[srcPlane]         IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]          IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]         IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[countType]        IN   IdpCountType�^: �J�E���g�����w��
'[loLim]            IN   Variant�^:      �����l
'[hiLim]            IN   Variant�^:      ����l
'[limitType]        IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[retResult()]      OUT  Double�^:       ���ʊi�[�p�z��(���I�z��)
'[pFlgName]         IN   CImgPlane�^:       �o�̓t���O��
'[pFlgBit]        �@IN   Long�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call pFlgPlane.SetPMD(srcZone)
    Call srcPlane.CountForFlgBitImgPlane(retResult, countType, loLim, hiLim, pFlgPlane, pFlgBit, limitType, srcColor, pInputFlgName)

End Sub
'=======2009/05/19 Add Maruyama �����܂�==============

Public Sub PutFlag( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, _
    ByVal pFlgName As String, Optional ByVal pInputFlgName As String _
)
'���e:
'   �����ɊY������_�Ƀt���O�𗧂Ă�B(Count����t���O�𗧂Ă邱�Ƃ����ɓ���)
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[countType]   IN   IdpCountType�^: �J�E���g�����w��
'[loLim]       IN   Variant�^:      �����l
'[hiLim]       IN   Variant�^:      ����l
'[limitType]   IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[pFlgName]         IN   String�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'   srcColor��EEE_COLOR_ALL���w�肷���All�ŏ���������Flat�̌��ʂ�Ԃ��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PutFlag(pFlgName, countType, loLim, hiLim, limitType, srcColor, pInputFlgName)

End Sub

Public Sub Add( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜�����Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'   ��)
'       Add(dst, "ZONE3", EEE_COLOR_FLAT, src, "ZONE3" EEE_COLOR_FLAT, dst, "ZONE3_2", EEE_COLOR_FLAT)
'       ���̏ꍇdst�� ZONE3_2 �������ΏۂɂȂ�B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Add(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub AddConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal addVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜�Ɏw��l�𑫂��B
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[addVal]      IN   Variant�^:      ���Z�l
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Add(srcPlane, addVal, dstColor, srcColor)

End Sub

Public Sub Subtract( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜�����Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Subtract(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub SubtractConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal subVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜����w��l�������B
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[subVal]      IN   Variant�^:      ���Z�l
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Subtract(srcPlane, subVal, dstColor, srcColor)

End Sub

Public Sub Multiply( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜����Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub MultiplyConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal mulVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜�Ɏw��l���|����B
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[mulVal]      IN   Variant�^:      ��Z�l
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, mulVal, dstColor, srcColor)

End Sub

Public Sub MultiplyConstFlag( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal mulVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, Optional ByVal pInputFlgName As String = "" _
)
'���e:
'   �Ώۉ摜�Ɏw��l���|����B
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[mulVal]      IN   Variant�^:      ��Z�l
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Multiply(srcPlane, mulVal, dstColor, srcColor, , pInputFlgName)

End Sub

Public Sub Divide( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜�����Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Divide(srcPlane, srcPlane2, dstColor, srcColor, srcColor2)

End Sub

Public Sub DivideConst( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal divVal As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant _
)
'���e:
'   �Ώۉ摜���w��l�Ŋ���B
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[divVal]      IN   Variant�^:      ���Z�l
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'
'���l:
'   �e�v���[���ɂ͓������̂��w�肷�邱�Ƃ��\�B
'   ���̏ꍇ�]�[���w��͍Ō�̂��̂œ��ꂳ��Ă��܂��̂ŗv���ӁB
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Divide(srcPlane, divVal, dstColor, srcColor)

End Sub

Public Sub Median( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'���e:
'   �Ώۉ摜�Ƀ��f�B�A���t�B���^���|����B
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[Width]       IN   Long�^:         �t�B���^��
'[Height]      IN   Long�^:         �t�B���^����
'
'���l:
'
    Dim Center As Long

    Center = (width * height + 1) / 2

    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.RankFilter(srcPlane, width, height, Center, dstColor, srcColor)

End Sub

Public Sub MedianHV( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'���e:
'   �Ώۉ摜�Ƀ��f�B�A���t�B���^���|����B
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[Width]       IN   Long�^:         �t�B���^��
'[Height]      IN   Long�^:         �t�B���^����
'
'���l:
'
    Dim tmpPlane As CImgPlane
    Set tmpPlane = TheIDP.PlaneManager(srcPlane.planeGroup).GetFreePlane(srcPlane.BitDepth)
    
    Call Median(srcPlane, srcZone, srcColor, tmpPlane, srcZone, srcColor, width, 1)
    Call Median(tmpPlane, srcZone, srcColor, dstPlane, dstZone, dstColor, 1, height)
    
End Sub

Public Sub MedianVH( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal width As Long, ByVal height As Long _
)
'���e:
'   �Ώۉ摜�Ƀ��f�B�A���t�B���^���|����B
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[Width]       IN   Long�^:         �t�B���^��
'[Height]      IN   Long�^:         �t�B���^����
'
'���l:
'
    Dim tmpPlane As CImgPlane
    Set tmpPlane = TheIDP.PlaneManager(srcPlane.planeGroup).GetFreePlane(srcPlane.BitDepth)
    
    Call Median(srcPlane, srcZone, srcColor, tmpPlane, srcZone, srcColor, 1, height)
    Call Median(tmpPlane, srcZone, srcColor, dstPlane, dstZone, dstColor, width, 1)
    
End Sub

Public Sub Convolution( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal Kernel As String, Optional ByVal divVal As Long = 0 _
)
'���e:
'   �Ώۉ摜�ɃR���{�����[�V�����t�B���^���|����B
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[kernel]      IN   String�^:       �t�B���^��
'[divVal]      IN   Long�^:         ���߂��p�̒l
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Convolution(srcPlane, Kernel, dstColor, srcColor)

    If divVal <> 0 Then
        Call DivideConst(dstPlane, dstZone, dstColor, divVal, dstPlane, dstZone, dstColor)
    End If

End Sub

Public Sub ExecuteLUT( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal lutName As String _
)
'���e:
'   �Ώۉ摜�Ƀ��b�N�A�b�v�e�[�u�����|����B
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �Ώۃv���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[lutName]     IN   String�^:       LUT��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ExecuteLUT(srcPlane, lutName, dstColor, srcColor)

End Sub

Public Sub Copy( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional pmask As String = "")
'���e:
'   �Ώۉ摜�Ƀ��f�B�A���t�B���^���|����B
'
'[srcPlane]    IN   CImgPlane�^:    �R�s�[���v���[��
'[srcZone]     IN   String�^:       �R�s�[���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: �R�s�[���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    �R�s�[��v���[��
'[dstZone]     IN   String�^:       �R�s�[��v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: �R�s�[��v���[���̐F�w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.CopyPlane(srcPlane, dstColor, srcColor, , pmask)

End Sub

Public Sub WritePixel(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, ByVal writeVal As Double, Optional ByVal mask As Long = 0)
'���e:
'   �Ώۉ摜�ɒl���������ށB
'
'[dstPlane]     OUT CImgPlane�^:    �Ώۃv���[��
'[dstZone]      IN  String�^:       �Ώۃv���[���̃]�[���w��
'[dstColor]     IN  IdpColorType�^: �Ώۃv���[���̐F�w��
'[writeVal]     IN  Double�^:       �������ޒl
'[mask]         IN  Long�^:         �}�X�N�w��
'
'���l:
'   mask���w�肷���1�̗������r�b�g�͖��������B
'   ��)
'       WritePixel("vmcu00", "ZONE3", 0, &HFFF0)�@�Ƃ����ꍇ����4bit�݂̂�0���������܂��B
'
'   �{���͂��̋@�\�͎g��Ȃ��悤�ɂ��ׂ��Ǝv���B
'   ���v���O�����̒��ł�����g���Ă����̂œ��ꂽ���A
'   �r�b�g���Z���������̂ł����LOr,LAnd�Ȃǂ��p�ӂ���Ă��邩�炻����g���ׂ��B

    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.WritePixel(writeVal, dstColor, , , mask)

End Sub

Public Sub MultiMean( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal multiMeanFunc As IdpMultiMeanFunc, ByVal width As Long, ByVal height As Long _
)
'���e:
'   �}���`�~�[�����s��
'
'[srcPlane]    IN   CImgPlane�^:        ���v���[��
'[srcZone]     IN   String�^:           ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^:     ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:        ���ʊi�[�v���[��
'[dstZone]     IN   String�^:           ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^:     ���ʊi�[�v���[���̐F�w��
'[multiMeanFunc] IN IdpMultiMeanFunc�^: ���Z���@�w��(Max,Min,Mean,Sum)
'[Width]       IN   Long�^:             ���w��
'[Height]      IN   Long�^:             �����w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.MultiMean(srcPlane, width, height, multiMeanFunc, dstColor, srcColor)

End Sub

Public Sub MultiMeanByBlock( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal multiMeanFunc As IdpMultiMeanFunc, ByVal DivX As Long, ByVal DivY As Long _
)
'���e:
'   �}���`�~�[�����s��
'
'[srcPlane]    IN   CImgPlane�^:        ���v���[��
'[srcZone]     IN   String�^:           ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^:     ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:        ���ʊi�[�v���[��
'[dstZone]     IN   String�^:           ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^:     ���ʊi�[�v���[���̐F�w��
'[multiMeanFunc] IN IdpMultiMeanFunc�^: ���Z���@�w��(Max,Min,Mean,Sum)
'[DivX]        IN   Long�^:             �������w��
'[DivY]        IN   Long�^:             �c�����w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.MultiMeanByBlock(srcPlane, DivX, DivY, multiMeanFunc, dstColor, srcColor)

End Sub


Public Sub AccumulateRow( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal accOption As IdpAccumOption = idpAccumSum, Optional ByVal dstCol As Long = 1 _
)
'���e:
'   �������ɉ��Z
'
'[srcPlane]    IN   CImgPlane�^:        ���v���[��
'[srcZone]     IN   String�^:           ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^:     ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:        ���ʊi�[�v���[��
'[dstZone]     IN   String�^:           ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^:     ���ʊi�[�v���[���̐F�w��
'[accOption]   IN   IdpAccumOption�^:   ���Z���@�w��(Mean,Sum,StdDeviation)
'[dstCol]      IN   Long�^:             ���ʊi�[��w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.AccumulateRow(srcPlane, accOption, dstCol, dstColor, srcColor)

End Sub

Public Sub AccumulateColumn( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal accOption As IdpAccumOption = idpAccumSum, Optional ByVal dstRow As Long = 1 _
)
'���e:
'   �c�����ɉ��Z
'
'[srcPlane]    IN   CImgPlane�^:        ���v���[��
'[srcZone]     IN   String�^:           ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^:     ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:        ���ʊi�[�v���[��
'[dstZone]     IN   String�^:           ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^:     ���ʊi�[�v���[���̐F�w��
'[accOption]   IN   IdpAccumOption�^:   ���Z���@�w��(Mean,Sum,StdDeviation)
'[dstRow]      IN   Long�^:             ���ʊi�[�s�w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.AccumulateColumn(srcPlane, accOption, dstRow, dstColor, srcColor)

End Sub

Public Sub SubRows( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal diffRows As Long _
)
'���e:
'   �w�蕝���אڂ���s���m�����Z
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[diffRows]    IN   Long�^:         ���w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.SubRows(srcPlane, diffRows, dstColor, srcColor)

End Sub

Public Sub SubColumns( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal diffCols As Long _
)
'���e:
'   ###���̃v���V�[�W���̖����Ȃǂ��ł��邾���ڂ����L�q���Ă�������###
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[diffCols]    IN   Long�^:         ���w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.SubColumns(srcPlane, diffCols, dstColor, srcColor)

End Sub

Public Sub LOr( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal srcBit2 As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'���e:
'   �Ώۉ摜���m��OR���Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[srcBit]      IN   Long�^:         ���v���[��1�̃r�b�g�w��
'[srcBit2]     IN   Long�^:         ���v���[��2�̃r�b�g�w��
'[dstBit]      IN   Long�^:         ���ʊi�[�v���[���̃r�b�g�w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LOr(srcPlane, srcPlane2, dstColor, srcColor, srcColor2, dstBit, srcBit, srcBit2)

End Sub

Public Sub LAnd( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef srcPlane2 As CImgPlane, ByVal srcZone2 As Variant, ByRef srcColor2 As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal srcBit2 As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'���e:
'   �Ώۉ摜���m��AND���Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��1
'[srcZone]     IN   String�^:       ���v���[��1�̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[��1�̐F�w��
'[srcPlane2]   IN   CImgPlane�^:    ���v���[��2
'[srcZone2]    IN   String�^:       ���v���[��2�̃]�[���w��
'[srcColor2]   IN   IdpColorType�^: ���v���[��2�̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[srcBit]      IN   Long�^:         ���v���[��1�̃r�b�g�w��
'[srcBit2]     IN   Long�^:         ���v���[��2�̃r�b�g�w��
'[dstBit]      IN   Long�^:         ���ʊi�[�v���[���̃r�b�g�w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane2.SetPMD(srcZone2)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LAnd(srcPlane, srcPlane2, dstColor, srcColor, srcColor2, dstBit, srcBit, srcBit2)

End Sub

Public Sub ShiftLeft( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal shiftNum As Long _
)
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ShiftLeft(srcPlane, shiftNum, dstColor, srcColor)
End Sub

Public Sub ShiftRight( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal shiftNum As Long _
)
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.ShiftRight(srcPlane, shiftNum, dstColor, srcColor)
End Sub

Public Sub LNot( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    Optional ByVal srcBit As Long = 0, Optional ByVal dstBit As Long = 0 _
)
'���e:
'   �Ώۉ摜��NOT���Z����
'
'[srcPlane]    IN   CImgPlane�^:    ���v���[��
'[srcZone]     IN   String�^:       ���v���[���̃]�[���w��
'[srcColor]    IN   IdpColorType�^: ���v���[���̐F�w��
'[dstPlane]    OUT  CImgPlane�^:    ���ʊi�[�v���[��
'[dstZone]     IN   String�^:       ���ʊi�[�v���[���̃]�[���w��
'[dstColor]    IN   IdpColorType�^: ���ʊi�[�v���[���̐F�w��
'[srcBit]      IN   Long�^:         ���v���[���̃r�b�g�w��
'[dstBit]      IN   Long�^:         ���ʊi�[�v���[���̃r�b�g�w��
'
'���l:
'
    Call srcPlane.SetPMD(srcZone)
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.LNot(srcPlane, dstColor, srcColor, dstBit, srcBit)

End Sub

'=======2009/04/28 Add Maruyama �����ǉ��@���̊֐�==============
'=======2013/02/15 Add JOB�������Ή� �������ύX==============
Public Sub ReadPixel( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal site As Long, _
    ByVal dataNum As Long, ByVal pFlgName As String, _
    ByRef retPixArr() As T_PIXINFO, ByRef AddrMode As IdpAddrMode _
)
'���e:
'   �t���O�v���[���Ŏw�肳�ꂽ��f�̃f�[�^��ǂݍ���
'
'[site]        IN   Long�^:         �T�C�g�w��(�K�{)
'[dataNum]     IN   Long�^:         �ǂݍ��ރf�[�^�̌�
'[srcPlane]    IN   CImgPlane�^:    �f�[�^���v���[��
'[srcZone]     IN   String�^:       �f�[�^���v���[���̃]�[���w��
'[pFlgName]    IN   String�^:       �t���O��
'[retPixArr()] OUT  T_PIXINFO�^:    ���ʊi�[�p�z��
'[AddrMode]    IN   IdpAddrMode�^:  �A�h���X�̕Ԃ���
'
'���l:
'   retPixArr�͕s�蒷�z����w�肷��B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PixelLog(site, pFlgName, retPixArr, dataNum, AddrMode)

End Sub
'=======2009/04/28 Add Maruyama �����܂�==============

Public Sub EdgeCorrect(ByRef dstPlane As CImgPlane, ByRef dstColor As Variant, ByVal srcZoneArray As Variant, ByVal dstZoneArray As Variant)
'���e:
'   srcZoneArray�Ŏw�肳�ꂽ�f�[�^���AdstZoneArray�̑Ή�����]�[���ɃR�s�[����
'
'[dstPlane]     IN  �Ώۃv���[��
'[dstColor]     IN  �Ώۃv���[���̐F�w��
'[srcZoneArray] IN  �R�s�[���]�[���z��
'[dstZoneArray] IN  �R�s�[��]�[���z��
'
'���l:
'
    Dim i As Long
    Dim workPlane As CImgPlane

    Set workPlane = TheIDP.PlaneManager(dstPlane.planeGroup).GetFreePlane(dstPlane.BitDepth)
    For i = 0 To UBound(srcZoneArray)
        Call Copy(dstPlane, srcZoneArray(i), dstColor, workPlane, dstZoneArray(i), dstColor)
        Call Copy(workPlane, dstZoneArray(i), dstColor, dstPlane, dstZoneArray(i), dstColor)
    Next i

End Sub

Public Sub Extention( _
    ByRef pSrcPlane As CImgPlane, ByVal pZone As String, ByRef pDstPlane As CImgPlane, ByVal pExLeft As Long, ByVal pExRight As Long, ByVal pExTop As Long, ByVal pExBottom As Long, _
    Optional ByRef pColor As Variant = EEE_COLOR_FLAT _
)
'���e:
'   pSrcPlane��pZone�̃f�[�^��pDstPlane�ɃR�s�[���A�w�肵�����̕��g������B
'
'[pSrcPlane]    IN  �f�[�^���̃v���[��
'[pZone]        IN  �Ώۂ̃]�[��
'[pDstPlane]    IN  �Ώۂ̃v���[��(pSrcPlane�Ɠ������̂��\)
'[pExLeft]      IN  �����̊g����
'[pExRight]     IN  �E���̊g����
'[pExTop]       IN  �㑤�̊g����
'[pExBottom]    IN  �����̊g����
'[pColor]       IN  �F�w��
'
'���l:
'   ���̒l���w�肷��ƁA�]�[���̓����Ŋg������
'       ��:�]�[���̓����̃f�[�^���O���ɃR�s�[
'       ��:�]�[���̓����ɂ���ɓ�������R�s�[
'   pColor��EEE_COLOR_FLAT�ȊO���w�肷��ƁA�g�����~�J���[�}�b�v�̕��Ɋg�����A�w�肵���F�̃f�[�^���R�s�[����B
'
    Dim tmpSrcPMD As CImgPmdInfo
    Dim tmpDstPMD As CImgPmdInfo
    Dim tmpPlane As CImgPlane

    Set tmpPlane = TheIDP.PlaneManager(pSrcPlane.planeGroup).GetFreePlane(pSrcPlane.BitDepth)

    If pColor <> EEE_COLOR_FLAT Then
        With pSrcPlane.planeMap
            pExLeft = pExLeft * .width
            pExRight = pExRight * .width
            pExTop = pExTop * .height
            pExBottom = pExBottom * .height
        End With
    End If

    '�����R�s�[
    If Not pSrcPlane Is pDstPlane Then
        Call pSrcPlane.SetPMD(pZone)
        Call pDstPlane.SetPMD(pZone)
        Call pDstPlane.CopyPlane(pSrcPlane, pColor)
    End If

    With TheIDP.PMD(pZone)
        '���ӂ��R�s�[
        If pExLeft <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr + (Abs(pExLeft) - pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExLeft), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr - Abs(pExLeft), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '�E�ӂ��R�s�[
        If pExRight <> 0 Then
            Call tmpPlane.SetCustomPMD(.Right + 1 - Abs(pExRight) - (Abs(pExRight) - pExRight) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExRight), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr + Abs(pExRight), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '��ӂ��R�s�[
        If pExTop <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExTop))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr - Abs(pExTop), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If

        '���ӂ��R�s�[
        If pExBottom <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .Bottom + 1 - Abs(pExBottom) - (Abs(pExBottom) - pExBottom) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExBottom))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, pColor)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr + Abs(pExBottom), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane, pColor)
            End With
        End If
    End With

End Sub

Public Sub ExtentionMirror( _
    ByRef pSrcPlane As CImgPlane, ByVal pZone As String, ByRef pDstPlane As CImgPlane, ByVal pExLeft As Long, ByVal pExRight As Long, ByVal pExTop As Long, ByVal pExBottom As Long _
)
'���e:
'   pSrcPlane��pZone�̃f�[�^��pDstPlane�ɃR�s�[���A�w�肵�����̕��g������B
'   �g���͋����R�s�[�ł����B
'
'[pSrcPlane]    IN  �f�[�^���̃v���[��
'[pZone]        IN  �Ώۂ̃]�[��
'[pDstPlane]    IN  �Ώۂ̃v���[��(pSrcPlane�Ɠ������̂��\)
'[pExLeft]      IN  �����̊g����
'[pExRight]     IN  �E���̊g����
'[pExTop]       IN  �㑤�̊g����
'[pExBottom]    IN  �����̊g����
'
'���l:
'   ���̒l���w�肷��ƁA�]�[���̓����Ŋg������
'       ��:�]�[���̓����̃f�[�^���O���ɃR�s�[
'       ��:�]�[���̓����ɂ���ɓ�������R�s�[
'   �F�̎w��͂ł��Ȃ��B�t���b�g�̂݁B
'
    Dim tmpSrcPMD As CImgPmdInfo
    Dim tmpDstPMD As CImgPmdInfo
    Dim tmpPlane As CImgPlane

    Set tmpPlane = TheIDP.PlaneManager(pSrcPlane.planeGroup).GetFreePlane(pSrcPlane.BitDepth)

    '�����R�s�[
    If Not pSrcPlane Is pDstPlane Then
        Call pSrcPlane.SetPMD(pZone)
        Call pDstPlane.SetPMD(pZone)
        Call pDstPlane.CopyPlane(pSrcPlane)
    End If

    With TheIDP.PMD(pZone)
        '���ӂ��R�s�[
        If pExLeft <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr + (Abs(pExLeft) - pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExLeft), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorHorizontal)
                Call pDstPlane.SetCustomPMD(.XAdr - Abs(pExLeft), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '�E�ӂ��R�s�[
        If pExRight <> 0 Then
            Call tmpPlane.SetCustomPMD(.Right + 1 - Abs(pExRight) - (Abs(pExRight) - pExRight) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, Abs(pExRight), .height - (Abs(pExTop) - pExTop) / 2 - (Abs(pExBottom) - pExBottom) / 2)
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorHorizontal)
                Call pDstPlane.SetCustomPMD(.XAdr + Abs(pExRight), .YAdr, .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '��ӂ��R�s�[
        If pExTop <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .YAdr + (Abs(pExTop) - pExTop) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExTop))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorVertical)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr - Abs(pExTop), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If

        '���ӂ��R�s�[
        If pExBottom <> 0 Then
            Call tmpPlane.SetCustomPMD(.XAdr - (Abs(pExLeft) + pExLeft) / 2, .Bottom + 1 - Abs(pExBottom) - (Abs(pExBottom) - pExBottom) / 2, .width + (Abs(pExLeft) + pExLeft) / 2 + (Abs(pExRight) + pExRight) / 2, Abs(pExBottom))
            With tmpPlane.CurrentPMD
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr, .width, .height)
                Call tmpPlane.CopyPlane(pDstPlane, , , idpCopyMirrorVertical)
                Call pDstPlane.SetCustomPMD(.XAdr, .YAdr + Abs(pExBottom), .width, .height)
                Call pDstPlane.CopyPlane(tmpPlane)
            End With
        End If
    End With

End Sub

Private Sub Extention_(ByRef pSrcPlane As CImgPlane, ByRef pDstPlane As CImgPlane, ByRef pWrkPlane As CImgPlane, ByRef pSrcZone As CImgPmdInfo, ByRef pDstZone As CImgPmdInfo)

    Call pSrcPlane.SetPMD(pSrcZone)
    Call pWrkPlane.SetPMD(pSrcZone)
    Call pDstPlane.SetPMD(pDstZone)

    Call pWrkPlane.CopyPlane(pSrcPlane)
    Call pDstPlane.CopyPlane(pWrkPlane)

End Sub

Public Sub GetFreePlane(ByRef pDst As CImgPlane, ByVal pType As String, ByVal pBitDepth As IdpBitDepth, Optional ByVal pClear As Boolean = False, Optional pComment As String = "-")
'���e:
'   �v���[�����擾����B
'
'[pDst]         IN  �Ώۂ̃v���[��
'[pType]        IN  �}�l�[�W�����w��
'[pBitDepth]    IN  �r�b�g�[���w��
'[pClear]       IN  �f�[�^��0�N���A���邩�ǂ���
'[pComment]     IN  �R�����g
'
'���l:
'

    On Error GoTo flagcheck
    
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
'IDP���O�ɏ��o�͂��邽�ߕύX    pDst.Comment = pComment

    Exit Sub

'2014/04/02 T.Akasaka
'Flag�g���@�\�ׁ̈APlane����������Ȃ��ꍇ�͋󂫂�FlagPlane���m�F���鏈����ǉ�
flagcheck:
    Dim ReleaseCount As Long
    
    ReleaseCount = TheIDP.PlaneManager(pType).ReleaseUnusedFlagPlane
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
    
End Sub

'=======2009/04/30 �ǉ� Maruyama ��������==============
Public Function GetFreePlaneForTOPT(ByRef pDst As CImgPlane, ByVal pType As String, ByVal pBitDepth As IdpBitDepth, _
        Optional ByVal pClear As Boolean = False, Optional pComment As String = "-") As Boolean
'���e:
'   �v���[�����擾����B
'
'[pDst]         IN  �Ώۂ̃v���[��
'[pType]        IN  �}�l�[�W�����w��
'[pBitDepth]    IN  �r�b�g�[���w��
'[pClear]       IN  �f�[�^��0�N���A���邩�ǂ���
'[pComment]     IN  �R�����g
'
'���l:
'
    
    Set pDst = Nothing
    On Error GoTo ErrExit
    Set pDst = TheIDP.PlaneManager(pType).GetFreePlane(pBitDepth, pClear, pComment)
'IDP���O�ɏ��o�͂��邽�ߕύX    pDst.Comment = pComment
    
    GetFreePlaneForTOPT = True
    Exit Function
    
ErrExit:
    Dim Err As CErrInfo
    Set Err = TheError.LastError
    TheExec.Datalog.WriteComment pType & " : There is no free plane."
    TheExec.Datalog.WriteComment Err.Message
    GetFreePlaneForTOPT = False
    Exit Function

End Function
'=======2009/04/30 �ǉ� Maruyama �����܂�==============

Public Sub ReleasePlane(ByRef pDst As CImgPlane)
'���e:
'   �v���[�����������B
'
'���l:
'
    Set pDst = Nothing
End Sub

Public Sub GetRegisteredPlane(ByVal pName As String, ByRef pDst As CImgPlane, Optional pComment As String, Optional IsDelete As Boolean = False)
'���e:
'   PlaneBank�ɓo�^���ꂽ�v���[�����擾����B
'
'[pName]        IN  �o�^�������O���w��
'[pDst]         IN  �Ώۂ̃v���[��
'
'���l:
'
    Set pDst = Nothing
    Set pDst = TheIDP.PlaneBank.Item(pName)
    
'=======2009/04/28 Add Maruyama ��������==============
    pDst.Comment = pComment
        
    If IsDelete Then
        TheIDP.PlaneBank.Delete (pName)
        pDst.ReadOnly = False
    End If
'=======2009/04/28 Add Maruyama �����܂�==============
End Sub

'=======2009/04/30 Add Maruyama ��������==============
Public Sub SharedFlagNot(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, _
        ByVal pDstName As String, ByVal pSrcName As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'���e:
'   SharedFlag��Flag�𔽓]����
'
'[pPlaneGroup]         IN Variant�^�@�@�@�@�Ώۂ̃v���[��,�^�C�v
'[pZone]        IN Variant�^:       �Ώۃv���[���̃]�[���w��
'[pSrcName]     IN String�^:        �f�[�^���̖��O
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName��pDstName������ł��\�B
'

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LNot(pDstName, pSrcName, pColor)
    End With
    
End Sub

Public Sub SharedFlagOr(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��Or���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pPlaneGroup]         IN Variant�^�@�@�@�@�Ώۂ̃v���[��,�^�C�v
'[pZone]        IN Variant�^:       �Ώۃv���[���̃]�[���w��
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B
'
    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LOr(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub SharedFlagAnd(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��And���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pPlaneGroup]         IN Variant�^�@�@�@�@�Ώۂ̃v���[��,�^�C�v
'[pZone]        IN Variant�^:       �Ώۃv���[���̃]�[���w��
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LAnd(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub SharedFlagXor(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ByVal pSrcName1 As String, ByVal pSrcName2 As String, _
        Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   pSrcName1�̃r�b�g��pSrcName2�̃r�b�g��XOr���Z�̌��ʂ�pDstName�̃r�b�g�ɓ����B
'   pDstName���o�^����Ă��Ȃ��ꍇ�A�V���ɓo�^����B
'   ���ɓo�^����Ă���ꍇ�́A���̃r�b�g�ɓ����B
'
'[pPlaneGroup]         IN Variant�^�@�@�@�@�Ώۂ̃v���[��,�^�C�v
'[pZone]        IN Variant�^:       �Ώۃv���[���̃]�[���w��
'[pSrcName1]    IN String�^:        �f�[�^���̖��O1
'[pSrcName2]    IN String�^:        �f�[�^���̖��O2
'[pDstName]     IN String�^:        ���ʂ̖��O
'[pColor]       IN IdpColorType�^:  �F�w��
'
'���l:
'   pSrcName1��pSrcName2��pDstName������ł��\�B

    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        Call .LXor(pDstName, pSrcName1, pSrcName2, pColor)
    End With

End Sub

Public Sub FlagNot(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   ���p�t���O�v���[����Not���Z�̌��ʂ��擾
'
'[pDst]         IN   CImgPlane�^�@�@�@�@�Ώۂ̃v���[��
'[pZone]        IN   Variant�^:       �Ώۃv���[���̃]�[���w��
'[pFlgName]     IN   String�^:          �f�[�^���̃t���O��
'[pDstBit]      IN   Long�^:            �ۑ���r�b�g�w��
'[pColor]       IN   IdpColorType�^:    �F�w��
'
'���l:
    With pDst
        Call .SetPMD(pZone)
        Call .FlagNot(pFlgName, pDstBit, pColor)
    End With
End Sub

Public Sub FlagOr(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   ���p�t���O�v���[����Or���Z�̌��ʂ��擾
'
'[pDst]         IN   CImgPlane�^�@�@�@�@�Ώۂ̃v���[��
'[pZone]        IN   Variant�^:         �Ώۃv���[���̃]�[���w��
'[pFlgName1]    IN   String�^:          �f�[�^��1�̃t���O��
'[pFlgName2]    IN   String�^:          �f�[�^��2�̃t���O��
'[pDstBit]      IN   Long�^:            �ۑ���r�b�g�w��
'[pColor]       IN   IdpColorType�^:    �F�w��
'
'���l:
'   pFlgName1�ApFlgName2�͓����̎w��\�B
    With pDst
        Call .SetPMD(pZone)
        Call .FlagOr(pFlgName1, pFlgName2, pDstBit, pColor)
    End With

End Sub

Public Sub FlagAnd(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   ���p�t���O�v���[����And���Z�̌��ʂ��擾
'
'[pDst]         IN   CImgPlane�^�@�@�@�@�Ώۂ̃v���[��
'[pZone]        IN   Variant�^:       �Ώۃv���[���̃]�[���w��
'[pFlgName1]    IN   String�^:          �f�[�^��1�̃t���O��
'[pFlgName2]    IN   String�^:          �f�[�^��2�̃t���O��
'[pDstBit]      IN   Long�^:            �ۑ���r�b�g�w��
'[pColor]       IN   IdpColorType�^:    �F�w��
'
'���l:
'   pFlgName1�ApFlgName2�͓����̎w��\�B
    With pDst
        Call .SetPMD(pZone)
        Call .FlagAnd(pFlgName1, pFlgName2, pDstBit, pColor)
    End With
End Sub

Public Sub FlagXor(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName1 As String, ByVal pFlgName2 As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_FLAT)
'���e:
'   ���p�t���O�v���[����Xor���Z�̌��ʂ��擾
'
'[pDst]         IN   CImgPlane�^�@�@�@�@�Ώۂ̃v���[��
'[pZone]        IN   Variant�^:         �Ώۃv���[���̃]�[���w��
'[pFlgName1]    IN   String�^:          �f�[�^��1�̃t���O��
'[pFlgName2]    IN   String�^:          �f�[�^��2�̃t���O��
'[pDstBit]      IN   Long�^:            �ۑ���r�b�g�w��
'[pColor]       IN   IdpColorType�^:    �F�w��
'
'���l:
'   pFlgName1�ApFlgName2�͓����̎w��\�B
    With pDst
        Call .SetPMD(pZone)
        Call .FlagXor(pFlgName1, pFlgName2, pDstBit, pColor)
    End With

End Sub

Public Sub FlagCopy(ByRef pDst As CImgPlane, ByVal pZone As Variant, ByVal pFlgName As String, _
    Optional ByVal pDstBit As Long = 1, Optional ByRef pColor As Variant = EEE_COLOR_ALL)
'���e:
'   ���p�t���O�v���[���̎w�肵���r�b�g�̒l���R�s�[
'
'[pDst]         IN   CImgPlane�^�@�@�@�@�Ώۂ̃v���[��
'[pZone]        IN   Variant�^:         �Ώۃv���[���̃]�[���w��
'[pFlgName]     IN   String�^:          �f�[�^���̃t���O��
'[pDstBit]      IN   Long�^:            �ۑ���r�b�g�w��
'[pColor]       IN   IdpColorType�^:    �F�w��
    With pDst
        Call .SetPMD(pZone)
        Call .FlagCopy(pFlgName, pDstBit, pColor)
    End With

End Sub
'=======2009/04/30 Add Maruyama �����܂�==============

'=======2009/06/03 Add Maruyama ��������==============
'���コ��쐬
Public Function IsFlgExist( _
        ByRef pPlaneGroup As Variant, ByVal pFlgName As String) As Boolean

    Dim Bit As Long
    
    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pFlgName)
        Bit = .FlagBit(pFlgName)
        If Bit = 0 Then
            IsFlgExist = False
            Exit Function
        End If
    End With
    IsFlgExist = True
    
End Function
'=======2009/06/03 Add Maruyama �����܂�==============

'2009/09/15 D.Maruyama ColorAll�����ǉ� ��������
Public Sub AverageColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   ���ϒl���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AverageColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub SumColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   ���v�l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.SumColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub StdDevColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �W���΍����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.StdDevColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub GetPixelCountColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �Ώۂ̉�f�����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.NumColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MinColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MinColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ő�l���擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub MinMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retMin As CImgColorAllResult, ByRef retMax As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�A�ő�l����x�Ɏ擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retMin]      OUT  CImgColorAllResult�^:       �ŏ��l�i�[�p�z��(�T�C�g��)
'[retMax]      OUT  CImgColorAllResult�^:       �ő�l�i�[�p�z��(�T�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResutMin As ColorAllResult
    Dim tempResutMax As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.MinMaxColorAll(tempResutMin, tempResutMax, pFlgName)

    If retMax Is Nothing Then
        Set retMax = New CImgColorAllResult
    End If
    
    Call retMax.SetParam(srcPlane, tempResutMax)

    If retMin Is Nothing Then
        Set retMin = New CImgColorAllResult
    End If
    
    Call retMin.SetParam(srcPlane, tempResutMin)

End Sub

Public Sub DiffMinMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�ƍő�l�̍����擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CCImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.DiffMinMaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub AbsMaxColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
'���e:
'   �ŏ��l�ƍő�l�̓���Βl�̑傫�������擾
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[retResult]   OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]    IN   String�^:       �t���O��
'
'���l:
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.AbsMaxColorAll(tempResut, pFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub


Public Sub CountColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, _
    ByRef retResult As CImgColorAllResult, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "" _
)
'���e:
'   �����ɊY������_�̌����擾����B
'
'[srcPlane]         IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]          IN   String�^:       �Ώۃv���[���̃]�[���w��
'[countType]        IN   IdpCountType�^: �J�E���g�����w��
'[loLim]            IN   Variant�^:      �����l
'[hiLim]            IN   Variant�^:      ����l
'[pCountLimMode]    IN   IdpCountLimitMode�^: ���E�l�̂Ƃ肩���i�T�C�g�ʁH�T�C�g�F�ʁH�j
'[limitType]        IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[retResult]        OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]         IN   String�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.CountColorAll(tempResut, countType, loLim, hiLim, pCountLimMode, limitType, pFlgName, pInputFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub CountColorAllForFlgBitImgPlane( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, ByRef retResult As CImgColorAllResult, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pInputFlgName As String = "" _
)
'���e:
'   �����ɊY������_�̌����擾����B(�t���O�r�b�g���C���[�W�v�����ɗ��Ă�)
'
'[srcPlane]         IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]          IN   String�^:       �Ώۃv���[���̃]�[���w��
'[countType]        IN   IdpCountType�^: �J�E���g�����w��
'[loLim]            IN   Variant�^:      �����l
'[hiLim]            IN   Variant�^:      ����l
'[pCountLimMode]    IN   IdpCountLimitMode�^: ���E�l�̂Ƃ肩���i�T�C�g�ʁH�T�C�g�F�ʁH�j
'[limitType]        IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[retResult]        OUT  CImgColorAllResult�^:       ���ʊi�[�p�\����(�F�ʃT�C�g��)
'[pFlgName]         IN   CImgPlane�^:       �o�̓t���O��
'[pFlgBit]        �@IN   Long�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'
    Dim tempResut As ColorAllResult
    Call srcPlane.SetPMD(srcZone)
    Call pFlgPlane.SetPMD(srcZone)
    Call srcPlane.CountColorAllForFlgBitImgPlane(tempResut, countType, loLim, hiLim, pCountLimMode, pFlgPlane, pFlgBit, limitType, pInputFlgName)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    Call retResult.SetParam(srcPlane, tempResut)

End Sub

Public Sub PutFlagColorAll( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, _
    ByVal limitType As IdpLimitType, _
    ByVal pFlgName As String, Optional ByVal pInputFlgName As String _
)
'���e:
'   �����ɊY������_�Ƀt���O�𗧂Ă�B(Count����t���O�𗧂Ă邱�Ƃ����ɓ���)
'
'[srcPlane]    IN   CImgPlane�^:    �Ώۃv���[��
'[srcZone]     IN   String�^:       �Ώۃv���[���̃]�[���w��
'[countType]   IN   IdpCountType�^: �J�E���g�����w��
'[loLim]       IN   Variant�^:      �����l
'[hiLim]       IN   Variant�^:      ����l
'[pCountLimMode]    IN   IdpCountLimitMode�^: ���E�l�̂Ƃ肩���i�T�C�g�ʁH�T�C�g�F�ʁH�j
'[limitType]   IN   IdpLimitType�^: ���E�l���܂ށA�܂܂Ȃ��w��
'[pFlgName]         IN   String�^:       �o�̓t���O��
'[pInputFlgName]    IN   String�^:       ���̓t���O��
'
'���l:
'   hiLim,loLim��Variant�Œ�`���Ă���̂́A�T�C�g�ʂŋ��E�l���Ⴄ�ꍇ�ɑΉ����邽�߁B
'   �T�C�g�z�������Ɗe�T�C�g���ƕʁX�ɁA�萔������ƑS�T�C�g�ɓ����l���K�p�����B
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.PutFlagColorAll(pFlgName, countType, loLim, hiLim, pCountLimMode, limitType, pInputFlgName)

End Sub
'2009/09/15 D.Maruyama ColorAll�����ǉ� �����܂�

'�ȉ��̓��W���[�����ł݂̂̎g�p ##################################################################################################################################
Private Sub SetOptionalPMD(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByVal refZone As Variant)

    If dstPlane Is Nothing Then Exit Sub

    If IsEmpty(dstZone) Then
        Call dstPlane.SetPMD(refZone)
    Else
        Call dstPlane.SetPMD(dstZone)
    End If

End Sub

Private Function Var2PlaneNameFlag(ByVal pVal As Variant) As String
    
    If IsObject(pVal) Then
        Var2PlaneNameFlag = pVal.Manager.Name
    Else
        Var2PlaneNameFlag = pVal
    End If
End Function



