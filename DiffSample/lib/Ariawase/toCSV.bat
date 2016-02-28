@if (1==1) /*
@echo off

if "%~2"=="" goto :USAGE
if "%~1"=="/?" goto :USAGE

rem ********************************************************************************
:MAIN
CScript //nologo //E:JScript "%~f0" %*
If ERRORLEVEL 1 goto :USAGE
goto :eof

rem ********************************************************************************
:USAGE
echo USAGE:%~n0 [-s �V�[�g��] ���̓t�@�C�� �o�̓t�@�C��
echo       xls�t�@�C�����J���ACSV(�J���}��؂�)�`���ŕۑ�(SaveAs)���܂��B
echo.
echo  �y�I�v�V�����z
echo     -s �V�[�g��
echo       ���̎w�肪�Ȃ��ꍇ�A1�V�[�g�ڂ�ϊ��ΏۂƂ��܂�
echo.
echo     ���̓t�@�C��
echo       ���̓t�@�C�����w�肵�܂�
echo.
echo     �o�̓t�@�C��
echo       �o�̓t�@�C�����w�肵�܂�
goto :eof

rem ********************************************************************************
rem */
@end
//---------------------------------------------------------- �Z�b�g�A�b�v
var Args = WScript.Arguments;
var EXCEL = WScript.CreateObject("EXCEL.Application");
var SHELL = WScript.CreateObject("WScript.Shell");

function echo(o){ WScript.Echo(o); }

// EXCEL�̒萔
var xlCSV = 6;

//---------------------------------------------------------- ��������
var sheet = null;
var infile = null;
var outfile = null;
for (var i=0; i<Args.Length; i++){
	var p = Args(i);
	switch (p) {
	case "-s":
		sheet = Args(++i);
		break;
	default:
		if (!infile) {
			infile = p;
		} else {
			outfile = p;
		}
		break;
	}
}
if (!infile){
	echo("���̓t�@�C�����w�肳��Ă��܂���");
	WScript.Quit(9);
}
if (!outfile){
	echo("�o�̓t�@�C�����w�肳��Ă��܂���");
	WScript.Quit(9);
}

//---------------------------------------------------------- �又��
// �J�����g�f�B���N�g���̐؂�ւ�
if (EXCEL.DefaultFilePath != SHELL.CurrentDirectory){
	EXCEL.DefaultFilePath = SHELL.CurrentDirectory;
	delete EXCEL;
	EXCEL = WScript.CreateObject("EXCEL.Application");
}

// �t�@�C�����J��
var book = EXCEL.Workbooks.Open(infile);
EXCEL.DisplayAlerts = false;

try{
	// �V�[�g�؂�ւ�
	if (sheet != null){
		book.Worksheets(sheet).Activate();
	}
	// xls�ϊ�
	book.SaveAs(outfile, xlCSV);
} catch(e){
	echo(e.number + ":" + e.description);
} finally {
	book.Close(false);
}