set EXEPATH=C:\IISroot\Hits\ExcelCreator\outfiles\bat

REM XPS�t�@�C���̈ړ���i�l�b�g���[�N��̃T�[�o�F���ꃆ�[�U/�p�X���[�h���O��j
set MOVEPATH=\\NAS-30\KowanG\Doc\Work\HiTS�����[
set MOVEPATH2=\\NAS-30\KowanG\Doc\Work\HiTS�����[2

echo ************************************** >>%EXEPATH%\bat.log
echo �����J�n %date% %time% >>%EXEPATH%\bat.log
c:\windows\SysWOW64\cscript.exe %EXEPATH%\xlspdfchg.vbs %MOVEPATH%  %MOVEPATH2%>>%EXEPATH%\bat.log
echo �����I�� %date% %time% >>%EXEPATH%\bat.log
echo ************************************** >>%EXEPATH%\bat.log
