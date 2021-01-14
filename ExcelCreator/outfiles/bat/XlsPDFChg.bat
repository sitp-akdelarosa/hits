set EXEPATH=C:\IISroot\Hits\ExcelCreator\outfiles\bat

REM XPSファイルの移動先（ネットワーク上のサーバ：同一ユーザ/パスワードが前提）
set MOVEPATH=\\NAS-30\KowanG\Doc\Work\HiTS搬入票
set MOVEPATH2=\\NAS-30\KowanG\Doc\Work\HiTS搬入票2

echo ************************************** >>%EXEPATH%\bat.log
echo 処理開始 %date% %time% >>%EXEPATH%\bat.log
c:\windows\SysWOW64\cscript.exe %EXEPATH%\xlspdfchg.vbs %MOVEPATH%  %MOVEPATH2%>>%EXEPATH%\bat.log
echo 処理終了 %date% %time% >>%EXEPATH%\bat.log
echo ************************************** >>%EXEPATH%\bat.log
