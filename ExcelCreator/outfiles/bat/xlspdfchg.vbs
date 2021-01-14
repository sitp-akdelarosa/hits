Option Explicit
On Error Resume Next

Const queDir = "C:\IISroot\Hits\ExcelCreator\outfiles\que"
Const logfile = "C:\IISroot\Hits\ExcelCreator\outfiles\bat\xlspdfchg.log"
Const cLogsize=1000000

Dim objExcelApp, objWbk1, objParm, xlsfile, xpsfile, movepath, xpsfilenames, xpsfilename
Dim fso, dir, file, quefile, inputFile, quetxt, tf
Dim wshNetwork
Dim movepath2				'2016/12/02 H.Yoshikawa Add

	'ファイル操作
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'ログファイルオープン
	set tf = fso.OpenTextFile(logfile, 8, True, False)
	'エラー（既に起動中）の場合終了
	If Err.Number <> 0 Then
		' スクリプト終了
		Wscript.Quit(-1)
	end if
	gfputTrace "********** 処理開始 **********"

	'パラメータ取得
	if WScript.Arguments.Count < 2 then					'2016/12/02 H.Yoshikawa Upd(引数の数を変更 1→2)
		gfputTrace "パラメータが正しくありません"
		' スクリプト終了
		Wscript.Quit(-1)
	end if
	movepath = WScript.Arguments(0)			'XPSファイルの移動先：ネットワーク上のサーバ（同一ユーザ/パスワードが前提）
	movepath2 = WScript.Arguments(1)		'XPSファイルの移動先：ネットワーク上のサーバ（同一ユーザ/パスワードが前提）2016/12/02 H.Yoshikawa Add

	' Excelのオブジェクトの参照を取得
	Set objExcelApp = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		gfputTrace "XPS作成エラー1：" & Err.Description
		' スクリプト終了
		Wscript.Quit(-1)
	end if

	Set dir = fso.getFolder(queDir)
	For Each file In dir.Files
	    quefile = file.Name
		Set inputFile = fso.OpenTextFile(queDir & "\" & quefile, 1, False, 0)
		quetxt = inputFile.ReadLine
		objParm = Split(quetxt, "/")
		if Ubound(objParm) < 1 Then
			gfputTrace "XPS作成エラー2：引数が正しくありません。(" & quefile & ")"
			' スクリプト終了
			Wscript.Quit(-1)
		end if
		xlsfile = objParm(0)
		xpsfile = objParm(1)
		inputFile.Close
		Set inputFile = Nothing
		
		gfputTrace "   XPS変換開始：(" & xlsfile & " ⇒ " & xpsfile & ")"

		' Excelウィンドウを非表示
		objExcelApp.Visible = false
		'Excelオープン
		Set objWbk1 = objExcelApp.Workbooks.Open(xlsfile, False, True)
		If Err.Number <> 0 Then
			gfputTrace "XPS作成エラー3：" & Err.Description
		else
			'XPS保存
			Call objWbk1.ExportAsFixedFormat(1, xpsfile)
			If Err.Number <> 0 Then
				gfputTrace "XPS作成エラー4：" & Err.Description
			else
				'XPS作成成功なら、QUEファイルを削除
				fso.DeleteFile queDir & "\" & quefile, True
				
				'XPSファイルを移動
				xpsfilenames = Split(xpsfile, "\")
				xpsfilename = xpsfilenames(UBound(xpsfilenames))
				fso.copyFile xpsfile, movepath & "\" & xpsfilename, true
				If Err.Number <> 0 Then
					gfputTrace "XPS移動エラー：" & Err.Description
				End If
				
				'2016/12/02 H.Yoshikawa Add Start
				fso.copyFile xpsfile, movepath2 & "\" & xpsfilename, true
				If Err.Number <> 0 Then
					gfputTrace "XPS移動エラー2：" & Err.Description
				End If
				'2016/12/02 H.Yoshikawa Add End
			End If
		end if
		err.clear
		
		' 指定ブックを閉じる
		objWbk1.Saved = True
		objWbk1.Close False
		Set objWbk1 = Nothing
		
		
	Next

	' Excel終了
	'objExcelApp.Quit
	Set objExcelApp = Nothing
	Set dir = Nothing

	gfputTrace "********** 処理終了 **********"

	'ログファイルクローズ
	tf.Close
	set tf=Nothing

	'ログファイルバックアップ
	dim fi, sz
	set fi=fso.getfile(logfile)
	sz = fi.size
	if cLogsize < sz then
		fso.copyFile logfile, logfile & "_bk" & Replace(Left(FormatDateTime(Now, 2), 10), "/", ""), true
		fso.deletefile logfile
	end if
	set fi = nothing
	Set fso = Nothing


function gfputTrace(str)
'On Error Resume Next
	dim logtime

	logtime=trim(year(now)*10000 + month(now)*100 + day(now)) & mid(trim(1000000 + hour(now)*10000 + minute(now)*100 + second(now)),2)
	tf.WriteLine logtime & ":" & str

end function
