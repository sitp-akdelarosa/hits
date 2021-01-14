Option Explicit
On Error Resume Next

Const queDir = "C:\IISroot\Hits\ExcelCreator\outfiles\que"
Const logfile = "C:\IISroot\Hits\ExcelCreator\outfiles\bat\xlspdfchg.log"
Const cLogsize=1000000

Dim objExcelApp, objWbk1, objParm, xlsfile, xpsfile, movepath, xpsfilenames, xpsfilename
Dim fso, dir, file, quefile, inputFile, quetxt, tf
Dim wshNetwork
Dim movepath2				'2016/12/02 H.Yoshikawa Add

	'�t�@�C������
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'���O�t�@�C���I�[�v��
	set tf = fso.OpenTextFile(logfile, 8, True, False)
	'�G���[�i���ɋN�����j�̏ꍇ�I��
	If Err.Number <> 0 Then
		' �X�N���v�g�I��
		Wscript.Quit(-1)
	end if
	gfputTrace "********** �����J�n **********"

	'�p�����[�^�擾
	if WScript.Arguments.Count < 2 then					'2016/12/02 H.Yoshikawa Upd(�����̐���ύX 1��2)
		gfputTrace "�p�����[�^������������܂���"
		' �X�N���v�g�I��
		Wscript.Quit(-1)
	end if
	movepath = WScript.Arguments(0)			'XPS�t�@�C���̈ړ���F�l�b�g���[�N��̃T�[�o�i���ꃆ�[�U/�p�X���[�h���O��j
	movepath2 = WScript.Arguments(1)		'XPS�t�@�C���̈ړ���F�l�b�g���[�N��̃T�[�o�i���ꃆ�[�U/�p�X���[�h���O��j2016/12/02 H.Yoshikawa Add

	' Excel�̃I�u�W�F�N�g�̎Q�Ƃ��擾
	Set objExcelApp = CreateObject("Excel.Application")
	If Err.Number <> 0 Then
		gfputTrace "XPS�쐬�G���[1�F" & Err.Description
		' �X�N���v�g�I��
		Wscript.Quit(-1)
	end if

	Set dir = fso.getFolder(queDir)
	For Each file In dir.Files
	    quefile = file.Name
		Set inputFile = fso.OpenTextFile(queDir & "\" & quefile, 1, False, 0)
		quetxt = inputFile.ReadLine
		objParm = Split(quetxt, "/")
		if Ubound(objParm) < 1 Then
			gfputTrace "XPS�쐬�G���[2�F����������������܂���B(" & quefile & ")"
			' �X�N���v�g�I��
			Wscript.Quit(-1)
		end if
		xlsfile = objParm(0)
		xpsfile = objParm(1)
		inputFile.Close
		Set inputFile = Nothing
		
		gfputTrace "   XPS�ϊ��J�n�F(" & xlsfile & " �� " & xpsfile & ")"

		' Excel�E�B���h�E���\��
		objExcelApp.Visible = false
		'Excel�I�[�v��
		Set objWbk1 = objExcelApp.Workbooks.Open(xlsfile, False, True)
		If Err.Number <> 0 Then
			gfputTrace "XPS�쐬�G���[3�F" & Err.Description
		else
			'XPS�ۑ�
			Call objWbk1.ExportAsFixedFormat(1, xpsfile)
			If Err.Number <> 0 Then
				gfputTrace "XPS�쐬�G���[4�F" & Err.Description
			else
				'XPS�쐬�����Ȃ�AQUE�t�@�C�����폜
				fso.DeleteFile queDir & "\" & quefile, True
				
				'XPS�t�@�C�����ړ�
				xpsfilenames = Split(xpsfile, "\")
				xpsfilename = xpsfilenames(UBound(xpsfilenames))
				fso.copyFile xpsfile, movepath & "\" & xpsfilename, true
				If Err.Number <> 0 Then
					gfputTrace "XPS�ړ��G���[�F" & Err.Description
				End If
				
				'2016/12/02 H.Yoshikawa Add Start
				fso.copyFile xpsfile, movepath2 & "\" & xpsfilename, true
				If Err.Number <> 0 Then
					gfputTrace "XPS�ړ��G���[2�F" & Err.Description
				End If
				'2016/12/02 H.Yoshikawa Add End
			End If
		end if
		err.clear
		
		' �w��u�b�N�����
		objWbk1.Saved = True
		objWbk1.Close False
		Set objWbk1 = Nothing
		
		
	Next

	' Excel�I��
	'objExcelApp.Quit
	Set objExcelApp = Nothing
	Set dir = Nothing

	gfputTrace "********** �����I�� **********"

	'���O�t�@�C���N���[�Y
	tf.Close
	set tf=Nothing

	'���O�t�@�C���o�b�N�A�b�v
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
