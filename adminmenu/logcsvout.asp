<%@Language="VBScript"%>

<!--#include file="./Common/Common.inc"-->
<%
'  �i�ύX�����j
'   2013-09-26   Y.TAKAKUWA   �X�}�[�g�t�H���̃J�E���g��ǉ��B
%>
<%
	'�ϐ��錾
	Dim sDateF,sDateT,sMode,tmpDate
	Dim iCount,iLoop,iFileCnt,iDateCnt,i,j,k,iHdRow,iGSum(),iTSum(),iMTSum(),iRSum
	Dim HDate()
	Dim iTKind,iMTKind
	Dim PageNum(),WkNum(),PageTitle(),SubTitle(),Count()
	Dim MPageNum(),MWkNum(),MPageTitle(),MSubTitle(),MCount()
	'2013-09-30 Y.TAKAKUWA Add-S
	Dim SPageNum(),SWkNum(),SPageTitle(),SSubTitle(),SCount(),iSTKind
	'2013-09-30 Y.TAKAKUWA Add-E
	Dim strTitleFileName,sHdValue,strFileName
	Dim FPageNum(),FWkNum(),FDate(),FCount()

	' Temp�t�@�C�������̃`�F�b�N

	' File System Object �̐���
	Set fs=Server.CreateObject("Scripting.FileSystemobject")

	' �_�E�����[�h�t�@�C���̎擾

	strFileName = Session.Contents("tempfile")
	If strFileName="" Then
		' �Z�b�V�������؂�Ă���Ƃ�
		Response.Redirect("accesstotal.asp")	 '���p����Top��
		Response.End
	End If
	strFileName="../temp/" & strFileName
	' �_�E�����[�h�t�@�C����Open
	Set ti=fs.OpenTextFile(Server.MapPath(strFileName),1,True)

	' �t�@�C���̃_�E�����[�h
	Response.ContentType="application/octet-stream"
	Response.AddHeader "Content-Disposition","attachment; filename=output.csv"

	'�w�b�_�̎擾

	anyTmp=Split(ti.ReadLine,",")
	sDateF=anyTmp(0)
	sDateT=anyTmp(1)
	sMode=anyTmp(2)
	
	iLoop=0
	iFileCnt=0
	'CSV�t�@�C���f�[�^��ϐ��Ɋi�[
	Do While Not ti.AtEndOfStream
		if iLoop<>0 then
			anyTmp=Split(ti.ReadLine,",")
			ReDim Preserve FPageNum(iFileCnt)
			ReDim Preserve FWkNum(iFileCnt)
			ReDim Preserve FDate(iFileCnt)
			ReDim Preserve FCount(iFileCnt)
			FPageNum(iFileCnt)=anyTmp(0)
			FWkNum(iFileCnt)=anyTmp(1)
			if sMode="D" then
				FDate(iFileCnt)=left(anyTmp(2),4) & "/" & mid(anyTmp(2),5,2) & "/" & right(anyTmp(2),2)
			else
				FDate(iFileCnt)=left(anyTmp(2),4) & "/" & mid(anyTmp(2),5,2) 
			end if
			FCount(iFileCnt)=anyTmp(3)
			iFileCnt=iFileCnt+1
		end if
		iLoop=iLoop+1
	Loop

	'�w�b�_��������
	if sMode="D" then
			Response.Write left(sDateF,4) & "�N" & mid(sDateF,5,2) & "��" & right(sDateF,2) & "������" & left(sDateT,4) & "�N" & mid(sDateT,5,2) & "��" & right(sDateT,2) & "���܂�,"
	Else
			Response.Write left(sDateF,4) & "�N" & mid(sDateF,5,2) & "������" & left(sDateT,4) & "�N" & mid(sDateT,5,2) & "���܂�,"
	End If
	Response.Write Chr(13) & Chr(10)
	Response.Write "���p�\�R����,"
	Response.Write Chr(13) & Chr(10)

	'�p�\�R���^�C�g���s�o��
	Response.Write "���j���[����,"
	Response.Write "���,"
	Response.Write "���No," 
	
	iDateCnt=0
	iLoop=0
	'�w�b�_���t�ݒ�
	'���ʂ̏ꍇ
	if sMode="D" then
		tmpDate=left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2)
		iCount=DateDiff("d", left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2), left(sDateT,4) & "/" & mid(sDateT,5,2) & "/" & right(sDateT,2))
		
		do 
			if iCount<iLoop then
				exit do
			end if
			ReDim Preserve HDate(iDateCnt)
			HDate(iDateCnt)=tmpDate
			Response.Write tmpDate & ","
			tmpDate=DateAdd("d", 1, tmpDate)
			iDateCnt=iDateCnt+1
			iLoop=iLoop+1
		Loop  
	else
		iCount=DateDiff("m", left(sDateF,4) & "/" & mid(sDateF,5,2) & "/01", left(sDateT,4) & "/" & mid(sDateT,5,2) & "/01")

		tmpDate=left(sDateF,4) & "/" & mid(sDateF,5,2) & "/" & right(sDateF,2)
		do 
			if iCount<iLoop then
				exit do
			end if
			ReDim Preserve HDate(iDateCnt)
			HDate(iDateCnt)=left(tmpDate,7)
			Response.Write left(tmpDate,7) & ","
			tmpDate=DateAdd("m", 1, tmpDate)
			iDateCnt=iDateCnt+1
			iLoop=iLoop+1
		Loop

	end if
	Response.Write "���v"

	Response.Write Chr(13) & Chr(10)

	'�p�\�R���p ���O�^�C�g���擾
	strTitleFileName="../logweb.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iTKind=0
	
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2) <> "" Then 
			ReDim Preserve PageNum(iTKind)
			ReDim Preserve WkNum(iTKind)
			ReDim Preserve PageTitle(iTKind)
			ReDim Preserve SubTitle(iTKind)
			PageTitle(iTKind) = anyTmpTitle(2)
			PageNum(iTKind) = ""
			WkNum(iTKind) = ""
			SubTitle(iTKind) = ""
			iTKind=iTKind+1
		end if
		ReDim Preserve PageNum(iTKind)
		ReDim Preserve WkNum(iTKind)
		ReDim Preserve PageTitle(iTKind)
		ReDim Preserve SubTitle(iTKind)
		PageNum(iTKind) = anyTmpTitle(0)
		WkNum(iTKind) = anyTmpTitle(1)
		PageTitle(iTKind) = ""
		SubTitle(iTKind) = anyTmpTitle(3)
		iTKind=iTKind+1
	Loop
	ti.Close

	'�g�їp ���O�^�C�g���擾
	strTitleFileName="../logija.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iMTKind=0
	
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2)<>"" Then 
			ReDim Preserve MPageNum(iMTKind)
			ReDim Preserve MWkNum(iMTKind)
			ReDim Preserve MPageTitle(iMTKind)
			ReDim Preserve MSubTitle(iMTKind)
			MPageTitle(iMTKind) = anyTmpTitle(2)
			MPageNum(iMTKind) = ""
			MWkNum(iMTKind) = ""
			'2017/10/13 Add-S CIS
			'SubTitle(iMTKind) = ""
			MSubTitle(iMTKind) = ""
			'2017/10/13 Add-E CIS
			iMTKind=iMTKind+1
		end if
		ReDim Preserve MPageNum(iMTKind)
		ReDim Preserve MWkNum(iMTKind)
		ReDim Preserve MPageTitle(iMTKind)
		ReDim Preserve MSubTitle(iMTKind)
		MPageNum(iMTKind) = anyTmpTitle(0)
		MWkNum(iMTKind) = anyTmpTitle(1)
		MPageTitle(iMTKind) = ""
		MSubTitle(iMTKind) = anyTmpTitle(3)
		iMTKind=iMTKind+1
	Loop
	ti.Close
	
	'2013-09-30 Y.TAKAKUWA Add-S
	'�g�їp ���O�^�C�g���擾
	strTitleFileName="../logsumafo.txt"
	Set ti=fs.OpenTextFile(Server.MapPath(strTitleFileName),1,True)
	iSTKind=0
	
	Do While Not ti.AtEndOfStream
		strTemp=ti.ReadLine
		anyTmpTitle=Split(strTemp,",")
		If anyTmpTitle(2)<>"" Then 
			ReDim Preserve SPageNum(iSTKind)
			ReDim Preserve SWkNum(iSTKind)
			ReDim Preserve SPageTitle(iSTKind)
			ReDim Preserve SSubTitle(iSTKind)
			SPageTitle(iSTKind) = anyTmpTitle(2)
			SPageNum(iSTKind) = ""
			SWkNum(iSTKind) = ""
			'2017/10/13 Add-S CIS
			'SubTitle(iSTKind) = ""
			SSubTitle(iSTKind) = ""
			'2017/10/13 Add-E CIS
			iSTKind=iSTKind+1
		end if
		ReDim Preserve SPageNum(iSTKind)
		ReDim Preserve SWkNum(iSTKind)
		ReDim Preserve SPageTitle(iSTKind)
		ReDim Preserve SSubTitle(iSTKind)
		SPageNum(iSTKind) = anyTmpTitle(0)
		SWkNum(iSTKind) = anyTmpTitle(1)
		SPageTitle(iSTKind) = ""
		SSubTitle(iSTKind) = anyTmpTitle(3)
		iSTKind=iSTKind+1
	Loop
	ti.Close
	'2013-09-30 Y.TAKAKUWA Add-E

	'�p�\�R���p�f�[�^����
	ReDim Count(iTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iTSum(iLoop)=0
	Next
	'�p�\�R���\�����ڕ����[�v
	For i=0 to iTKind-1
		'���j���[���ڂ��ς�����ꍇ
		if sHdValue<>PageTitle(i) and trim(PageTitle(i))<>"" then
			'�擪�s�ȊO
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					Count(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=PageTitle(i)
		end if
		'�J�E���g�N���A
		For iLoop=0 to iDateCnt-1
			Count(i,iLoop)=0
		Next
		
		'�t�@�C���s�������[�v
		For j=0 to iFileCnt-1
			'��ʔԍ��A��Ɣԍ��������ꍇ
			If PageNum(i)=FPageNum(j) and WkNum(i)=FWkNum(j) then
				'���t�����[�v
				For k=0 to iDateCnt-1
					'���t�������f�[�^�̏ꍇ
					if cstr(HDate(k))=cstr(FDate(j)) then
						Count(i,k)=Count(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iTSum(k)=iTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'�ŏI�s�̃f�[�^�𑫂�����
	For iLoop=0 to iDateCnt-1
		Count(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '-------------------------------------------------------------------------------------------------------
	'�g�їp�f�[�^����
	sHdValue=""
	ReDim MCount(iMTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iMTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iMTSum(iLoop)=0
	Next
	'�g�ѕ\�����ڕ����[�v
	For i=0 to iMTKind-1
		'���j���[���ڂ��ς�����ꍇ
		if sHdValue<>MPageTitle(i) and trim(MPageTitle(i))<>"" then
			'�擪�s�ȊO
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					MCount(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=MPageTitle(i)
		end if
		'�J�E���g�N���A
		For iLoop=0 to iDateCnt-1
			MCount(i,iLoop)=0
		Next
		
		'�t�@�C���s�������[�v
		For j=0 to iFileCnt-1
			'��ʔԍ��A��Ɣԍ��������ꍇ
			If MPageNum(i)=FPageNum(j) and MWkNum(i)=FWkNum(j) then
				'���t�����[�v
				For k=0 to iDateCnt-1
					'���t�������f�[�^�̏ꍇ
					if cstr(HDate(k))=cstr(FDate(j)) then
						MCount(i,k)=MCount(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iMTSum(k)=iMTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'�ŏI�s�̃f�[�^�𑫂�����
	For iLoop=0 to iDateCnt-1
		MCount(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-S
    '------------------------------------------------------------------------------------------------------------
    '�X�}�g�t�H���f�[�^����
	sHdValue=""
	ReDim SCount(iSTKind-1,iDateCnt-1)
	ReDim iGSum(iDateCnt-1)
	ReDim iSTSum(iDateCnt-1)
	for iLoop=0 to iDateCnt-1
		iSTSum(iLoop)=0
	Next
	'�g�ѕ\�����ڕ����[�v
	For i=0 to iSTKind-1
		'���j���[���ڂ��ς�����ꍇ
		if sHdValue<>SPageTitle(i) and trim(SPageTitle(i))<>"" then
			'�擪�s�ȊO
			if i<>0 then
				for iLoop=0 to iDateCnt-1
					SCount(iHdRow,iLoop)=iGSum(iLoop)
				Next
			end if

			iHdRow=i
			for iLoop=0 to iDateCnt-1
				iGSum(iLoop)=0
			Next
			sHdValue=SPageTitle(i)
		end if
		'�J�E���g�N���A
		For iLoop=0 to iDateCnt-1
			SCount(i,iLoop)=0
		Next
		
		'�t�@�C���s�������[�v
		For j=0 to iFileCnt-1
			'��ʔԍ��A��Ɣԍ��������ꍇ
			If SPageNum(i)=FPageNum(j) and SWkNum(i)=FWkNum(j) then
				'���t�����[�v
				For k=0 to iDateCnt-1
					'���t�������f�[�^�̏ꍇ
					if cstr(HDate(k))=cstr(FDate(j)) then
						SCount(i,k)=SCount(i,k)+FCount(j)
						iGSum(k)=iGSum(k)+FCount(j)
						iSTSum(k)=iSTSum(k)+FCount(j)
						Exit for
					end if
				Next
			end if
		Next
	Next
	'�ŏI�s�̃f�[�^�𑫂�����
	For iLoop=0 to iDateCnt-1
		SCount(iHdRow,iLoop)=iGSum(iLoop)
	Next	
    '------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-E
    '------------------------------------------------------------------------------------------------------------
    
	'�p�\�R���t�@�C���֏o��
	For iLoop=0 to iTKind-1
		Response.Write  PageTitle(iLoop) &","
		Response.Write  SubTitle(iLoop) &","
		if trim(PageNum(iLoop))<>"" then
			Response.Write  PageNum(iLoop) & "-" & WkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'���t�����[�v
		for j=0 to iDateCnt-1
			Response.Write  Count(iLoop,j) &","
			iRSum=iRSum+Count(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'���v��������
	Response.Write  "���v,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iTSum(j) &","
		iRSum=iRSum+iTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)

	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
	'�g�уt�@�C���o��
	Response.Write "���g�ѓd�b��,"
	Response.Write Chr(13) & Chr(10)

	'�g�у^�C�g���s�o��
	Response.Write "���j���[����,"
	Response.Write "���,"
	Response.Write "���No," 
	
	For iLoop=0 to iDateCnt-1
		Response.Write HDate(iLoop) & ","
	next
	Response.Write "���v"

	Response.Write Chr(13) & Chr(10)

	'�p�\�R���t�@�C���֏o��
	For iLoop=0 to iMTKind-1
		Response.Write  MPageTitle(iLoop) &","
		Response.Write  MSubTitle(iLoop) &","
		if trim(MPageNum(iLoop))<>"" then
			Response.Write  MPageNum(iLoop) & "-" & MWkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'���t�����[�v
		for j=0 to iDateCnt-1
			Response.Write  MCount(iLoop,j) &","
			iRSum=iRSum+MCount(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'���v��������
	Response.Write  "���v,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iMTSum(j) &","
		iRSum=iRSum+iMTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-S
    '-------------------------------------------------------------------------------------------------------------
    '�X�}�[�g�t�H���t�@�C���o��
	Response.Write "���X�}�[�g�t�H����,"
	Response.Write Chr(13) & Chr(10)

	'�g�у^�C�g���s�o��
	Response.Write "���j���[����,"
	Response.Write "���,"
	Response.Write "���No," 
	
	For iLoop=0 to iDateCnt-1
		Response.Write HDate(iLoop) & ","
	next
	Response.Write "���v"

	Response.Write Chr(13) & Chr(10)

	'�p�\�R���t�@�C���֏o��
	For iLoop=0 to iSTKind-1
		Response.Write  SPageTitle(iLoop) &","
		Response.Write  SSubTitle(iLoop) &","
		if trim(SPageNum(iLoop))<>"" then
			Response.Write  SPageNum(iLoop) & "-" & SWkNum(iLoop) &","		
		else
			Response.Write ","
		end if
		iRSum=0
		'���t�����[�v
		for j=0 to iDateCnt-1
			Response.Write  SCount(iLoop,j) &","
			iRSum=iRSum+SCount(iLoop,j)
		next
		Response.Write iRSum & ","
		Response.Write Chr(13) & Chr(10)
	Next

	'���v��������
	Response.Write  "���v,,,"
	iRSum=0
	for j=0 to iDateCnt-1
		Response.Write  iSTSum(j) &","
		iRSum=iRSum+iSTSum(j)
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)
	Response.Write Chr(13) & Chr(10)
    '-------------------------------------------------------------------------------------------------------------
    ' 2013-09-30 Y.TAKAKUWA Add-E
    '-------------------------------------------------------------------------------------------------------------
    
	'�����v��������
	Response.Write  "�����v,,,"
	iRSum=0
	for j=0 to iDateCnt-1
	    '2013-09-30 Y.TAKAKUWA Upd-S
		'Response.Write  iTSum(j)+iMTSum(j) &","
		'iRSum=iRSum+iTSum(j)+iMTSum(j)
		Response.Write  iTSum(j)+iMTSum(j)+iSTSum(j) &","
		iRSum=iRSum+iTSum(j)+iMTSum(j)+iSTSum(j)		
		'2013-09-30 Y.TAKAKUWA Upd-E
	next
	Response.Write iRSum & ","
	Response.Write Chr(13) & Chr(10)


	' �_�E�����[�h�I��
	Response.End

%>
