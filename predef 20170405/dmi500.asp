<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi500.asp									   _/
'_/	Function	:��Ɣ���mail�������M						   _/
'_/	Date		:2009/03/11									   _/
'_/	Code By		:Shibuta									   _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''Microsoft ADO�p��adovbs.inc�ɂĒ񋟂���Ă���
	Const adBoolean = 11
	Const adDBTimeStamp = 135
	Const adInteger = 3
	Const adChar = 129
	Const adParamInput = &H0001
	Const adParamReturnValue = &H0004
	Dim ErrCode
	
	ErrCode = 0
	
	'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH
	Session.Contents("SendMailSubmitted") = "False"
	'''���M���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("SendMailSubmitted") = "False" then

		'''�f�[�^�擾
		Dim USER, CALLPG, SENDUSER
		Dim Email1, Email2, Email3, Email4, Email5
		Dim UserName,ComInterval,rc
		
		USER = Session.Contents("userid")
		CALLPG = Session.Contents("callpg")
		SENDUSER = Session.Contents("senduser")

		'''DB�ڑ�
		Dim ObjConn, ObjRS, StrSQL
		ConnDBH ObjConn, ObjRS

		'''�ʐM�Ԋu�擾
		StrSQL = "SELECT ComInterval FROM mParam "

		ObjRS.Open StrSQL, ObjConn
		if err <> 0 then
			'''DB�ؒf
			DisConnDBH ObjConn, ObjRS
			jumpErrorP "2","c104","01","��Ɣ���mail�������M","101","SQL:<BR>"&strSQL
		end if

		ComInterval = ObjRS("ComInterval")
		ObjRS.Close

		if SENDUSER <> "" then
		''��Ɣ����z�M���̎擾
			StrSQL = "SELECT T.*, "
			StrSQL = StrSQL & "CASE WHEN U.NameAbrev IS NULL THEN U.FullName ELSE U.NameAbrev END AS USERNAME "
			StrSQL = StrSQL & "FROM mUsers U, "
			StrSQL = StrSQL & "(SELECT T.* FROM TargetOperation T, mUsers U WHERE T.UserCode = U.UserCode "
			StrSQL = StrSQL & "AND U.HeadCompanyCode =" & SENDUSER & ") T "
			StrSQL = StrSQL & "WHERE U.UserCode = '" & USER & "'"
			
			ObjRS.Open StrSQL, ObjConn
			if err <> 0 then
			'''DB�ؒf
				DisConnDBH ObjConn, ObjRS
				jumpErrorP "2","c104","01","��Ɣ���mail�������M","101","SQL:<BR>"&strSQL
			end if
			'ObjRS.close

			Dim svName, mailTo, mailFrom, attachedFiles, ObjMail
			Dim mailFlag1, mailFlag2, mailFlag3, mailFlag4, mailFlag5
			Dim mailSubject, mailBody,WorkName
			Dim SendTime, UpdateSendTime
		
			'''SMTP�T�[�o���̐ݒ�
			svName   = "slitdns2.hits-h.com"
			'svName = "192.168.17.61"
			attachedFiles = ""
			mailFlag1 = 0
			mailFlag2 = 0
			mailFlag3 = 0
			mailFlag4 = 0
			mailFlag5 = 0
			'''���[�����M���A�h���X�̐ݒ�
			mailFrom = "mrhits@hits-h.com"
'			mailFrom = "test@192.168.17.61"
			mailTo = ""

		Select Case CALLPG
			'''�����o���
			Case "dmi040"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagDelResults1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagDelResults2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagDelResults3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagDelResults4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagDelResults5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "�����o���"
				SendTime = ObjRS("DelResultsDate")
				UpdateSendTime = "DelResultsDate"

			'''��������
			Case "dmi140"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagRecEmp1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if
				
				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagRecEmp2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagRecEmp3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagRecEmp4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagRecEmp5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "��������"
				SendTime = ObjRS("RecEmpDate")
				UpdateSendTime = "RecEmpDate"

			'''���������
			Case "dmi340"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagRecResults1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagRecResults2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagRecResults3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagRecResults4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagRecResults5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "���������"
				SendTime = ObjRS("RecResultsDate")
				UpdateSendTime = "RecResultsDate"

			'''����o���
			Case "dmi240"
				if IsNULL(ObjRS("Email1")) = false AND ObjRS("FlagDelEmp1") = "1" then
					mailTo = mailTo & Trim(ObjRS("Email1"))
					mailFlag1 = 1
				else
					mailFlag1 = 0
				end if

				if IsNULL(ObjRS("Email2")) = false AND ObjRS("FlagDelEmp2") = "1" then
					if mailFlag1 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email1"))
					else
						mailTo = mailTo & Trim(ObjRS("Email1"))
					end if
					mailFlag2 = 1
				else
					mailFlag2 = 0
				end if

				if IsNULL(ObjRS("Email3")) = false AND ObjRS("FlagDelEmp3") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email3"))
					else
						mailTo = mailTo & Trim(ObjRS("Email3"))
					end if
					mailFlag3 = 1
				else
					mailFlag3 = 0
				end if

				if IsNULL(ObjRS("Email4")) = false AND ObjRS("FlagDelEmp4") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email4"))
					else
						mailTo = mailTo & Trim(ObjRS("Email4"))
					end if
					mailFlag4 = 1
				else
					mailFlag4 = 0
				end if

				if IsNULL(ObjRS("Email5")) = false AND ObjRS("FlagDelEmp5") = "1" then
					if mailFlag1 = 1 or mailFlag2 = 1 or mailFlag3 = 1 or mailFlag4 = 1 then
						mailTo = mailTo & vbtab & Trim(ObjRS("Email5"))
					else
						mailTo = mailTo & Trim(ObjRS("Email5"))
					end if
					mailFlag5 = 1
				else
					mailFlag5 = 0
				end if

				WorkName = "����o���"
				SendTime = ObjRS("DelEmpDate")
				UpdateSendTime = "DelEmpDate"
			End Select
			
			Set ObjMail = Server.CreateObject("BASP21")

			mailSubject = "HiTS ��ƈ˗�"
			mailBody = WorkName & "���� (" & Trim(ObjRS("USERNAME")) & "�l���)" & vbCrLf & vbCrLf
			mailBody = mailBody & " (" & Trim(ObjRS("USERNAME")) & "�l���)" & vbCrLf & vbCrLf
			mailBody = mailBody & WorkName & "���������܂����B" & vbCrLf
			mailBody = mailBody & "�ڂ�����HiTS�̎��O���o�^�̉�ʂ����Q�Ɖ������B"

			'���[�����M�������猻�݂̎������ʐM�Ԋu�ȏ�̏ꍇ�̓��[���𑗐M����B
			if Trim(mailTo) <> "" Then
				WriteLogH "c104", "��Ɣ���mail�������M","svName",svName
				WriteLogH "c104", "��Ɣ���mail�������M","mailTo",mailTo
				WriteLogH "c104", "��Ɣ���mail�������M","mailFrom",mailFrom
				WriteLogH "c104", "��Ɣ���mail�������M","mailSubject",mailSubject
				WriteLogH "c104", "��Ɣ���mail�������M","mailBody",mailBody

				if ComInterval < DateDiff("n",SendTime,Now) then
'					rc=ObjMail.Sendmail(svName, mailTo, mailFrom, mailSubject, mailBody, attachedFiles)
					sendTime=Now
				else
					ErrCode = 8
				end if

				If rc = 0 Then
					'''���[�����M���t�̍X�V���s���B
					StrSQL = "UPDATE TargetOperation SET UpdtTime='" & Now() & "', UpdtPgCd='dmi500',"
					StrSQL = StrSQL & " UpdtTmnl='" & USER & "',"&  UpdateSendTime & "='" & Now() & "'"
					StrSQL = StrSQL &"WHERE UserCode = '" & Trim(ObjRS("UserCode")) & "'"

					ObjConn.Execute(StrSQL)
					if err <> 0 then
						Set ObjRS = Nothing
						jumpErrorPDB ObjConn,"1","c104","14","��Ɣ���mail�������M","104","SQL:<BR>"&StrSQL
					end if
	
					'''���O�o��
					WriteLogH "c104", "��Ɣ���mail�������M","01",""
					ErrCode = 0
				else
					fp = Server.MapPath("./mailerror") & "\error.txt"
					set fobj = Server.CreateObject("Scripting.FileSystemObject")
						if rc<>"" then
							if fobj.FileExists(fp) = True then
								set tfile = fobj.OpenTextFile(fp,8)
							else
								set tfile = fobj.CreateTextFile(fp,True,False)
							end if
							tfile.WriteLine sendTime & " " & rc
							tfile.Close
							ErrCode = 8
						end if
				end if
			else
				ErrCode = 1
			end if
		end if

		'''DB�ڑ�����
		DisConnDBH ObjConn, ObjRS
		'''�G���[�g���b�v����
		on error goto 0

		Session.Contents("SendMailSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��Ɣ���mail�������M</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
<!--	window.close(); -->
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------�X�e�[�^�X�z�Mmail�������M���ʉ��--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="dmi500" method="POST">
	<TR><TD>�@</TD></TR>
<% if ErrCode=0 then %>
	<TR>
		<TD align="center">
			���[�����M���܂����B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
<% elseif ErrCode=1 then %>
	<TR>
		<TD align="center">
			���[�����M�悪�ݒ肳��Ă��܂���B<BR>�u�ݒ�v���j���[�ɂă��[���A�h���X��o�^���Ă��������B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
<% elseif ErrCode=8 then %>
	<TR>
		<TD align="center">
			���[�����M�Ɏ��s���܂����B<BR>
		</TD>
	</TR>
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
<% end if %>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%''' if Session.Contents("SendMailSubmitted") = "False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��Ɣ���mail�������M</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function CloseWin(){
	try{
		window.opener.parent.DList.location.href="sst100L.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------��Ɣ���mail�������M���ʉ��--------------------------->
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="dmi500" method="POST">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�����͊��Ɋ������Ă��܂��B<BR><BR><BR>
		</TD>
	</TR>
	<TR>
		<TD align="center">
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("SendMailSubmitted") = "False"��endif���� %>
<% end if %>
