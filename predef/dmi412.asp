<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits										   _/
'_/	FileName	:dmi411.asp									   _/
'_/	Function	:��Ɣ���mail�Ώۍ��ړo�^�E�X�V				   _/
'_/	Date		:2009/03/10									   _/
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
'''�Z�b�V�����̗L�������`�F�b�N
	CheckLoginH

	'''�f�[�^�o�^�^�X�V���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	'''�܂��f�[�^�o�^�^�X�V���܂�����ʂ͕\������Ă��Ȃ��ꍇ
if Session.Contents("ItemsSubmitted")="False" then

''�T�[�o���t�̎擾
	Dim DayTime
	getDayTime DayTime

''���[�U�f�[�^����
	Dim USER, sUN, Utype
	USER   = UCase(Session.Contents("userid"))
	sUN    = Session.Contents("sUN")
	Utype  = Session.Contents("UType")

''�f�[�^�擾
	Dim F_DelResults(4), F_RecEmp(4), F_RecResults(4), F_DelEmp(4)
	Dim Email1, Email2, Email3, Email4, Email5
	Dim iCnt,tmpstr

	F_DelResults(0) = Request.Form("F_DelResults1")
	F_DelResults(1) = Request.Form("F_DelResults2")
	F_DelResults(2) = Request.Form("F_DelResults3")
	F_DelResults(3) = Request.Form("F_DelResults4")
	F_DelResults(4) = Request.Form("F_DelResults5")
	
	F_RecEmp(0) = Request.Form("F_RecEmp1")
	F_RecEmp(1) = Request.Form("F_RecEmp2")
	F_RecEmp(2) = Request.Form("F_RecEmp3")
	F_RecEmp(3) = Request.Form("F_RecEmp4")
	F_RecEmp(4) = Request.Form("F_RecEmp5")
	
	F_RecResults(0) = Request.Form("F_RecResults1")
	F_RecResults(1) = Request.Form("F_RecResults2")
	F_RecResults(2) = Request.Form("F_RecResults3")
	F_RecResults(3) = Request.Form("F_RecResults4")
	F_RecResults(4) = Request.Form("F_RecResults5")
	
	F_DelEmp(0) = Request.Form("F_DelEmp1")
	F_DelEmp(1) = Request.Form("F_DelEmp2")
	F_DelEmp(2) = Request.Form("F_DelEmp3")
	F_DelEmp(3) = Request.Form("F_DelEmp4")
	F_DelEmp(4) = Request.Form("F_DelEmp5")
	
	Email1 = Request.Form("Email1")
	Email2 = Request.Form("Email2")
	Email3 = Request.Form("Email3")
	Email4 = Request.Form("Email4")
	Email5 = Request.Form("Email5")
	
'�G���[�g���b�v�J�n
	on error resume next
	''DB�ڑ�
	Dim ObjConn, ObjRS, StrSQL, RecordCNT
	ConnDBH ObjConn, ObjRS

	StrSQL = "SELECT count(*) AS NUM from TargetOperation where UserCode='"& USER &"'"
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		Session.Contents("dmi411") = "false"
		DisConnDBH ObjConn, ObjRS	'DB�ؒf
		jumpErrorP "2","c103","10","��Ɣ���mail���ړo�^�E�X�V","101","SQL:<BR>"&strSQL
	else
		RecordCNT = ObjRS("NUM")
		ObjRS.close
	end if

	'''�V�K�o�^�̏ꍇ
	if RecordCNT = 0 then
		StrSQL = "INSERT INTO TargetOperation(UserCode,UpdtTime,UpdtPgCd,UpdtTmnl,"
		StrSQL = StrSQL & "Email1,Email2,Email3,Email4,Email5,"
		StrSQL = StrSQL & "FlagDelResults1,FlagRecEmp1,FlagRecResults1,FlagDelEmp1,"
		StrSQL = StrSQL & "FlagDelResults2,FlagRecEmp2,FlagRecResults2,FlagDelEmp2,"
		StrSQL = StrSQL & "FlagDelResults3,FlagRecEmp3,FlagRecResults3,FlagDelEmp3,"
		StrSQL = StrSQL & "FlagDelResults4,FlagRecEmp4,FlagRecResults4,FlagDelEmp4,"
		StrSQL = StrSQL & "FlagDelResults5,FlagRecEmp5,FlagRecResults5,FlagDelEmp5)"
		StrSQL = StrSQL & "values('" & USER & "','" & Now() & "','STATUS01','" & USER & "',"

		StrSQL = StrSQL & "'" & Email1 & "','" &Email2 & "','" & Email3 & "','" & Email4 & "','" & Email5 & "'"

		For iCnt = 0 To 4
			StrSQL = StrSQL & ","
			StrSQL = StrSQL & "'" & F_DelResults(iCnt) & "','" & F_RecEmp(iCnt) & "','" & F_RecResults(iCnt) & "','" & F_DelEmp(iCnt) & "'"
		Next
		
		StrSQL = StrSQL & ")"
		
		ObjConn.Execute(StrSQL)
		if err <> 0 then
			Session.Contents("dmi411") = "false"
			Set ObjRS = Nothing
			jumpErrorPDB ObjConn,"2","c104","10","��Ɣ���mail���ړo�^�E�X�V","103","SQL:<BR>"&StrSQL
		end if
	''�X�V�̏ꍇ
	else
		StrSQL = "UPDATE TargetOperation SET UpdtTime='"& Now() &"', UpdtPgCd='STATUS01', UpdtTmnl='"& USER &"',"

		StrSQL = StrSQL & " Email1='" & Email1 & "',"
		StrSQL = StrSQL & " Email2='" & Email2 & "',"
		StrSQL = StrSQL & " Email3='" & Email3 & "',"
		StrSQL = StrSQL & " Email4='" & Email4 & "',"
		StrSQL = StrSQL & " Email5='" & Email5 & "',"
		
		StrSQL = StrSQL & " FlagDelResults1='" & F_DelResults(0) & "',"
		StrSQL = StrSQL & " FlagRecEmp1='" & F_RecEmp(0) & "',"
		StrSQL = StrSQL & " FlagRecResults1='" & F_RecResults(0) & "',"
		StrSQL = StrSQL & " FlagDelEmp1='" & F_DelEmp(0) & "',"

		StrSQL = StrSQL & " FlagDelResults2='" & F_DelResults(1) & "',"
		StrSQL = StrSQL & " FlagRecEmp2='" & F_RecEmp(1) & "',"
		StrSQL = StrSQL & " FlagRecResults2='" & F_RecResults(1) & "',"
		StrSQL = StrSQL & " FlagDelEmp2='" & F_DelEmp(1) & "',"
		
		StrSQL = StrSQL & " FlagDelResults3='" & F_DelResults(2) & "',"
		StrSQL = StrSQL & " FlagRecEmp3='" & F_RecEmp(2) & "',"
		StrSQL = StrSQL & " FlagRecResults3='" & F_RecResults(2) & "',"
		StrSQL = StrSQL & " FlagDelEmp3='" & F_DelEmp(2) & "',"
		
		StrSQL = StrSQL & " FlagDelResults4='" & F_DelResults(3) & "',"
		StrSQL = StrSQL & " FlagRecEmp4='" & F_RecEmp(3) & "',"
		StrSQL = StrSQL & " FlagRecResults4='" & F_RecResults(3) & "',"
		StrSQL = StrSQL & " FlagDelEmp4='" & F_DelEmp(3) & "',"
		
		StrSQL = StrSQL & " FlagDelResults5='" & F_DelResults(4) & "',"
		StrSQL = StrSQL & " FlagRecEmp5='" & F_RecEmp(4) & "',"
		StrSQL = StrSQL & " FlagRecResults5='" & F_RecResults(4) & "',"
		StrSQL = StrSQL & " FlagDelEmp5='" & F_DelEmp(4) & "'"
		
		StrSQL = StrSQL & " WHERE UserCode = '" & USER & "'"
		
		
		
		ObjConn.Execute(StrSQL)
		if err <> 0 then
		Session.Contents("dmi411") = "false"
			Set ObjRS = Nothing
			jumpErrorPDB ObjConn,"2","d104","11","��Ɣ���mail���ړo�^�E�X�V","103","SQL:<BR>" & StrSQL
		end if
	end if

'DB�ڑ�����
	DisConnDBH ObjConn, ObjRS
'�G���[�g���b�v����
	on error goto 0

	Session.Contents("dmi411") = "false"

	tmpstr = Email1 & "," & Email2 & "," & Email3 & "," & Email4 & "," & Email5
						
	For iCnt = 0 To 4
		tmpStr = tmpStr & "," & _
		F_DelResults(iCnt) & "," & _ 
		F_RecEmp(iCnt) & "," & F_RecResults(iCnt) & "," & F_DelEmp(iCnt)
	Next
	
	if RecordCNT = 0 then
		WriteLogH "d103", "��Ɣ���mail���ړo�^�E�X�V","10",tmpstr
	else
		WriteLogH "d103", "��Ɣ���mail���ړo�^�E�X�V","11",tmpstr
	end if

	Session.Contents("ItemsSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��Ɣ���mail�Ώۍ��ړo�^�E�X�V</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------��Ɣ���mail�ݒ荀�ړo�^�E�X�V--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD align=center>
<% if RecordCNT = "0" then %>
			<BR><BR>�o�^���܂����B<BR><BR><BR>
<% end if %>
<% if RecordCNT = "1" then %>
			<BR><BR>�X�V���܂����B<BR><BR><BR>
<% end if %>
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("ItemsSubmitted")="False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��Ɣ���mail�ݒ荀�ړo�^�E�X�V</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------��Ɣ���mail�ݒ荀�ړo�^�E�X�V--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
	<TR>
		<TD align=center>
			<BR><BR>�o�^�E�X�V�͊��Ɋ������Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="window.close()">
		</TD>
	</TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("ItemsSubmitted")="False"��endif���� %>
<% end if %>
