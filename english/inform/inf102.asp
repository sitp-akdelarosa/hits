<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf102.asp                                      _/
'_/	Function	:���m�点���[���A�h���X���V�K�o�^����            _/
'_/	Date			:2005/03/09                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
	'''�o�^���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("InsertSubmitted")="False" then

		'''�f�[�^�擾
		Dim EMAIL, DANTAI_CODE, COMPANY_NAME, USER_NAME, TEL, ADDRESS
		EMAIL = Trim(Request.Form("email"))
		DANTAI_CODE = Request.Form("dantai")
		COMPANY_NAME = Trim(Request.Form("company_name"))
		USER_NAME = Trim(Request.Form("user_name"))
		TEL = Trim(Request.Form("tel"))
		ADDRESS = Trim(Request.Form("address"))

		'''�G���[�g���b�v�J�n
		on error resume next
		'''DB�ڑ�
		Dim cn, rs, sql
		ConnDBH cn, rs

		sql="insert into send_information(email,UpdtPgCd,UpdtTmnl,UpdtTime,group_code,company_name,user_name,tel,address) "
		sql=sql & " values("
		sql=sql & "'" & EMAIL & "',"
		sql=sql & "'Sendinfo',"
		sql=sql & "'Sendinfo',"
		sql=sql & "'" & Now() & "',"
		sql=sql & "'" & DANTAI_CODE & "',"
		sql=sql & "'" & COMPANY_NAME & "',"
		sql=sql & "'" & USER_NAME & "',"
		sql=sql & "'" & TEL & "',"
		sql=sql & "'" & ADDRESS & "')"

		cn.Execute(sql)
		if err <> 0 then
			set rs = Nothing
			response.write("inf101.asp:send_information�e�[�u��insert�G���[!")
			response.end
		end if

		'''DB�ڑ�����
		DisConnDBH cn, rs
		'''�G���[�g���b�v����
		on error goto 0

		Session.Contents("InsertSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���[���A�h���X�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="inf101.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------���m�点���[���A�h���X�V�K�o�^--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf101">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�o�^���܂����B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("InsertSubmitted")="False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���m�点���[���A�h���X�V�K�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT Language="JavaScript">
<!--
function CloseWin(){
	try{
		window.opener.parent.List.location.href="inf101.asp"
	}catch(e){}
	window.close();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------���m�点���[���A�h���X�V�K�o�^--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf100">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�o�^�͂��łɊ������Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("InsertSubmitted")="False"��endif���� %>
<% end if %>
