<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf103.asp                                      _/
'_/	Function	:���m�点���M����̍X�V����                    _/
'_/	Date			:2005/03/10                                      _/
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
	'''�X�V���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("UpdateSubmitted")="False" then

		'''�f�[�^�擾
		Dim EMAIL, DANTAI_CODE, COMPANY_NAME, USER_NAME, TEL, ADDRESS
		EMAIL = Request.Form("email")
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

		sql="update send_information set "
		sql=sql & " UpdtTime='" & Now() & "',"
		sql=sql & " group_code='" & DANTAI_CODE & "',"
		sql=sql & " company_name='" & COMPANY_NAME & "',"
		sql=sql & " user_name='" & USER_NAME & "',"
		sql=sql & " tel='" & TEL & "',"
		sql=sql & " address='" & ADDRESS & "'"
		sql=sql & " where email='" & EMAIL & "'"

		cn.Execute(sql)
		if err <> 0 then
			set rs = Nothing
			response.write("inf103.asp:send_information�e�[�u��update�G���[!")
			response.end
		end if

		'''DB�ڑ�����
		DisConnDBH cn, rs
		'''�G���[�g���b�v����
		on error goto 0

		Session.Contents("UpdateSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���m�点���M����X�V</TITLE>
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
<!-------------���m�点���M����X�V--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�X�V���܂����B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>

<%'''if Session.Contents("UpdateSubmitted")="False"��else���� %>
<% else %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���m�点���M����X�V</TITLE>
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
<!-------------���m�点���M����X�V--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�X�V�͂��łɊ������Ă��܂��B<BR><BR><BR>
			<INPUT type="button" value="����" onClick="CloseWin()">
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
<%'''if Session.Contents("UpdateSubmitted")="False"��endif���� %>
<% end if %>
