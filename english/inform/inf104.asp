<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf104.asp                                      _/
'_/	Function	:���m�点���M����̍폜����                    _/
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
	'''�폜�X�V���܂�����ʂɂāu�ŐV�̏��ɍX�V�v��Submit���ꂽ�ꍇ�̑΍�
	if Session.Contents("DeleteSubmitted")="False" then

		'''�f�[�^�擾
		Dim EMAIL
		EMAIL = Request.Form("email")

		'''�G���[�g���b�v�J�n
		on error resume next
		'''DB�ڑ�
		Dim cn, rs, sql
		ConnDBH cn, rs

		sql="delete from send_information where email='" & EMAIL & "'"

		cn.Execute(sql)
		if err <> 0 then
			set rs = Nothing
			response.write("inf104.asp:send_information�e�[�u��delete�G���[!")
			response.end
		end if

		'''DB�ڑ�����
		DisConnDBH cn, rs
		'''�G���[�g���b�v����
		on error goto 0

		Session.Contents("DeleteSubmitted") = "True"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���m�点���M����폜</TITLE>
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
<!-------------���m�点���M����폜--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�폜���܂����B<BR><BR><BR>
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
<TITLE>���m�点���M����폜</TITLE>
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
<!-------------���m�点���M����폜--------------------------->
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
<FORM name="inf103">
	<TR><TD>�@</TD></TR>
	<TR>
		<TD align="center">
			�폜�͂��łɊ������Ă��܂��B<BR><BR><BR>
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
