<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits                                          _/
'_/	FileName	:inf100.asp                                      _/
'_/	Function	:���m�点���[���A�h���X���͉��                  _/
'_/	Date			:2005/03/03                                      _/
'_/	Code By		:aspLand HARA                                    _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'''HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���[���A�h���X�o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<!--<SCRIPT src="./js/common.js"></SCRIPT>-->
<SCRIPT language=JavaScript>
<!--
window.resizeTo(500,150);
window.focus();

function GoNext(){
	f=document.inf100;

	if(f.email.value == ""){
		alert("���[���A�h���X����͂��Ă��������B");
		f.email.focus();
		return false;
	}else{
		if(gfisMailAddr(f.email.value)==false){
			alert("���[���A�h���X���s���ł��B\n���[���A�h���X���m�F���Ă��������B");
			f.email.focus();
			return false;
		}
	}
	f.submit();
}

//���[���A�h���X�`�F�b�N
function gfisMailAddr(a){
	if(a==""){
		return(true);
	}
	var b=a.replace(/[a-zA-Z0-9_@\.\-]/g,'');
	if(b.length!=0){
		return(false);
	}
	var p1=a.indexOf("@");
	var p2=a.lastIndexOf("@");
	var p3=a.lastIndexOf(".");
	if(0<p1 && p1==p2 && p1<p3 && p3<a.length-1 ){
		return(true);
	}
	return(false);
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY bgcolor="DEE1FF" text="#000000" link="#3300FF" leftMargin="0" topMargin="0" marginheight="0" marginwidth="0">
<!-------------���[���A�h���X���͉��--------------------------->
<% Session.Contents("InsertSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="inf100" action="inf101.asp" method="post" onsubmit="return false;">
	<TR>
		<TD colspan="2">
			<b><font color="navy">�o�^�^�Q�Ɓ^�폜�������A�h���X����͂��Ă��������B</font></b><BR>
		</TD>
	</TR>
	<TR>
		<TD width="30%" align="right">���[���A�h���X�F</TD>
		<TD width="70%">
			<INPUT type="text" name="email" value="" size="40" maxlength="50">
		</TD>
	</TR>
	<TR>
		<TD colspan="2" align="center">
			<INPUT type="button" value="����" onClick="javascript:GoNext()">
			<INPUT type="button" value="���~" onClick="window.close()">�@
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
