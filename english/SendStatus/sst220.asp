<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst200.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�˗��폜���			_/
'_/	Date			:2004/01/13				_/
'_/	Code By		:aspLand HARA			_/
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

	'''�f�[�^�擾
	Dim USER, ContBLNo, ContORBL
	USER = Session.Contents("userid")
	ContORBL = Request.Form("ContORBL")
	ContBLNo = Request.Form("ContBLNo")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�˗��폜</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
	window.resizeTo(450,180);
	bgset(target);
	window.focus();
}
//�폜
function GoDelete(){
	f=document.sst220;
	if(!confirm("�폜���Ă���낵���ł����H")){
		return;
	}
	f.action="sst221.asp";
	f.submit();
}
//�R���e�i���Ɖ�
function GoInfo(){
	f=document.sst220;
	f.action="sst900.asp";
	newWin = window.open("", "ConInfo", "status=yes,scrollbars=yes,resizable=yes,menubar=yes");
	f.target="ConInfo";
	f.submit();
	f.target="_self";
}
//mail�������M
function GoSendmail(){
	f=document.sst220;
	if(!confirm("���M���Ă���낵���ł����H")){
		return;
	}
	f.Mode.value=2;		//�폜��ʂ��mail�������M�����s�����ꍇ
	f.action="sst500.asp";
	f.submit();
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"  onLoad="setParam(document.sst220)">
<!-------------�X�e�[�^�X�z�M�˗��폜���--------------------------->
<% Session.Contents("DeleteSubmitted")="False"  %>
<% Session.Contents("SendMailSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst220" method="POST">
	<TR>
		<TD colspan="3">
			<B>Import Status Delivery Request Information</B><BR>
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">Login User</DIV></TD>
		<TD width="60%" colspan="2">
			<INPUT type="text" name="LoginUser" value="<%=USER%>" size="15" readonly style="background-color:#E0E0E0;color:#000000;">
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">Container No. / BL No.</DIV></TD>
		<TD width="40%">
			<INPUT type="text" name="ContBLNo" value="<%=ContBLNo%>" size="30" readonly style="background-color:#E0E0E0;color:#000000;">
		</TD>
		<TD width="20%">
			<INPUT type="hidden" name="ContORBL" value="<%=ContORBL%>" >
			<INPUT type="button" value="SEARCH" onClick="GoInfo()">
		</TD>
	</TR>
	<TR>
		<TD colspan="3" align="center">
			<INPUT type="hidden" name="Mode" value="">
			<INPUT type="button" value="Delete" onClick="GoDelete()">
			<INPUT type="button" value="Close" onClick="window.close()">�@
			<A HREF="javascript:GoSendmail();">Real Time Delivery</A>
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
