<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst200.asp				_/
'_/	Function	:�X�e�[�^�X�z�M�˗��o�^���			_/
'_/	Date			:2004/01/07				_/
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
	Dim USER
	USER = Session.Contents("userid")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�X�e�[�^�X�z�M�˗��o�^</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
	window.resizeTo(450,180);
	bgset(target);
	window.focus();
}

function GoNext(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("�R���e�i�ԍ��܂��͂a�k�ԍ�����͂��Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("�R���e�i�ԍ��ɔ��p�p�����ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("�R���e�i�ԍ��͂P�Q�����ȓ��Ŏw�肵�Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("�a�k�ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTY�͓o�^�ł��܂���B");
		f.ContBLNo.focus();
		return;
	}

	changeUpper(f);
	f.action="sst201.asp";
	f.submit();
}

function GoSendmail(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("�R���e�i�ԍ��܂��͂a�k�ԍ�����͂��Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("�R���e�i�ԍ��ɔ��p�p�����ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("�R���e�i�ԍ��͂P�Q�����ȓ��Ŏw�肵�Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("�a�k�ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTY�͎w��ł��܂���B");
		f.ContBLNo.focus();
		return;
	}

	if(!confirm("���M���Ă���낵���ł����H")){
		f.ContBLNo.focus();
		return;
	}
	f.Mode.value=1;		//�V�K�o�^��ʂ��mail�������M�����s�����ꍇ
	f.action="sst500.asp";
	f.submit();
}
//�R���e�i���Ɖ�
function GoInfo(){
	f=document.sst200;
	Number=LTrim(f.ContBLNo.value);
	if(Number.length==0){
		alert("�R���e�i�ԍ��܂��͂a�k�ԍ�����͂��Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && !CheckEisuji(Number)){
		alert("�R���e�i�ԍ��ɔ��p�p�����ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[0].checked && Number.length>12){
		alert("�R���e�i�ԍ��͂P�Q�����ȓ��Ŏw�肵�Ă��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && !CheckEisu(Number)){
		alert("�a�k�ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕�������͂��Ȃ��ł��������B");
		f.ContBLNo.focus();
		return;
	}
	if(f.ContORBL[1].checked && Number.toUpperCase() == "EMPTY"){
		alert("EMPTY�͎w��ł��܂���B");
		f.ContBLNo.focus();
		return;
	}

	f.action="sst900.asp";
	newWin = window.open("", "ConInfo", "status=yes,scrollbars=yes,resizable=yes,menubar=yes");
	f.target="ConInfo";
	f.submit();
	f.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"  onLoad="setParam(document.sst200)">
<!-------------�X�e�[�^�X�z�M�˗��o�^���--------------------------->
<% Session.Contents("InsertSubmitted")="False"  %>
<% Session.Contents("SendMailSubmitted")="False"  %>
<TABLE border="0" cellPadding="3" cellSpacing="1" width="100%">
<FORM name="sst200" method="POST">
	<TR>
		<TD colspan="3">
			<B>�A���X�e�[�^�X�z�M�˗��o�^</B><BR>
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">���O�C�����[�U</DIV></TD>
		<TD width="60%" colspan="2">
			<INPUT type="text" name="LoginUser" value="<%=USER%>" size="10" readonly style="background-color:#E0E0E0;color:#000000;">
		</TD>
	</TR>
	<TR>
		<TD width="40%"><DIV class="bgb">�ΏۃR���e�iNo.�^�a�kNo.</DIV></TD>
		<TD width="40%">
			<INPUT type="text" name="ContBLNo" value="" size="27" maxlength="20">
		</TD>
		<TD width="20%">
			<INPUT type="button" value="���Ɖ�" onClick="GoInfo()">
		</TD>
	</TR>
	<TR>
		<TD colspan="3" align="center">
			<INPUT type="radio" name="ContORBL" value="1" checked>�R���e�i�@
			<INPUT type="radio" name="ContORBL" value="2">�a�k
		</TD>
	</TR>
	<TR>
		<TD colspan="3" align="center">
			<INPUT type="hidden" name="Mode" value="">
			<INPUT type="button" value="�o�^" onClick="GoNext()">
			<INPUT type="button" value="���~" onClick="window.close()">�@
			<A HREF="javascript:GoSendmail();">mail�������M</A>
		</TD>
	</TR>
</FORM>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY>
</HTML>
