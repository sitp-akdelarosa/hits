<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi810.asp				_/
'_/	Function	:���O����oCSV���̓t�@�C���ݒ�		_/
'_/	Date		:2003/05/30				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b302", "����o���O������","05",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O����oCSV���̓t�@�C���ݒ�</TITLE>
<SCRIPT language=JavaScript>
<!--
window.focus();
//CW-025 ADD
function SendCSV(){
  if(document.dmi820F.fln.value.length==0){
    alert("�t�@�C�����w�肵�Ă��������B");
    return;
  }else{
    document.dmi820F.submit();
  }
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------���O����oCSV���̓t�@�C���ݒ�--------------------------->
<B>CSV�t�@�C�����M</B>
<CENTER>
  <FORM ACTION="./dmi820.asp" NAME="dmi820F" METHOD="POST" ENCTYPE="multipart/form-data">
    <P>���M����t�@�C�����w�肵�Ă�������<BR>
    <input type="file" name="fln" enctype="multipart/form-data" ><BR>
    <INPUT TYPE="HIDDEN" NAME="perm" SIZE="-1" VALUE="forb"></P>
    <P><INPUT TYPE="Button" VALUE="���M" onClick="SendCSV()">
       <INPUT type=button value="����" onClick="window.close()"></P>
  </FORM>
</CENTER>
</BODY></HTML>