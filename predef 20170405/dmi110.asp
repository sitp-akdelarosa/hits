<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi110.asp				_/
'_/	Function	:���O��������̓R���e�i�I�����		_/
'_/	Date		:2003/05/28				_/
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
  WriteLogH "b202", "��������O������","00",""	'CW-046
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O�o�^�E�����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(200,450);
window.focus();

function GoNext(){
  target=document.dmi110F
  Num=LTrim(target.CONnum.value);
  if(Num.length==0){
    alert("�R���e�i�ԍ����L�����Ă�������");
    target.CONnum.focus();
    return;
  }
  if(!CheckEisu2(target.CONnum.value)){
    alert("�R���e�i�ԍ��ɔ��p�p�����ȊO�̕������L�����Ȃ��ł�������");
    target.CONnum.focus();
    return;
  }
  chengeUpper(target);
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi110F.CONnum.focus();
}
//2008-01-31 Add-E M.Marquez
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit();">
<!-------------��������̓R���e�i�I�����--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD height="300" align=center>
      <FORM name="dmi110F" method="POST" action="./dmi115.asp">
        <B>�R���e�i�ԍ��Ŏw��</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><BR>
	  <A HREF="JavaScript:GoNext()">���s</A><P>
	<A HREF="JavaScript:window.close()">����</A><P>
      </FORM>
  </TD></TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
