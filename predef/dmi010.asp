<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi010.asp				_/
'_/	Function	:���O�����o���͕��@�I�����		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%><% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
  WriteLogH "b102", "�����o���O������(����)","00",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O�o�^�E�����o</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(600,400); // Edited by AK.DELAROSA 2021/01/12
window.focus();

function GoNext(n,m){
  if(n==1){
    Num=LTrim(document.dmi010F.CONnum.value);
    if(Num.length==0){
      alert("�R���e�i�ԍ����L�����Ă�������");
      document.dmi010F.CONnum.focus();
      return;
    }
    if(!CheckEisu2(document.dmi010F.CONnum.value)){
      alert("�R���e�i�ԍ��ɔ��p�p�����ȊO�̕������L�����Ȃ��ł�������");
      document.dmi010F.CONnum.focus();
      return;
    }
    switch(m){
	case 1:
          document.dmi010F.flag.value="1";
	  break;
	case 2:
          document.dmi010F.flag.value="2";
	  break;
        case 3:
          document.dmi010F.flag.value="3";
        break;
      }
  } else {
    Num=LTrim(document.dmi010F.BLnum.value);
    if(Num.length==0){
      alert("�a�k�ԍ����L�����Ă�������");
      document.dmi010F.BLnum.focus();
      return;
    }
    if(!CheckEisu(document.dmi010F.BLnum.value)){
      alert("�a�k�ԍ��ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������L�����Ȃ��ł�������");
      document.dmi010F.BLnum.focus();
      return;
    }
    document.dmi010F.flag.value="4";
  }
  chengeUpper(document.dmi010F);
  document.dmi010F.submit();
}
//2008-01-29 Add-S M.Marquez
function finit(){
    document.dmi010F.CONnum.focus();
}
//2008-01-29 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit();">
<!-------------�����o�w����--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD align=center>
      <FORM name="dmi010F" method="POST" action="./dmi015.asp">
        <B>�R���e�i�ԍ��Ŏw��</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><BR>
	  <A HREF="JavaScript:GoNext(1,1)">�w�肠����s</A><BR>
	  <A HREF="JavaScript:GoNext(1,2)">�w��Ȃ����s</A><BR>
<% If Session.Contents("UType")<>5 Then %>
	  <A HREF="JavaScript:GoNext(1,3)">�ꗗ����I�����s</A>
<% End If %>
          <P>
        <B>�a�k�ԍ��Ŏw��</B><BR>
	  <INPUT type=text  name="BLnum" maxlength=20><BR>
	  <A HREF="JavaScript:GoNext(2,0)">���s</A><P>
	<A HREF="JavaScript:window.close()">����</A><P>
        <INPUT type=hidden name="flag">
      </FORM>
  </TD></TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
