<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi310.asp				_/
'_/	Function	:���O�������ԍ����͉��		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
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
  WriteLogH "b402", "���������O������","00",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O�o�^�E������</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(200,400);
window.focus();

function GoNext(){
  strA    = new Array("�u�b�L���O�ԍ�","�R���e�i�ԍ�");
  target=document.dmi310F;
  targetA    = new Array();
  targetA[0] = target.BookNo;
  targetA[1] = target.CONnum;
  for(k=0;k<2;k++){
    Num=LTrim(targetA[k].value);
    if(Num.length==0){
      alert(strA[k]+"���L�����Ă�������");
      targetA[k].focus();
      return;
    }
    if(k==0){
      if(!CheckEisu(targetA[k].value)){
        alert(strA[k]+"�ɔ��p�p�����Ɣ��p�X�y�[�X�A�u-�v�A�u/�v�ȊO�̕������L�����Ȃ��ł�������");
        targetA[k].focus();
        return;
      }
    }else{
      if(!CheckEisu2(targetA[k].value)){
        alert(strA[k]+"�ɔ��p�p�����ȊO�̕������L�����Ȃ��ł�������");
        targetA[k].focus();
        return;
      }
    }
  }
  chengeUpper(target);
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi310F.BookNo.focus();
}
//2008-01-31 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0  onload="finit();">
<!-------------�������ԍ����͉��--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD height="300" align=center>
<%'Mod-s 2006/03/06 h.matsuda%>
<!-----<FORM name="dmi310F" method="POST" action="./dmi315.asp">--->
      <FORM name="dmi310F" method="POST" action="./dmi312.asp">
	  <INPUT type=hidden name="ShoriMode" value="FLin">
<%'Mod-e 2006/03/06 h.matsuda%>
        <B>�u�b�L���O�ԍ�</B><BR>
	  <INPUT type=text  name="BookNo" maxlength=20 size=27><BR>
        <B>�R���e�i�ԍ�</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><P>
	  <A HREF="JavaScript:GoNext()">���s</A><P>
	  <A HREF="JavaScript:window.close()">����</A><P>
      </FORM>
  </TD></TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
