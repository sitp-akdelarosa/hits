<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo220.asp				_/
'_/	Function	:���O����o���͕\�����			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/08/06	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���S�ʉ��C	_/
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
  WriteLogH "b302", "����o���O������","11",""

'�f�[�^����
  dim COMPcd0,ret,compF,i
  COMPcd0= Request("COMPcd0")
  compF  = Request("compF")

'�X�V���[�h�t���O�ݒ�
  ret=true
  If compF<>0 AND COMPcd0 <> UCase(Session.Contents("userid")) Then
    ret=false
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>����o���\���m�F</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  window.resizeTo(550,680);
  window.focus();
  bgset(target);
}
//�X�V��ʂ�
function GoReEntry(){
  target=document.dmo220F;
  target.action="./dmi220.asp";
  target.submit();
}
//�u�b�L���O���
function GoBookI(){
  target=document.dmo220F
  BookInfo(target);
}
//�w�������������ʂ�
function GoSijiPrint(){
  target=document.dmo220F;
  target.action="./dmo291.asp";
//  newWin = window.open("", "Print", "width=500,height=700,left=30,top=10,resizable=yes,scrollbars=yes,top=0");
//  target.target="Print";
  target.submit();
//  target.target="_self";
}
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo220F)">
<!-------------����o���\���m�F���--------------------------->
<FORM name="dmo220F" method="POST">
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2>
      <B>����o������(�\�����[�h)</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>�u�b�L���O�m���D</DIV></TD>
    <TD><INPUT type=text name="BookNoM" value="<%=Request("BookNoM")%>" readOnly size=40>
        <INPUT type=hidden name="BookNo" value="<%=Request("BookNo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipFact" value="<%=Request("shipFact")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>�D��</DIV></TD>
    <TD><INPUT type=text name="shipName" value="<%=Request("shipName")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>�d���n</DIV></TD>
    <TD><INPUT type=text name="delivTo" value="<%=Request("delivTo")%>" readOnly size=40></TD></TR>
  <TR>
    <TD><DIV class=bgb>��ЃR�[�h(���^)</DIV></TD>
    <TD><INPUT type=text name="COMPcd1" value="<%=Request("COMPcd1")%>" size=5  readOnly>
        <INPUT type=hidden name="oldCOMPcd1" value="<%=Request("oldCOMPcd1")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����Ɩ{��</DIV></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>
    <TABLE border=0 cellPadding=0 cellSpacing=0 width=400 align=center>
      <TR><TD></TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR>
<% For i=0 To 4%>
      <TR><TD>(<%=i+1%>)</TD>
          <TD><INPUT type=text name="ContSize<%=i%>"   value="<%=Request("ContSize"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContType<%=i%>"   value="<%=Request("ContType"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="ContHeight<%=i%>" value="<%=Request("ContHeight"&i)%>" size=4  readOnly></TD>
          <TD><INPUT type=text name="Material<%=i%>"   value="<%=Request("Material"&i)%>"   size=4  readOnly></TD>
          <TD><INPUT type=text name="PickPlace<%=i%>"  value="<%=Request("PickPlace"&i)%>"  size=25 readOnly></TD>
          <TD>�E�E�E</TD>
          <TD><INPUT type=text name="PickNum<%=i%>" value="<%=Request("PickNum"&i)%>" size=4  readOnly></TD></TR>
<% Next %>
    </TABLE>
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߓ���</DIV></TD>
    <TD><INPUT type=text name="vanMon" value="<%=Request("vanMon")%>" size=3  readOnly>��
        <INPUT type=text name="vanDay" value="<%=Request("vanDay")%>" size=3  readOnly>��
        <INPUT type=text name="vanHou" value="<%=Request("vanHou")%>" size=3  readOnly>��
        <INPUT type=text name="vanMin" value="<%=Request("vanMin")%>" size=3  readOnly>��
        </TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�P</DIV></TD>
    <TD><INPUT type=text name="vanPlace1" value="<%=Request("vanPlace1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�o���l�ߏꏊ�Q</DIV></TD>
    <TD><INPUT type=text name="vanPlace2" value="<%=Request("vanPlace2")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>�i��</DIV></TD>
    <TD><INPUT type=text name="goodsName" value="<%=Request("goodsName")%>" size=30  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>������b�x�D�b�x�J�b�g��</DIV></TD>
    <TD><INPUT type=text name="Terminal" value="<%=Request("Terminal")%>" readOnly>
        <INPUT type=text name="CYCut" value="<%=Request("CYCut")%>" readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�P</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73  readOnly></TD></TR>
  <TR>
    <TD><DIV class=bgb>���l�Q</DIV></TD>
    <TD><INPUT type=text name="Comment2" value="<%=Request("Comment2")%>" size=73  readOnly></TD></TR>
    
  <TR>
<!-- 2009/03/10 R.Shibuta Add-S -->
   <TD><DIV class=bgy>�o�^�S����</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName@" readonly = "readonly" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  </TR>
  
  <TR>
    <TD colspan=2 align=center>
<%'Add-S 2009/10/22 H.Fujiyama%>
       <INPUT type=hidden name="TruckerSubName" value="<%=Request("TruckerSubName")%>" >
<%'Add-E 2009/10/22 H.Fujiyama%>
       <INPUT type=hidden name=Mord value="<%=Request("Mord")%>" >
       <INPUT type=hidden name=COMPcd0 value="<%=COMPcd0%>" >
       <INPUT type=hidden name="TFlag" value="<%=Request("TFlag")%>">
<%'Add-s 2006/03/06 h.matsuda%>
       <INPUT type=hidden name=shipline value="<%=Request("shipline")%>" >
	   <INPUT type=hidden name="ShoriMode" value="EMoutInf">
<%'Add-e 2006/03/06 h.matsuda%>
<%' If COMPcd0 = UCase(Session.Contents("userid")) Then  '''Del 20040301%>
       <INPUT type=button value="�w�������" onClick="GoSijiPrint()">
<%' End If '''Del 20040301%>
<% If ret Then %>
       <INPUT type=hidden name="compFlag" value="<%=compF%>">
       <INPUT type=submit value="�X�V���[�h" onClick="GoReEntry()">
<% End If %>
       <INPUT type=submit value="����" onClick="window.close()">
       <P>
       <INPUT type=button value="�u�b�L���O���" onClick="GoBookI()">
    </TD></TR>

</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>