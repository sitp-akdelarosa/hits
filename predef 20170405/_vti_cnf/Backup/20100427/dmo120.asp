<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo120.asp				_/
'_/	Function	:���O��������͕\�����			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:C-002	2003/07/29	���l���ǉ�	_/
'_/	Modify		:3th	2003/01/31	3���ύX	_/
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
  WriteLogH "b202", "��������O������","11",""

'�f�[�^�擾
  dim Mord,CONnum,CMPcd(5),HedId,Rmon,Rday,UpFlag
  dim param,i,j
  Mord   = Request("Mord")
  CONnum = Request("CONnum")
  UpFlag=Request("UpFlag")
  For Each param In Request.Form
    If Left(param,5) = "CMPcd" Then
      j = Right(param,1)
      CMPcd(j) = Request.Form(param)
    End If
  Next
'�\����������
'3th  Rmon = Request("Rmon")
'3th  Rday = Request("Rday")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��������\��</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>

window.resizeTo(640,530);
window.focus();
<!--
function setParam(target){
  for (i=0; i<20; i++) target.elements[i].readOnly = true;
  bgset(target);
}

//�R���e�i�ڍ׉��
function GoConInfo(){
  target=document.dmo120F;
  ConInfo(target,1,0);
  return false;
}
//�X�V��ʂ�
function GoReEntry(){
  target=document.dmo120F;
  target.action="./dmi120.asp";
  return true;
}

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onLoad="setParam(document.dmo120F)">
<!-------------����������͕\�����--------------------------->
<FORM name="dmo120F" method="POST">
<TABLE border=0 cellPadding=5 cellSpacing=0 width="100%">
  <TR>
    <TD colspan=2><B>�����������(�\�����[�h)</B></TD></TR>
  <TR>
    <TD><DIV class=bgb>�R���e�i�m���D</DIV></TD>
    <TD><INPUT type=text name="CONnum" value="<%=CONnum%>"></TD></TR>
  <TR>
    <TD width=230><BR><DIV class=bgb>��ЃR�[�h</DIV></TD>
    <TD>�o�^��<BR>
        <INPUT type=text name="CMPcd0" value="<%=CMPcd(0)%>" size=7>
        <INPUT type=text name="CMPcd1" value="<%=CMPcd(1)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd2" value="<%=CMPcd(2)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd3" value="<%=CMPcd(3)%>" size=5 maxlength=2>
        <INPUT type=text name="CMPcd4" value="<%=CMPcd(4)%>" size=5 maxlength=2>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD Align=right>�w�����S����</TD>
    <TD>
        <INPUT type=text name="SubName" readonly = "readonly" value="<%=Request("TruckerSubName")%>" maxlength=16>
    </TD></TR>
<!-- 2009/10/09 Add-S Fujiyama -->
  <TR>
    <TD><DIV class=bgb>�w�b�h�h�c</DIV></TD>
    <TD><INPUT type=text name="HedId" value="<%=Request("HedId")%>" maxlength=5></TD></TR>
  <TR>
    <TD><DIV class=bgb>�ԋp��</DIV></TD>
    <TD><INPUT type=text name="HTo" value="<%=Request("HTo")%>"></TD></TR>
  <TR>
    <TD><DIV class=bgb>�����\���</DIV></TD>
    <TD><INPUT type=text name="Rmon" value="<%=Request("Rmon")%>" size=2>��
        <INPUT type=text name="Rday" value="<%=Request("Rday")%>" size=2>��
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�T�C�Y�A�^�C�v�A�����A�ގ��A�e�A�E�F�C�g</DIV></TD>
    <TD><INPUT type=text name="CONsize" value="<%=Request("CONsize")%>" size=5>
        <INPUT type=text name="CONtype" value="<%=Request("CONtype")%>" size=5>
        <INPUT type=text name="CONhite" value="<%=Request("CONhite")%>" size=5>
        <INPUT type=text name="CONsitu" value="<%=Request("CONsitu")%>" size=5>
        <INPUT type=text name="CONtear" value="<%=Request("CONtear")%>" size=5>kg
    </TD></TR>
  <TR>
    <TD><DIV class=bgb>�戵�D��</DIV></TD>
    <TD><INPUT type=text name="TrhkSen" value="<%=Request("TrhkSen")%>" size=27></TD></TR>
  <TR>
    <TD><DIV class=bgb>�ۊ�</DIV></TD>
    <TD><INPUT type=text name="MrSk" value="<%=Request("MrSk")%>" size=5>
�@�@</TD></TR>
  <TR>
    <TD><DIV class=bgb>�l�`�w�d��</DIV></TD>
    <TD><INPUT type=text name="MaxW" value="<%=Request("MaxW")%>" maxlength=6>kg</TD></TR>
<%'C-002 ADD Start %>
  <TR>
    <TD><DIV class=bgb>���l</DIV></TD>
    <TD><INPUT type=text name="Comment1" value="<%=Request("Comment1")%>" size=73></TD></TR>
<%'C-002 ADD End %>
<!-- 2009/03/10 R.Shibuta Add-S -->
  <TR>
   <TD><DIV class=bgy>�o�^�S����</DIV></TD>
   <TD><INPUT type=text name="TruckerSubName" readonly = "readonly" maxlength=16></TD>
<!-- 2009/03/10 R.Shibuta Add-E -->
  <TR>
    <TD colspan=2 align=center>
       <INPUT type=hidden name=Mord value="<%=Mord%>" >
       <INPUT type=hidden name=UpUser value="<%=Request("UpUser")%>" >
       <INPUT type=hidden name="UpFlag"  value="<%=UpFlag%>">
 <!-- 2009/08/04 Tanaka Add-S -->
       <INPUT type=hidden name="TruckerName" value="<%=Request("TruckerName")%>" >
 <!-- 2009/08/04 Tanaka Add-E -->
<%'20030909 IF Request("compFlag") AND (UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1) Then %>
<%'Chang 20050303 STAT fro Fourth Recon By SEIKO N.Oosige
' IF UCase(Session.Contents("userid"))=CMPcd(0) Or Request("TruckerFlag")<>1 Then 
  IF UCase(Session.Contents("userid"))=CMPcd(0) Or (Request("compFlag") AND Request("TruckerFlag")<>1) Then %>
       <INPUT type=hidden name="compFlag" value="<%=Request("compFlag")%>">
       <INPUT type=hidden name="WkCNo"     value="<%=Request("WkCNo")%>">
       <INPUT type=hidden name="TruckerFlag" value="<%=Request("TruckerFlag")%>">
       <INPUT type=submit value="�X�V���[�h" onClick="return GoReEntry()">
<%End IF%>
       <INPUT type=submit value="����" onClick="window.close()">
       <P>
       <INPUT type=submit value="�R���e�i���" onClick="return GoConInfo()">
    </TD></TR>
</TABLE>
</FORM>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
