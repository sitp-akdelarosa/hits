<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo092.asp				_/
'_/	Function	:���O�����o�w����������		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co ��d			_/
'_/	Modify		:								_/
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
'���O�o��
  WriteLogH "b109", "�����o�w�������", "02",""
'�R���e�i�f�[�^�擾
  dim preConInfo,preNum,ConInfo,Num,i,j
  Get_Data preNum,preConInfo
  
  If Request("checkNum")="" OR Request("checkNum")=Null Then
    Num=1
    ReDim ConInfo(1)
    ConInfo(0)=preConInfo(0)
  Else
    dim strChecks,tmptarget,targetNo
    Num=Request("checkNum")
    strChecks=Request("checkeds")
    ReDim ConInfo(Num)
    tmptarget=Split(strChecks, ",", -1, 1)
    For i=0 To Num-1
      targetNo=Mid(tmptarget(i),3)
      ConInfo(i)=preConInfo(targetNo)
    Next
  End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./styleprint.css">
<TITLE>�w����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.focus();
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------�����o�w����������--------------------------->
<CENTER><B class=titleB>�����o�w����</B></CENTER>
<DIV style=text-align:right;>�쐬&nbsp;<%=Request("day")%></DIV><BR>
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD>��Ɣԍ�</TD><TD>��<%=Request("SakuNo")%></TD><TD></TD></TR>
  <TR>
    <TD valign=top>�w����</TD><TD valign=top>��<%=Request("SjManN")%></TD>
    <TD>�i�S���ҁF�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�j<BR>
        <%=Request("TelNo")%></TD></TR>
  <TR>
    <TD>��Ǝ�</TD><TD>��<%=Request("WkManN")%></TD>
    <TD>�i�w�b�h�h�c��<%=Request("HedId")%>�j</TD></TR>
  <TR>
    <TD>�w����@</TD><TD>��<%=Request("Way")%></TD><TD></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>�P�D</TH>
    <TD nowrap><B>�R���e�i���</B>&nbsp;</TD><TD></TD></TR>
  <TR>
    <TD>�i�D�Ёj</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>�i�D���j</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>�i�i���j</TD><TD><%=Request("HinName")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>�Q�D</TH>
    <TD><B>���o���</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�b�x�j</TD><TD><%=Request("Hfrom")%></TD></TR>
  <TR>
    <TD nowrap>�i���o�\������j&nbsp;</TD><TD><%=Request("RDate")%></TD></TR>
  <TR>
    <TD valign=top>�i�[����P�j</TD><TD><%=Request("Nonyu1")%></TD></TR>
  <TR>
    <TD valign=top>�i�[����Q�j</TD><TD><%=Request("Nonyu2")%></TD></TR>
  <TR>
    <TD>�i�[���������j</TD><TD><%=Request("NoDate")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�R�D</TH>
    <TD><B>��R���ԋp���</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�ԋp��j</TD><TD><%=Request("RPlace")%></TD></TR>
  <TR>
    <TD nowrap>�i�ԋp�\������j&nbsp;</TD><TD><%=Request("Rnissu")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�S�D</TH>
    <TD><B>���l</B></TD><TD></TD></TR>
  <TR>
    <TD valign=top nowrap>�i���l�P�j&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD valign=top>�i���l�Q�j</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=<%=Num+3%> valign=top>�T�D</TH>
    <TD colspan=7><B>�R���e�i�ԍ�</B></TD></TR>
  <TR><TD width=20></TD><TD>����</TD><TD>&nbsp;�R���e�i�ԍ�&nbsp;</TD><TD>&nbsp;�T�C�Y&nbsp;</TD>
      <TD>&nbsp;�^�C�v&nbsp;</TD><TD>&nbsp;����&nbsp;</TD><TD>&nbsp;�O���X&nbsp;</TD>
  <TR align=center><TD></TD>
    <TD>1</TD>
    <TD><%=ConInfo(0)(0)%></TD><TD><%=ConInfo(0)(1)%>'</TD><TD><%=ConInfo(0)(2)%></TD>
    <TD><%=ConInfo(0)(3)%></TD><TD><%=ConInfo(0)(4)%>kg</TD></TR>
<% For i=1 To Num-1 %>
  <TR align=center><TD></TD>
    <TD><%=i+1%></TD>
    <TD><%=ConInfo(i)(0)%></TD><TD><%=ConInfo(i)(1)%>'</TD><TD><%=ConInfo(i)(2)%></TD>
    <TD><%=ConInfo(i)(3)%></TD><TD><%=ConInfo(i)(4)%>kg</TD></TR>
<%Next%>
</TABLE><P>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
