<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo292.asp				_/
'_/	Function	:���O����o�w����������		_/
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
  WriteLogH "b309", "����o�w�������", "02",""
  
  dim i,j,conInfo,Num
  Redim conInfo(5)
  j=0
  For i=0 To 4
    conInfo(j)=Array("","","","","","")
    conInfo(j)(0)=Request("ContSize"&i)
    conInfo(j)(1)=Request("ContType"&i)
    conInfo(j)(2)=Request("ContHeight"&i)
    conInfo(j)(3)=Request("Material"&i)
    conInfo(j)(4)=Request("PickPlace"&i)
    conInfo(j)(5)=Request("PickNum"&i)
    If conInfo(j)(0)="" AND conInfo(j)(1)="" AND conInfo(j)(2)="" AND conInfo(j)(3)="" AND conInfo(j)(4)="" AND conInfo(j)(5)="" Then
    Else
      j=j+1
    End If
  Next
  Num=j
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
<!-------------����o�w����������--------------------------->
<CENTER><B class=titleB>����o�w����</B></CENTER>
<DIV style=text-align:right;>�쐬&nbsp;<%=Request("day")%></DIV><BR>
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD valign=top>�w����</TD><TD valign=top>��<%=Request("SjManN")%></TD>
    <TD>�i�S���ҁF�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�j<BR>
        <%=Request("TelNo")%></TD></TR>
  <TR>
    <TD>��Ǝ�</TD><TD>��<%=Request("WkManN")%></TD>
    <TD></TD></TR>
  <TR>
    <TD colspan=2>�u�b�L���O�ԍ��@�@�E�E�E�E�E�E</TD><TD><%=Request("BookNo")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=85% align=center>
  <TR><TD>����</TD><TD>�T�C�Y</TD><TD>�^�C�v</TD><TD>����</TD><TD>�ގ�</TD><TD>�s�b�N�ꏊ</TD><TD></TD><TD>�{��</TD></TR>
<% For i=0 To Num-1%>
  <TR><TD><%=i+1%></TD>
      <TD><%=conInfo(i)(0)%>'</TD><TD><%=conInfo(i)(1)%></TD>
      <TD><%=conInfo(i)(2)%></TD><TD><%=conInfo(i)(3)%></TD>
      <TD><%=conInfo(i)(4)%></TD><TD>�E�E�E</TD>
      <TD><%=conInfo(i)(5)%></TD></TR>
<% Next %>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=4 valign=top>�P�D</TH>
    <TD nowrap><B>�u�b�L���O���&nbsp;</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�D�Ёj</TD><TD><%=Request("shipFact")%></TD></TR>
  <TR>
    <TD>�i�D���j</TD><TD><%=Request("shipName")%></TD></TR>
  <TR>
    <TD>�i�d���n�j</TD><TD><%=Request("delivTo")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=6 valign=top>�Q�D</TH>
    <TD><B>�o���l�ߏ��</B></TD><TD></TD></TR>
  <TR>
    <TD>�i�o���l�ߓ����j</TD><TD><%=Request("vanDate")%></TD></TR>
  <TR>
    <TD valign=top nowrap>�i�o���l�ߏꏊ�P�j&nbsp;</TD><TD><%=Request("vanPlace1")%></TD></TR>
  <TR>
    <TD valign=top nowrap>�i�o���l�ߏꏊ�Q�j</TD><TD><%=Request("vanPlace2")%></TD></TR>
  <TR>
    <TD>�i�i���j</TD><TD><%=Request("goodsName")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�R�D</TH>
    <TD><B>�������</B></TD><TD></TD></TR>
  <TR>
    <TD>�i������b�x�j</TD><TD><%=Request("Terminal")%></TD></TR>
  <TR>
    <TD nowrap>�i�b�x�J�b�g���j&nbsp;</TD><TD><%=Request("CYCut")%></TD></TR>
</TABLE><P>

<TABLE border=0 cellPadding=1 cellSpacing=0>
  <TR><TH rowspan=3 valign=top>�S�D</TH>
    <TD><B>���l</B></TD><TD></TD></TR>
  <TR>
    <TD valign=top nowrap>�i���l�P�j&nbsp;</TD><TD><%=Request("Comment1")%></TD></TR>
  <TR>
    <TD valign=top>�i���l�Q�j</TD><TD><%=Request("Comment2")%></TD></TR>
</TABLE><P>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
