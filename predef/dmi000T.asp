<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000T.asp				_/
'_/	Function	:���O���ꗗ��ʃw�b�_			_/
'_/	Date		:2003/05/26				_/
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
  dim DayTime,day
  '�T�[�o�����̎擾
  getDayTime DayTime
  day = DayTime(0) & "�N" & DayTime(1) & "��" & DayTime(2) & "��" &_
        DayTime(3) & "��" & DayTime(4) & "�����݂̏��"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>���O���ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("../codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=350");
  CodeWin.focus();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------���O���͏������Top--------------------------->
<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%" height="40">
   <TR>
     <TD rowspan="3" width="506" valign=top><IMG src="Image/predef_title.gif" width="506" height="40"></TD>
     <TD align="right" bgColor="#000099" height="14" colspan="2"><IMG src="Image/logo_hits_ver2.gif" height="14" width="300"></TD>
   </TR>
   <TR>
   	  <TD height="2" colspan="2"></TD>
   </TR>
   <TR>
   	   <TD align="center" valign="top"><%=day%></TD>
       <TD align=right>
	   	<Form>
          <INPUT type="button" value="�R�[�h�ꗗ" OnClick="OpenCodeWin()" style="height:22;">
          <SPAN style="width:20;"></SPAN>
        </Form>
	   </TD>
   </TR>
</TABLE>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
