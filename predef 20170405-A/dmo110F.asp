<%
	@LANGUAGE = VBScript
	@CODEPAGE = 932
%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo110F.asp				_/
'_/	Function	:��������ꗗ��ʃt���[��		_/
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>��������ꗗ</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<!-------------��������ꗗ���Frame--------------------------->
<frameset rows="100,*,35" border="0" frameborder="0" name="110Frame">
  <frame src="./dmo110T.asp" name="Top" scrolling="no" noresize>
  <frame src="./dmo110L.asp" name="DList" scrolling="no">
  <frame src="./dmo110B.asp" name="Bottom" scrolling="no" noresize>
  <noframes>
  ���̃y�[�W�̓t���[���Ή��̃u���E�U�ł������������B
  </noframes>
</frameset>
<BODY>
<!-------------��ʏI���--------------------------->
</BODY></HTML>
