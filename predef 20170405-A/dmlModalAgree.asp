<% @LANGUAGE = VBScript %>
<%
%><% Option Explicit %>
<%
	'HTTP�R���e���c�^�C�v�ݒ�
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
  Dim ChkAgree
  Dim ChkSolas
  Dim BookChk
  Dim IMDGChk										'2016/11/03 H.Yoshikawa Add
  ChkAgree = Trim(Request.QueryString("ChkAgr"))
  ChkSolas = Trim(Request.QueryString("ChkSls"))
  BookChk = Trim(Request.QueryString("BookChk"))
  IMDGChk = Trim(Request.QueryString("IMDGChk"))	'2016/11/03 H.Yoshikawa Add
'�Z�b�V�����̗L�������`�F�b�N
  CheckLoginH
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE></TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function finit()
{
    if("<%=ChkAgree%>" == "1" && "<%=ChkSolas%>" == "1" && "<%=BookChk%>" == "0"  && "<%=IMDGChk%>" == "0" ){		// 2016/11/03 H.Yoshikawa Upd(IMDGChk�ǉ�)
        fRgst();
    }
}
function fBack()
{
   returnValue = false;
   window.close();
}
function fRgst()
{
  returnValue = true;
  window.close();
}
-->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit()">
<form name="frm" method="post">

<table border=0 cellPadding=1 cellSpacing=0 width="100%">
<tr>
<td align=center>
<TABLE border=0 cellPadding=4 cellSpacing=0>
  <tr>
  <td colspan=2 align=center>
	<div><BR></div>
    <div>���̓o�^�́A���L�̗��R�ɂ��u���o�^�v��ԂƂȂ�A�Q�[�g��t�͂ł��܂���B</div>
	<div><BR></div>
	<div><BR></div>
  <% If Trim(ChkSolas) <> "1" Then %>
	<div align=left>���u�����ɓ��͂����R���e�i�O���X��SOLAS���Ɋ�Â����@�Ōv�����ꂽ���l�ł��B�v��<BR>�@�`�F�b�N������܂���B</div>
  <%End If%>
  <% If Trim(ChkAgree) <> "1" Then %>
	<div align=left>���u�{��ʂ̓��͓��e���Q�[�g�ł̔����[�̑���Ƃ��Ďg�p���邱�Ƃɓ��ӂ��܂��B�v��<BR>�@�`�F�b�N������܂���B</div>
  <%End If%>
  <% If Trim(BookChk) <> "0" Then %>
	<div align=left>���u�b�L���O���ƈقȂ�l������܂��B�i�ԐF�\���j</div>
  <%End If%>
  <% '2016/11/03 H.Yoshikawa Add Start %>
  <% If Trim(IMDGChk) <> "0" Then %>
	<div align=left>���댯�i�R�[�h�Ɍ�肪����܂��B�i�ԐF�\���j</div>
  <%End If%>
  <% '2016/11/03 H.Yoshikawa Add End %>
	<div><BR></div>
	<div><BR></div>
    <div>�{�o�^���s���ꍇ�́A�u�߂�v�{�^���������āA���L�̒l���C���̂������o�^���������B</div>
    <div>���̂܂܁u���o�^�v���s���܂����H</div>
  </td>
  </tr>
  <tr><td><BR /></td></tr>
  <tr>
  <td align=center><input type="button" name="Back" value="�߂�" Onclick="fBack();" onkeypress="return true"></td>
  <td align=center><input type="button" name="Rgst" value="���o�^" Onclick="fRgst();" onkeypress="return true"></td>
  </tr>
</TABLE>
</td>
</tr>
</table>
</div>
</form>
</BODY>
</HTML>
