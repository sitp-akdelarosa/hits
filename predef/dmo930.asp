<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi930.asp				_/
'_/	Function	:�A�o���(�ꗗ)��ʐڑ�			_/
'_/	Date		:2003/05/29				_/
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
  dim inPutStr,formStr
'�f�[�^�擾
  dim BookNo,CONnum
  BookNo = Request("BookNo")
  CONnum = Request("CONnum")
'2006/03/06 add-s h.matsuda
  dim Num,ShipLine,ShoriMode
	Num=BookCountSL(BookNo)'�u�b�L���O�ԍ��̏d���`�F�b�N
	ShoriMode = Trim(Request("ShoriMode"))
	ShipLine = Trim(Request("ShipLine"))
'2006/03/06 add-e h.matsuda

  If BookNo <> "" Then
    formStr="<FORM method=post action='../bookcheck.asp' name='dmi930F'>"		'CW-019
    inPutStr="<INPUT type=hidden name='booking' value='"& BookNo &"'>"
    Session.Contents("route") = "�A�o�R���e�i���Ɖ�i��ƑI���j > �u�b�L���O���Ɖ� >  "'CW-011
  Else
    formStr="<FORM action='../expcntnr.asp' name='dmi930F'>"		'CW-019
    inPutStr="<INPUT type=hidden name='cntnrno' value='"& CONnum &"'>"
    Session.Contents("route") = "�A�o�R���e�i���Ɖ�i��ƑI���j "	'CW-011
  End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>�]����</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function opnewin(){
  //2006/03/06 add-s h.matsuda	�D�БI����ʂ��Ăяo���ꂽ�ꍇ�̂ݕύX���̏������s��
	<% If Num>1 and ShipLine="" and BookNo<>"" and ShoriMode<>"" then %>
		document.dmi930F.action="./dmi312.asp";
		document.dmi930F.submit();
	<% Else %>
		window.focus();
		document.dmi930F.submit();
	<% End If %>
//  window.focus();
//  document.dmi930F.submit();
  //2006/03/06 add-e h.matsuda
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY onLoad="opnewin()">
<P>�]����...���΂炭���҂����������B</P>
<!--CW-019 <FORM action="../expcntnr.asp" name="dmi930F"> -->
<%= formStr%><%'CW-019%>
<%= inPutStr %>
<%'Mod-s 2006/03/06 h.matsuda%>
	  <INPUT type=hidden name="BookNo" value="<%=BookNo%>">
	  <INPUT type=hidden name="ShoriMode" value="EMoutInf">
	  <INPUT type=hidden name="ShipLine" value="<%=ShipLine%>">
<%'Mod-e 2006/03/06 h.matsuda%>
</FORM>
</BODY></HTML>
