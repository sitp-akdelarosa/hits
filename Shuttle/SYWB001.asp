<%@ LANGUAGE="VBScript" %>

<html>

<head>
<title>���o���Ɖ���</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
function ClickInquiry() {
}
//--->
</SCRIPT>

</head>

<body >
<IMG border=0 height=42 src="image/title01.gif" width=311>
<center>
<p><IMG border=0 height=66 src="image/title11.gif" width=503><p>

<%
Set conn = Server.CreateObject("ADODB.Connection")
'conn.Open "HakataDB", "sa", "hakata"	'D20040314
conn.Open "HakataDB", "sa", ""		'I20040314
Set rsd = Server.CreateObject("ADODB.Recordset")
rsd.Open "sUseDB", conn, 0, 1, 2
if rsd.eof then
	rsd.Close
	conn.Close
	Response.Write "�V�X�e���G���[:�g�pDB�Ǘ��e�[�u���Ƀ��R�[�h������܂���B"
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
else
	wOutUpdtTime = rsd("OutUpdtTime" & rsd("EnableDB")) 
%>
	��&nbsp; ���݂̏���&nbsp; <u><b><%=Month(wOutUpdtTime)%>                              ��<%=Day(wOutUpdtTime)%>                                       ��
										<%=FormatDateTime(wOutUpdtTime, vbShortTime)%></b></u>&nbsp; �̂��̂ł��B<br><br> 
	   (&nbsp; ����X�V�\���&nbsp; <b><%=Month(rsd("OutPUpdtTime"))%>                              ��<%=Day(rsd("OutPUpdtTime"))%>                                       ��
										<%=FormatDateTime(rsd("OutPUpdtTime"), vbShortTime)%></b>&nbsp; �ł��B&nbsp;) 
<%
end if
rsd.Close
conn.Close

contval=""
blval=""
tsubmit="��    ��"
%>

<!--#include file="ComnForm.inc"-->

<br>
<br>
<p><A href="index.asp">���j���[�ɖ߂�</A></p>

</center>

</body>

</html>
