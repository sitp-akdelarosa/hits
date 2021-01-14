<%@ LANGUAGE="VBScript" %>

<html>

<head>
<title>搬出許可照会画面</title>
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
	Response.Write "システムエラー:使用DB管理テーブルにレコードがありません。"
	Response.Write "</body>"
	Response.Write "</html>"
	Response.End
else
	wOutUpdtTime = rsd("OutUpdtTime" & rsd("EnableDB")) 
%>
	★&nbsp; 現在の情報は&nbsp; <u><b><%=Month(wOutUpdtTime)%>                              月<%=Day(wOutUpdtTime)%>                                       日
										<%=FormatDateTime(wOutUpdtTime, vbShortTime)%></b></u>&nbsp; のものです。<br><br> 
	   (&nbsp; 次回更新予定は&nbsp; <b><%=Month(rsd("OutPUpdtTime"))%>                              月<%=Day(rsd("OutPUpdtTime"))%>                                       日
										<%=FormatDateTime(rsd("OutPUpdtTime"), vbShortTime)%></b>&nbsp; です。&nbsp;) 
<%
end if
rsd.Close
conn.Close

contval=""
blval=""
tsubmit="照    会"
%>

<!--#include file="ComnForm.inc"-->

<br>
<br>
<p><A href="index.asp">メニューに戻る</A></p>

</center>

</body>

</html>
