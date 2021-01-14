<%@Language="VBScript" %>
<!--#include file="Common.inc"-->

<html>

<style type="text/css">
<!--
	/* 説明文 */
	td.explain{
		font-size:12px;
		color:#000000;
		font-weight:bold;
	}
-->
</style>

<head>
	<title>着離岸情報照会</title>
	<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>

<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table border="0">
  <tr>
	<td width="110">
	</td>
	<td class="explain" valign="bottom">
	  福岡市港湾局　入出港予定船情報より
	</td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td valign="bottom">
<%
	DispMenuBar
%>
	</td>
  </tr>

</table>

<%
	DispMenuBarBack2 "javascript:history.back()"
%>

</body>
</html>
