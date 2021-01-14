<%@Language="VBScript" %>

<!--#include file="Common.inc"-->


<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから照会画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/bookingt.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
%>
					</td>
				</tr>
			</table>
			<center>
			<table width=95% cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td align="right">
						<font color="#333333" size="-1">
						<%=strRoute%>
						</font>
					</td>
				</tr>
			</table>
			<BR><BR><BR><BR><BR>
			<table border="0">
				<tr>
					<td align="center">
						セキュリティ上運用を中止しています。
						ただし、事前情報入力機能の中からは照会可能です。<br>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td align="center">
						<input type="button" value=" OK " onclick="javascript:history.back();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td valign="bottom">
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------照会画面終わり--------------------------->
<%
    DispMenuBarBack "index.asp"
%>
</body>
</html>
