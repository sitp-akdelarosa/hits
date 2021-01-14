<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>”Ào“ü—\–ñU‚è•ª‚¯‰æ–Ê</title>
</head>

<body>
<%
	Dim sYMD
	Dim conn, rsd
	Dim sName, sTerm_Name, sTerminal, sType
	Dim sql

	'w’è“ú•tæ“¾(Ÿ‰æ–ÊˆøŒp)
	sYMD = TRIM(Request.QueryString("YMD"))

	'ì‹ÆŠÔ‘Ñ(Ÿ‰æ–ÊˆøŒp)
	sName = TRIM(Request.QueryString("NAME"))

	'‘I‘ğ‚b‚x^‚u‚o
	sTerminal = TRIM(Request.Form("SELECT"))

	'‚c‚aÚ‘±
	Call ConnectSvr(conn, rsd)

	'‚b‚x^‚u‚o‹æ•ªEƒ^[ƒ~ƒiƒ‹æ“¾
	Call GetTerminal2(conn, rsd, sTerminal, sType, sTerm_Name)
'''	sql = "SELECT Terminal, Type, Name FROM sTerminal" & _
'''		  " WHERE RTRIM(Terminal) = '" & sTerminal & "'"
'''	rsd.Open sql, conn, 0, 1, 1

'''	sType  = Trim(rsd("Type"))		
'''	sTerm_Name  = Trim(rsd("Name"))
'''	rsd.Close

	'‚c‚aØ’f
	conn.Close
%>
	<INPUT TYPE=hidden NAME="Terminal" VALUE=<%=sTerminal%>>
	<INPUT TYPE=hidden NAME="Type" VALUE=<%=sType%>>
	<INPUT TYPE=hidden NAME="Term_Name" VALUE=<%=sTerm_Name%>>
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
	<INPUT TYPE=hidden NAME="Name" VALUE=<%=sName%>>
<%
	If sType = "Y" Then	'‚b‚x
%>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB010.asp?YMD=" + YMD.value + 
		             "&NAME=" + Name.value + "&Term_Name=" + Term_Name.value +
                     "&Terminal=" + Terminal.value);
</SCRIPT>
<%
	Else %>				'‚u‚o
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB062.asp?YMD=" + YMD.value + 
		             "&NAME=" + Name.value + "&Term_Name=" + Term_Name.value +
                     "&Terminal=" + Terminal.value);
</SCRIPT>
<%	End If
%>
</body>
</html>
