<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>���o���\��U�蕪�����</title>
</head>

<body>
<%
	Dim sYMD
	Dim conn, rsd
	Dim sName, sTerm_Name, sTerminal, sType
	Dim sql

	'�w����t�擾(����ʈ��p)
	sYMD = TRIM(Request.QueryString("YMD"))

	'��Ǝ��ԑ�(����ʈ��p)
	sName = TRIM(Request.QueryString("NAME"))

	'�I���b�x�^�u�o
	sTerminal = TRIM(Request.Form("SELECT"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'�b�x�^�u�o�敪�E�^�[�~�i���擾
	Call GetTerminal2(conn, rsd, sTerminal, sType, sTerm_Name)
'''	sql = "SELECT Terminal, Type, Name FROM sTerminal" & _
'''		  " WHERE RTRIM(Terminal) = '" & sTerminal & "'"
'''	rsd.Open sql, conn, 0, 1, 1

'''	sType  = Trim(rsd("Type"))		
'''	sTerm_Name  = Trim(rsd("Name"))
'''	rsd.Close

	'�c�a�ؒf
	conn.Close
%>
	<INPUT TYPE=hidden NAME="Terminal" VALUE=<%=sTerminal%>>
	<INPUT TYPE=hidden NAME="Type" VALUE=<%=sType%>>
	<INPUT TYPE=hidden NAME="Term_Name" VALUE=<%=sTerm_Name%>>
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
	<INPUT TYPE=hidden NAME="Name" VALUE=<%=sName%>>
<%
	If sType = "Y" Then	'�b�x
%>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB010.asp?YMD=" + YMD.value + 
		             "&NAME=" + Name.value + "&Term_Name=" + Term_Name.value +
                     "&Terminal=" + Terminal.value);
</SCRIPT>
<%
	Else %>				'�u�o
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB062.asp?YMD=" + YMD.value + 
		             "&NAME=" + Name.value + "&Term_Name=" + Term_Name.value +
                     "&Terminal=" + Terminal.value);
</SCRIPT>
<%	End If
%>
</body>
</html>
