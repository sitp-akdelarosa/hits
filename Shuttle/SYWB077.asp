<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB077.inc"-->
<html>

<head>
<title>��V���[�V�����ʉ��</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>

<center>
��V���[�V�����ʉ��<br><br>

		<font face="�l�r �S�V�b�N">
<%
	Dim conn, rsd, sql
	Dim sYMD, sHHName, sLackChassis
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim i20, i40

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'�w����t�擾
	sYMD    = TRIM(Request.QueryString("TDATE"))

	'�w�莞�ԑю擾
	sHHName = TRIM(Request.QueryString("HHName"))

	'�󂫃V���[�V���̎擾
	sLackChassis = GetEmptychassis(conn, rsd, sGrpID, sYMD, sHHName, i20, i40)

	Response.Write "�O���[�v�h�c�@�@�@�@" & sGrpID
	Response.Write "<br><br>"

	Response.Write "�w����t�@�@�@�@�@�@" & TRIM(Request.QueryString("TDATE"))
	Response.Write "<br><br>"

%>
		<A HREF="JavaScript:history.back()">
		<BR>���o���\���ʂ֖߂�</A>
</center>

</body>
</html>
 