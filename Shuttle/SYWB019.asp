<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->

<html>

<head>

<title>���Ԙg�J�����</title>
</head>

<body>
<%
	Dim conn, rsd
	Dim sYMD, sHH, sHHName
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sTrgDate, sDateNow, sDate, iSTime, iETime,contval
 	Dim sOpenFlag(23)

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Right(sYMD, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("NAME"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'�O���[�v���ԑя��̎擾��J���t���O�̎擾
	Call GetGrpSlot(conn, rsd, sGrpID, sYMD, sOpenFlag)

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title24.gif" width="236" height="34"><p>
<table border="1">   
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">��Ǝ���</td>
		<td width="360" ><%=ChgYMDStr2(sYMD)%>�@<%=sHHName%></td>
	</tr>
</table>
</center>
   
			<br>
			<br>
    <form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB030.asp?YMD=<%=sYMD%>&HH=<%=sHH%>" >
<center>
	<SELECT NAME="SELECT1">
	<%  if sOpenFlag(int(sHH)) = "Y" then	%>
			<OPTION selected VALUE="No1" >�J������
			<OPTION VALUE="No2" >�J�����Ȃ�
	<%	else	%>
			<OPTION VALUE="No1" >�J������
			<OPTION selected VALUE="No2" >�J�����Ȃ�
	<%	end if	%>
	</SELECT>

			<br>
			<br>
</center>

<center>
<table border=0>
		<td><input type="submit" value="�@���s�@" id=submit4 name=submit4 ></td>
	</form>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="�@���~�@" id=submit4 name=submit4></td>
	</form>
</table>
</center>

<br>     
<br>     
</body>     
</html>     
