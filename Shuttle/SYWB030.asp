<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>���Ԙg�J����ʍX�V���</title>
</head>
<body>
<%
	Dim conn, rsd, sql
	Dim sYMD, sHH
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sTrgDate, sDateNow, sDate, iSTime, iETime
 	Dim sOpenFlag(23), sAns,i

	'�w����t�擾
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = TRIM(Request.QueryString("HH"))

	'�c�a�ڑ�
	Call ConnectSvr(conn, rsd)

	'���[�U���̎擾
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'�O���[�v���ԑя��̎擾��J���t���O�̎擾
	Call GetGrpSlot(conn, rsd, sGrpID, sYMD, sOpenFlag)

	if Request.Form("select1") = "No1" then
		sAns = "Y"
	else
		sAns = " "
	end if

	'���ϐ��擾
	sql = "SELECT * FROM sGrpSlot" & _
		  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
		  "   AND Date = '" & sYMD & "'"
	rsd.Open sql, conn, 0, 2, 1

	If Not rsd.EOF Then
		rsd("OpenFlag" & Trim(Cstr(int(sHH)))) = sAns
		rsd("UpdtTime") = now()
		rsd.update
	Else
		rsd.addnew

		rsd("GroupID") = sGrpID
		rsd("Date") = sYMD
		rsd("UpdtPgCd") = "SYWB0030"
		rsd("UpdtTmnl") = "WEB"
		rsd("UpdtTime") = now()
				
		For i = 0 To 23
			if int(i) = int(sHH) then
				rsd("OpenFlag" & Trim(CLng(i))) = sAns
			else
				rsd("OpenFlag" & Trim(CLng(i))) = "Y"
			end if	
		Next
		rsd.update
	End If

	rsd.Close
%>
<CENTER>
<B>�X�V��</B>
</CENTER>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>

</body>
</html>
 