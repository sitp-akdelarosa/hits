<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>時間枠開放画面更新画面</title>
</head>
<body>
<%
	Dim conn, rsd, sql
	Dim sYMD, sHH
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sTrgDate, sDateNow, sDate, iSTime, iETime
 	Dim sOpenFlag(23), sAns,i

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = TRIM(Request.QueryString("HH"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'グループ時間帯情報の取得後開放フラグの取得
	Call GetGrpSlot(conn, rsd, sGrpID, sYMD, sOpenFlag)

	if Request.Form("select1") = "No1" then
		sAns = "Y"
	else
		sAns = " "
	end if

	'環境変数取得
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
<B>更新中</B>
</CENTER>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>

</body>
</html>
 