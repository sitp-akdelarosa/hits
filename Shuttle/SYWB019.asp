<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->

<html>

<head>

<title>時間枠開放画面</title>
</head>

<body>
<%
	Dim conn, rsd
	Dim sYMD, sHH, sHHName
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sTrgDate, sDateNow, sDate, iSTime, iETime,contval
 	Dim sOpenFlag(23)

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Right(sYMD, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("NAME"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'グループ時間帯情報の取得後開放フラグの取得
	Call GetGrpSlot(conn, rsd, sGrpID, sYMD, sOpenFlag)

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title24.gif" width="236" height="34"><p>
<table border="1">   
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">作業時間</td>
		<td width="360" ><%=ChgYMDStr2(sYMD)%>　<%=sHHName%></td>
	</tr>
</table>
</center>
   
			<br>
			<br>
    <form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB030.asp?YMD=<%=sYMD%>&HH=<%=sHH%>" >
<center>
	<SELECT NAME="SELECT1">
	<%  if sOpenFlag(int(sHH)) = "Y" then	%>
			<OPTION selected VALUE="No1" >開放する
			<OPTION VALUE="No2" >開放しない
	<%	else	%>
			<OPTION VALUE="No1" >開放する
			<OPTION selected VALUE="No2" >開放しない
	<%	end if	%>
	</SELECT>

			<br>
			<br>
</center>

<center>
<table border=0>
		<td><input type="submit" value="　実行　" id=submit4 name=submit4 ></td>
	</form>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="　中止　" id=submit4 name=submit4></td>
	</form>
</table>
</center>

<br>     
<br>     
</body>     
</html>     
