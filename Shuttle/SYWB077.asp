<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB077.inc"-->
<html>

<head>
<title>空シャーシ数結果画面</title>
</head>
<body>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>

<center>
空シャーシ数結果画面<br><br>

		<font face="ＭＳ ゴシック">
<%
	Dim conn, rsd, sql
	Dim sYMD, sHHName, sLackChassis
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim i20, i40

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	
	'指定日付取得
	sYMD    = TRIM(Request.QueryString("TDATE"))

	'指定時間帯取得
	sHHName = TRIM(Request.QueryString("HHName"))

	'空きシャーシ数の取得
	sLackChassis = GetEmptychassis(conn, rsd, sGrpID, sYMD, sHHName, i20, i40)

	Response.Write "グループＩＤ　　　　" & sGrpID
	Response.Write "<br><br>"

	Response.Write "指定日付　　　　　　" & TRIM(Request.QueryString("TDATE"))
	Response.Write "<br><br>"

%>
		<A HREF="JavaScript:history.back()">
		<BR>搬出入予約画面へ戻る</A>
</center>

</body>
</html>
 