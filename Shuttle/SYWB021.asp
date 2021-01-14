
<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>シャーシ属性設定画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
/* 決定ボタン */
function ClickSend() {

	/* 入力チェック */

	if  (document.UPLOAD1.SELECT1.value == "No0" && document.UPLOAD1.SELECT2.value == "No0" &&
		document.UPLOAD1.sy_zaiko.value == "") {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
	if ((document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT1.value != "No0") || 
		(document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT2.value != "No0") ||
		(document.UPLOAD1.SELECT1.value != "No0" && document.UPLOAD1.SELECT2.value != "No0")) {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
}

/* 詳細確認ボタン */
function ClickSend1(go) {

	/* 入力チェック */

	if  (document.UPLOAD1.SELECT1.value == "No0" && document.UPLOAD1.SELECT2.value == "No0" &&
		document.UPLOAD1.sy_zaiko.value == "") {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
	if ((document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT1.value != "No0") || 
		(document.UPLOAD1.sy_zaiko.value != ""   && document.UPLOAD1.SELECT2.value != "No0") ||
		(document.UPLOAD1.SELECT1.value != "No0" && document.UPLOAD1.SELECT2.value != "No0")) {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}

/*手入力*/
	if (document.UPLOAD1.sy_zaiko.value != "")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.sy_zaiko.value.toUpperCase()
		return true;
	}

/*ＳＹ在庫選択入力*/
	if (document.UPLOAD1.SELECT1.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.SELECT1.value
		return true;
	}

/*ＳＹ非在庫選択入力*/
	if (document.UPLOAD1.SELECT2.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.SELECT2.value
		return true;
	}
}
//--->
</SCRIPT>

</head>

<body>
<%
	Dim conn, rsd, sql											'ＤＢ接続
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator			'ユーザ情報
	Dim sYMD													'指定日付
	Dim sDispChassis1, sDispChassis2, sPlateNo, sChk1, sChk2	'シャーシ表示情報
	Dim i
	Dim sNO, sChkChassisID										'未使用
	Dim sErr1, sErr2, sErr3, sChassis							'未使用

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))

	'シャーシ表示情報の取得
	sDispChassis1 = ""
	sDispChassis2 = ""
	sChk1 = ""
	sChk2 = ""
	
	sPlateNo = sDispChassis1 & "　" & sDispChassis2

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="ＭＳ ゴシック">
   
<center>
<form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB024.asp?YMD=<%=sYMD%>" onSubmit="return ClickSend()">
<table border="1" width="420"  >
<b><font color=#000080>対象シャーシ</font></b>
		<tr bgcolor=#ffff99><td><br>
				ＳＹ在庫より選択　　<SELECT NAME="SELECT1">
						<OPTION VALUE="No0" >　
						<%	i = 1
							sql = "SELECT * FROM sChassis" & _
								  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
								  "  AND StackFlag <> ' '"
							sql = sql & "  Order By ChassisId"
							rsd.Open sql, conn, 0, 1, 1
			
							if not rsd.eof then
								do while not rsd.EOF%>
									<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
									<%rsd.MoveNext
									i = i + 1
								loop
							end if
							rsd.Close
						%>
					</SELECT><br>
				ＳＹ非在庫より選択　<SELECT NAME="SELECT2">
						<OPTION VALUE="No0" >　
						<%	i = 1
							sql = "SELECT * FROM sChassis" & _
								  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
								  "  AND StackFlag = ' '"
							sql = sql & "  Order By ChassisId"
							rsd.Open sql, conn, 0, 1, 1
							if not rsd.eof then
								do while not rsd.EOF%> 
									<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
									<%rsd.MoveNext
									i = i + 1
								loop
							end if
							rsd.Close
						%>
					</SELECT><br>
	手入力する場合　　　<INPUT TYPE="text" NAME="sy_zaiko" SIZE="9" MAXLENGTH="5" value=<%=sDispChassis1%>>　<input type="button" value="シャーシの詳細確認" id=submit5 name=submit5 onclick="return ClickSend1(this)">
		<br>
		</td></tr>
</table><br><br>
</center>
<br>
<center>
<table border=0>
		<td><input type="submit"  value="　決定　" id=submit4 name=submit4></td>
	</form>
	<td>　</td>
	<td>　</td>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="　中止　" id=submit6 name=submit6></td>
	</form>
</table>
</center>

<br>     
<br>     
</body>     
</html>