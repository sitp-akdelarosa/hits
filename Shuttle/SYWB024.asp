
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
var  f1=false;
var  f2=false;
var  f3=false;

/* 登録ボタン */
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


	/* 処理選択ワーニングチェック */
	/*if  ((f1==true) && (f2==true)) {
		window.alert("属性の設定に矛盾があります。");
		return false;
	}*/
	if  ((document.UPLOAD1.check1.checked==true) && (document.UPLOAD1.check2.checked==true)) {
		window.alert("属性の設定に矛盾があります。");
		return false;
	}
}
/* 詳細確認ボタン */
function ClickSend2(go) {
		location.href = "SYWB023.asp?sCassis=" + document.UPLOAD1.sy_zaiko.value.toUpperCase()
		return true;
}
//--->
</SCRIPT>

</head>

<body>
<%
	Dim conn, rsd, sql											'ＤＢ接続
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator			'ユーザ情報
	Dim sYMD, sChassisID										'指定日付、シャーシID
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

	'シャーシID取得
	sChassisID = TRIM(Request.QueryString("TRGID"))
	If sChassisID = "" Then
		If Request.Form("sy_zaiko") <> "" Then
			sChassisID = Request.Form("sy_zaiko")	'手入力
		ElseIf Request.Form("SELECT1")  <> "No0" Then
			sChassisID = Request.Form("SELECT1")		'在庫選択
		Else
			sChassisID = Request.Form("SELECT2")		'非在庫選択
		End If
	End If

	'シャーシ表示情報の取得
	sDispChassis1 = ""
	sDispChassis2 = ""
	sChk1 = ""
	sChk2 = ""
	sql = "SELECT * FROM sChassis" & _
			" WHERE ChassisId = '" & sChassisID & "'"
	rsd.Open sql, conn, 0, 1, 1

	If Not rsd.EOF Then
		sDispChassis1 = sChassisID				'指定シャーシID
		sDispChassis2 = trim(rsd("PlateNo"))	'プレート番号
		If rsd("NotDelFlag") = "Y" Then
			sChk1 = "1"		'搬出を載せないシャーシ
		End If
		If rsd("NightFlag") = "Y" Then
			sChk2 = "1"		'夕積みシャーシ
		End If
	End If
	rsd.Close
	
	sPlateNo = sDispChassis1 & "　" & sDispChassis2

%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title26.gif" width="236" height="34"><p>
</center>

		<font face="ＭＳ ゴシック">
   
<center>
<form  METHOD="post"  NAME="UPLOAD1" ACTION="SYWB032.asp?TDATE=<%=sYMD%>" onSubmit="return ClickSend()">
<table border="1" width="420"  >
<b><font color=#000080>対象シャーシ</font></b>
		<tr bgcolor=#ffff99><td><br>　<%=sPlateNo%>
		<br>　　　　　　　　　　　　　　　<input type="button" value="シャーシの詳細確認" id=submit5 name=submit5 onclick="return ClickSend2(this)">
		<br>		<INPUT TYPE=hidden NAME="sy_zaiko" VALUE=<%=sDispChassis1%>>
					<INPUT TYPE=hidden NAME="SELECT1" VALUE="No0">
					<INPUT TYPE=hidden NAME="SELECT2" VALUE="No0">
		</td></tr>
</table><br><br>
<table border="1" width="420">
<b><font color=#000080>処理選択</font></b>	<tr bgcolor=#ccffcc><td><br>
				<%	if sChk1 = "1" then %>
<INPUT TYPE=checkbox NAME="check1" checked onClick="f1=!f1">搬出コンテナを載せない<br>
				<%	else	%>
<INPUT TYPE=checkbox NAME="check1" onClick="f1=!f1">搬出コンテナを載せない<br>
				<%	end if
					if sChk2 = "1" then%>
<INPUT TYPE=checkbox NAME="check2" checked onClick="f2=!f2">夕積のみ載せる<br>
				<%	else	%>
<INPUT TYPE=checkbox NAME="check2" onClick="f2=!f2">夕積のみ載せる<br>
				<%	end if	%>
<INPUT TYPE=checkbox NAME="check3" onClick="f3=!f3">グループ変更
			<SELECT NAME="SELECT3">
				<%	sql = "SELECT * FROM sMGroup" & _
						  " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
					rsd.Open sql, conn, 0, 1, 1
					do while not rsd.EOF
						%><OPTION VALUE=<%=RTRIM(rsd("GroupID"))%>><%=rsd("GroupName")%>
						<%rsd.MoveNext
					loop
					rsd.Close
'ほかのグループ
					sql = "SELECT * FROM sMGroup" & _
						  " WHERE RTRIM(GroupID) <> '" & sGrpID & "'"
					rsd.Open sql, conn, 0, 1, 1
					do while not rsd.EOF
						%><OPTION VALUE=<%=RTRIM(rsd("GroupID"))%>><%=rsd("GroupName")%>
						<%rsd.MoveNext
					loop
					rsd.Close %>
			</SELECT><br><br>
			<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
	</td></tr>
</table>
</center>
<br>
<center>
<table border=0>
		<td><input type="submit"  value="　実行　" id=submit4 name=submit4></td>
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