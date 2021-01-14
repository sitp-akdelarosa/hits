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
function ClickSend(go) {

	/* 入力チェック */
//未入力チェック
	if  (SELECT1.value == "No0" && SELECT2.value == "No0" && 
		SELECT3.value == "No0" && 
		sy_zaiko.value == "" && 
		sy_Change.value == "") {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
	if ((sy_zaiko.value != ""   && SELECT1.value != "No0") || 
		(sy_zaiko.value != ""   && SELECT2.value != "No0") ||
		(sy_zaiko.value != ""   && sy_Change.value != "" ) ||
		(sy_zaiko.value != ""   && SELECT3.value != "No0") ||
		(SELECT1.value != "No0" && SELECT2.value != "No0") ||
		(SELECT1.value != "No0" && sy_Change.value != "")  ||
		(SELECT1.value != "No0" && SELECT3.value != "No0") ||
		(SELECT2.value != "No0" && sy_Change.value != "")  ||
		(SELECT2.value != "No0" && SELECT3.value != "No0") ||
		(sy_Change.value != ""  && SELECT3.value != "No0")) {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
/*変更手入力*/
	if (sy_zaiko.value != "")  {
		location.href = "SYWB033.asp?sCassis=" + sy_zaiko.value.toUpperCase() + "H" + 
		"&YMD=" + YMD.value + "&OPENO=" + OpeNo.value + "&M_ChassisId=" + M_ChassisId.value
		return true;
	}

/*ＳＹ在庫選択入力*/
	if (SELECT1.value != "No0")  {
		location.href = "SYWB033.asp?sCassis=" + SELECT1.value + "H" + 
		"&YMD=" + YMD.value + "&OPENO=" + OpeNo.value + "&M_ChassisId=" + M_ChassisId.value
		return true;
	}

/*ＳＹ非在庫選択入力*/
	if (SELECT2.value != "No0")  {
		location.href = "SYWB033.asp?sCassis=" + SELECT2.value + "H" + 
		"&YMD=" + YMD.value + "&OPENO=" + OpeNo.value + "&M_ChassisId=" + M_ChassisId.value
		return true;
	}
/*交換手入力*/
	if (sy_Change.value != "")  {
		location.href = "SYWB033.asp?sCassis=" + sy_Change.value.toUpperCase() + "K" + 
		"&YMD=" + YMD.value + "&OPENO=" + OpeNo.value + "&M_ChassisId=" + M_ChassisId.value
		return true;
	}
/*交換選択入力*/
	if (SELECT3.value != "No0")  {
		location.href = "SYWB033.asp?sCassis=" + SELECT3.value + "K" + 
		"&YMD=" + YMD.value + "&OPENO=" + OpeNo.value + "&M_ChassisId=" + M_ChassisId.value
		return true;
	}
}
function ClickSend1(go) {

//未入力チェック
	if  (SELECT1.value == "No0" && SELECT2.value == "No0" && 
		SELECT3.value == "No0" && 
		sy_zaiko.value == "" && 
		sy_Change.value == "") {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}
	if ((sy_zaiko.value != ""   && SELECT2.value != "No0") || 
		(sy_zaiko.value != ""   && SELECT2.value != "No0") ||
		(sy_zaiko.value != ""   && sy_Change.value != "" ) ||
		(sy_zaiko.value != ""   && SELECT3.value != "No0") ||
		(SELECT1.value != "No0" && SELECT2.value != "No0") ||
		(SELECT1.value != "No0" && sy_Change.value != "")  ||
		(SELECT1.value != "No0" && SELECT3.value != "No0") ||
		(SELECT2.value != "No0" && sy_Change.value != "")  ||
		(SELECT2.value != "No0" && SELECT3.value != "No0") ||
		(sy_Change.value != ""  && SELECT3.value != "No0")) {
		window.alert("シャーシＩＤに矛盾があります。");
		return false;
	}


/*変更手入力*/
	if (sy_zaiko.value != "")  {
		location.href = "SYWB023.asp?sCassis=" + sy_zaiko.value.toUpperCase()
		return true;
	}

/*ＳＹ在庫選択入力*/
	if (SELECT1.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + SELECT1.value
		return true;
	}

/*ＳＹ非在庫選択入力*/
	if (SELECT2.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + SELECT2.value
		return true;
	}
/*交換手入力*/
	if (sy_Change.value != "")  {
		location.href = "SYWB023.asp?sCassis=" + sy_Change.value.toUpperCase()
		return true;
	}
/*交換選択入力*/
	if (SELECT3.value != "No0")  {
		location.href = "SYWB023.asp?sCassis=" + SELECT3.value
		return true;
	}
}
//--->
</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sOpeNo, sChassisID							'指定値
	Dim sPlateNo, sSize20Flag, sMixSizeFlag					'シャーシ情報
	Dim conn, rsd, sql										'ＤＢ接続
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator		'ユーザ情報
	Dim sStatus, sDelFlag, sWorkFlag, sLockFlag, sContSize	'申請情報
	Dim i
	Dim sNO, sChk1, sChk2, sChkChassisID			'未使用
	Dim sErr1, sErr2, sErr3							'未使用

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))

	'作業番号取得
	sOpeNo = TRIM(Request.QueryString("OPENO"))

	'シャーシID取得
	sChassisID = TRIM(Request.QueryString("CID"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

	'予約申請情報読み込み
	Call GetAppInfoOpeNo(conn, rsd, sOpeNo)
	'申請情報にシャーシが設定されている場合はそれを優先
	If Trim(rsd("ChassisID")) <> "" Then
		sChassisID = Trim(rsd("ChassisID"))
	End If
	sStatus   = Trim(rsd("Status"))		'状態
	sDelFlag  = Trim(rsd("DelFlag"))	'削除フラグ
	sWorkFlag = Trim(rsd("WorkFlag"))	'作業中フラグ
	sLockFlag = Trim(rsd("LockFlag"))	'ロックフラグ
	sContSize = Trim(rsd("ContSize"))	'コンテナサイズ	'2000/02/22
	rsd.Close

	'シャーシ情報（プレート番号・２０フィートフラグ・２０／４０兼用シャーシ)の取得
	If sChassisID <> "" Then	'2001/02/22 NICS
		sql = "SELECT PlateNo, Size20Flag, MixSizeFlag FROM sChassis" & _
			  " WHERE RTRIM(GroupID) = '" & sGrpID & "'" & _
			    " AND ChassisId = '" & sChassisID & "'"
		rsd.Open sql, conn, 0, 1, 1

		If Not rsd.EOF Then
			sPlateNo     = Trim(rsd("PlateNo"))			'プレート番号
			sSize20Flag  = Trim(rsd("Size20Flag"))		'２０フィートフラグ
			sMixSizeFlag = Trim(rsd("MixSizeFlag"))		'２０／４０兼用シャーシ
		End if
		rsd.Close
		sPlateNo = sChassisID & "　" & sPlateNo
	End If
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title27.gif" width="236" height="34"><p>
</center>

		<font face="ＭＳ ゴシック">
   
<center>
<%	If sChassisID = "" Then	%>
	<INPUT TYPE="hidden" NAME="sy_Change">　
	<INPUT TYPE="hidden" NAME="SELECT3" Value="No0">　
<%	End If	%>

<%	If sChassisID <> "" Then	%>
	<form  METHOD="post"  NAME="UPLOAD0" ACTION="SYWB023.asp?sCassis=<%=sChassisID%>">
	<table border="1" width="500"  >
	<b><font color=#000080>対象シャーシ</font></b><br>
		<tr bgcolor=#ffff99><td><br>
	　<%=sPlateNo%><br>　　　　　　　　　　　　　　　　　　　<input type="submit" value="シャーシの詳細確認">
		<br><br></td></tr>
	</table><br>
	</form>
<%	End If %>
<table border="1" width="500"  >
<%	If sChassisID <> "" Then	%>
<b><font color=#000080>変更・交換相手</font></b><br>
<%	Else	%>
<b><font color=#000080>シャーシ設定</font></b><br>
<%	End If	%>
	<tr bgcolor=#ccffcc><td><br>
<%	If sChassisID <> "" Then	%>
＜変更の場合＞・・・作業にリンクしていないもの<br><br>
<%	End If	%>
　ＳＹ在庫より選択　　<SELECT NAME="SELECT1">
						<OPTION VALUE="No0" >　<%
							if sStatus =  "02" and  sDelFlag = "" _
								and sWorkFlag = ""  and sLockFlag = "" then

								sql = "SELECT ChassisID FROM sChassis"
								sql = sql & " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
								sql = sql & "  AND StackFlag <> ' '"
								sql = sql & "  AND ContFlag = ' '"
								If sContSize = "20" Then	'20f 2/22
									sql = sql & "  AND ( Size20Flag = 'Y' OR MixSizeFlag = 'Y' )"
								Else						'40f 2/22
									sql = sql & "  AND Size20Flag <> 'Y'"
								End IF
								If sChassisID <> "" Then	'2/22
									sql = sql & "  AND ChassisID <> '" & sChassisID & "'"
								End If
								sql = sql & "  Order By ChassisID"

								rsd.Open sql, conn, 0, 1, 1
			
								if not rsd.eof then
									do while not rsd.EOF%>
										<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
										<%rsd.MoveNext
										i = i + 1
									loop
								end if
								rsd.Close
							end if
						%>
				</SELECT><br>
　ＳＹ非在庫より選択　<SELECT NAME="SELECT2">
					<OPTION VALUE="No0" >　<%
							if sStatus =  "02" and  sDelFlag = "" _
								and sWorkFlag = ""  and sLockFlag = "" then
								sql = "SELECT ChassisID FROM sChassis"
								sql = sql & " WHERE RTRIM(GroupID) = '" & sGrpID & "'"
								sql = sql & "  AND StackFlag = ' '"
								sql = sql & "  AND ContFlag = ' '"
								If sContSize = "20" Then
									sql = sql & "  AND ( Size20Flag = 'Y' OR MixSizeFlag = 'Y' )"
								Else
									sql = sql & "  AND Size20Flag <> 'Y'"
								End IF
								If sChassisID <> "" Then	'2/22
									sql = sql & "  AND ChassisID <> '" & sChassisID & "'"
								End If
								sql = sql & "  Order By ChassisID"
								rsd.Open sql, conn, 0, 1, 1
			
								if not rsd.eof then
									do while not rsd.EOF%>
										<OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
										<%rsd.MoveNext
										i = i + 1
									loop
								end if
								rsd.Close
							end if
						%>
				</SELECT><br>
　手入力する場合　　　<INPUT TYPE="text" NAME="sy_zaiko" SIZE="9" MAXLENGTH="5">　
<%	If sChassisID <> "" Then	%><br><br>
＜交換の場合＞・・・作業にリンクしているもの<br><br>
　リストより選択　　　<SELECT NAME="SELECT3">
					<OPTION VALUE="No0" >　<%
							if sStatus =  "02" and  sDelFlag = "" _
								and sWorkFlag = ""  and sLockFlag = "" then
								sql = "SELECT distinct sAppliInfo.ChassisID FROM sAppliInfo, sChassis"
								sql = sql & " WHERE RTRIM(sAppliInfo.GroupID) = '" & sGrpID & "'"
								sql = sql & "  AND sAppliInfo.Status   = '02'"
								sql = sql & "  AND sAppliInfo.DelFlag  = ' '"
								sql = sql & "  AND sAppliInfo.WorkFlag = ' '"
								sql = sql & "  AND sAppliInfo.LockFlag = ' '"
								sql = sql & "  AND sAppliInfo.WorkDate = '" & cdate(ChgYMDStr(sYMD)) & "'"
								If sContSize = "20" Then
									sql = sql & "  AND ( sAppliInfo.Size20Flag = 'Y' OR sAppliInfo.MixSizeFlag = 'Y' )"
								Else
									sql = sql & "  AND sAppliInfo.Size20Flag <> 'Y'"
								End IF
''''''''''''''''sql = sql & "  AND sChassis.Size20Flag = '" & sSize20Flag & "'"
''''''''''''''''sql = sql & "  AND sChassis.MixSizeFlag = '" & sMixSizeFlag & "'"
								sql = sql & "  AND RTRIM(sAppliInfo.ChassisID) <> ''"
								If sChassisID <> "" Then	'2/22
									sql = sql & "  AND RTRIM(sAppliInfo.ChassisID) <> '" & sChassisID & "'"
								End If
								sql = sql & "  AND RTRIM(sAppliInfo.ChassisID) = sChassis.ChassisID"
								sql = sql & "  Order By sAppliInfo.ChassisID"
								rsd.Open sql, conn, 0, 1, 1
								if not rsd.eof then
									do while not rsd.EOF
										%><OPTION VALUE=<%=rsd("ChassisId")%>><%=rsd("ChassisId")%>
										<%rsd.MoveNext
									loop
								end if
								rsd.Close
							end if
							%>
				</SELECT><br>
　手入力する場合　　　<INPUT TYPE="text" NAME="sy_Change" SIZE="9" MAXLENGTH="5">
<%	End If%>

<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
<INPUT TYPE=hidden NAME="OpeNo" VALUE=<%=sOpeNo%>>
<INPUT TYPE=hidden NAME="M_ChassisId" VALUE=<%=sChassisID%>>

<form  METHOD="post"  NAME="UPLOAD2" onclick="return ClickSend1(this)">
　　　　　　　　　　　　　　　　　　　<input type="button" value="シャーシの詳細確認" id=submit5 name=submit5>
</form>

	</td></tr>
</table>
</center>

			<br>
			
<center>
<table border=0>
	<form  METHOD="post"  NAME="UPLOAD1" onclick="return ClickSend(this)">
		<td><input type="button" value="　実行　" id=submit4 name=submit4></td>
	</form>
	<td>　</td>
	<td>　</td>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="　中止　" id=submit4 name=submit4></td>
	</form>
</table>
</center>

</body>     
</html>     
