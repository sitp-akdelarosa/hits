<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>前受け予約申請画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
function ClickSend() {

	/* 搬出入タイプを取得 */
	Type1 = Type2 = Type3 = Type4 = 0
	if (document.SEND.RDType1.checked) {
		Type1 = 2
	}
	if (document.SEND.RDType2.checked) {
		Type2 = 2
	}
	if (document.SEND.RDType3.checked) {
		Type3 = 2
	}
	if (document.SEND.RDType4.checked) {
		Type4 = 2
	}

	if (Type1 == 0 &&
		Type2 == 0 &&
		Type3 == 0 &&
		Type4 == 0) {
		window.alert("種類を入力してください。");
		return false;
	}

	if (ChkSend("予約１", Type1, 
				document.SEND.ContNoRec1.value, 
				document.SEND.BKNo1.value, 
				document.SEND.ContSizeRec1.value, 
				document.SEND.checkA1.checked, 
				document.SEND.checkB1.checked, 
				document.SEND.checkC1.checked, 
				document.SEND.ReceiveFrom1.value) &&
		ChkSend("予約２", Type2, 
				document.SEND.ContNoRec2.value, 
				document.SEND.BKNo2.value, 
				document.SEND.ContSizeRec2.value, 
				document.SEND.checkA2.checked, 
				document.SEND.checkB2.checked, 
				document.SEND.checkC2.checked, 
				document.SEND.ReceiveFrom2.value) &&
		ChkSend("予約３", Type3, 
				document.SEND.ContNoRec3.value, 
				document.SEND.BKNo3.value, 
				document.SEND.ContSizeRec3.value, 
				document.SEND.checkA3.checked, 
				document.SEND.checkB3.checked, 
				document.SEND.checkC3.checked, 
				document.SEND.ReceiveFrom3.value) &&
		ChkSend("予約４", Type4, 
				document.SEND.ContNoRec4.value, 
				document.SEND.BKNo4.value, 
				document.SEND.ContSizeRec4.value, 
				document.SEND.checkA4.checked, 
				document.SEND.checkB4.checked, 
				document.SEND.checkC4.checked, 
				document.SEND.ReceiveFrom4.value)) { 
		return true;
	}
	return false;
}

function ChkSend(Name, RDType, ContNoRec, BKNo, ContSizeRec, 
					ChkA, ChkB, ChkC, ReceiveFrom) {
	if (RDType == 0) {					/*選択なし*/
		if (ContNoRec != "" || BKNo != "" || ContSizeRec != "BL" || ReceiveFrom != "" ||
			ChkA || ChkB  || ChkC) {
				window.alert(Name + "の種類を選択してください。" + DeliverTo);
				return false;
		}
	}

	if (RDType == 2) {	/* 搬入の場合 */
		if (ContNoRec == "") {
			window.alert(Name + "の搬入コンテナ番号を入力してください。");
			return false;
		}
		if (BKNo == "") {
			window.alert(Name + "の搬入ブッキング番号を入力してください。");
			return false;
		}
		if (ContSizeRec == "BL") {
			window.alert(Name + "の搬入コンテナサイズを入力してください。");
			return false;
		}
		if (!ChkChara(ReceiveFrom)) {
			window.alert(Name + "のコンテナ搬入元は英字で入力して下さい。");
			return false;
		}
		if (ChkA && ChkB) {
			window.alert(Name + "の『搬出を載せない』と『夕積のみ載せる』が矛盾しています。");
			return false;
		}
	}
	return true;
}
function ChkChara(str) {
	/* 半角英字数字のみ許可 */
	sWk = str.toUpperCase()	/* 大文字変換 */
	for (i = 0; i < sWk.length; i++) {
		if (!((sWk.charAt(i) >= "A" && sWk.charAt(i) <= "Z") ||
 		      (sWk.charAt(i) >= "0" && sWk.charAt(i) <= "9"))) {
			return false;
		}
	}
	return true;
}
function ClickSend1(go) {
	/*クリア処理 予約１*/
	document.SEND.RDType1.checked = false

	/*クリア処理 予約１*/
	document.SEND.ContNoRec1.value = "" 
	document.SEND.BKNo1.value = ""  
	document.SEND.ContSizeRec1.value = "BL"  
	document.SEND.checkA1.checked = false 
	document.SEND.checkB1.checked = false 
	document.SEND.checkC1.checked = false 
	document.SEND.ReceiveFrom1.value = ""  
}
function ClickSend2(go) {
	/*クリア処理 予約２*/
	document.SEND.RDType2.checked = false

	/*クリア処理 予約２*/
	document.SEND.ContNoRec2.value = "" 
	document.SEND.BKNo2.value = ""  
	document.SEND.ContSizeRec2.value = "BL"  
	document.SEND.checkA2.checked = false 
	document.SEND.checkB2.checked = false 
	document.SEND.checkC2.checked = false 
	document.SEND.ReceiveFrom2.value = ""  
}
function ClickSend3(go) {
	/*クリア処理 予約３*/
	document.SEND.RDType3.checked = false
	/*クリア処理 予約３*/
	document.SEND.ContNoRec3.value = "" 
	document.SEND.BKNo3.value = ""  
	document.SEND.ContSizeRec3.value = "BL"  
	document.SEND.checkA3.checked = false 
	document.SEND.checkB3.checked = false 
	document.SEND.checkC3.checked = false 
	document.SEND.ReceiveFrom3.value = ""  
}
function ClickSend4(go) {
	document.SEND.RDType4.checked = false
	/*クリア処理 予約４*/
	document.SEND.ContNoRec4.value = "" 
	document.SEND.BKNo4.value = ""  
	document.SEND.ContSizeRec4.value = "BL"  
	document.SEND.checkA4.checked = false 
	document.SEND.checkB4.checked = false 
	document.SEND.checkC4.checked = false 
	document.SEND.ReceiveFrom4.value = ""  
}

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sHH, sHHName, sTerm_Name, sTerm_CD

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("Name"))
	sTerm_Name = Trim(Request.QueryString("Term_Name"))		'VP対応
	sTerm_CD = Trim(Request.QueryString("Terminal"))		'VP対応
%>

<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title36.gif" width="236" height="34"><p>
</center>
<center>
<table border="1">   
	<tr ALIGN=middle>
		<td width="200" bgcolor ="#e8ffe8">作業時間</td>
		<td width="360" bgcolor ="#ffffff"><%=ChgYMDStr2(sYMD)%>　<%=sHHName%></td>
	</tr>
	<tr ALIGN=middle>
		<td width="200" bgcolor ="#e8ffe8">搬入先</td>
		<td width="360" bgcolor ="#ffffff"><%=sTerm_Name%></td>
	</tr>
</table>
<br>
<center><font color="#ff0000"><small>
（注意）コンテナ搬入元は半角ローマ字で入力してください
</small></font>
</center>
<font face="ＭＳ ゴシック">
<!--	<form  METHOD="post" NAME="SEND" ACTION="SYWB012.asp?TDATE=<%=sYMD%>&HH=<%=sHH%>&HHNAME=<%=sHHName%>" onSubmit="return ClickSend()"> -->
<form  METHOD="post" NAME="SEND" ACTION="SYWB063.asp?TDATE=<%=sYMD%>&HH=<%=sHH%>&HHNAME=<%=sHHName%>
&Term_Name=<%=sTerm_Name%>&Terminal=<%=sTerm_CD%>" 
			onSubmit="return ClickSend()">
<center>
<%
	Dim idx, sRDType
	Dim sContNoRec, sBKNo, sContSizeRec, bChkA, bChkB, bChkC
	Dim sContNoDel, sChID, sBLNo, sContSizeDel, sDeliverTo, sReceiveFrom
	Dim sWk

	for idx = 1 to 4
%>
	<table border="0" width="700" bgcolor ="#ffffff">  
		<TR><th align=left><font color="#00008B">＜予約<%=idx%>＞</font></th></TR>
	</table>
	
	<table border="1" width="700" bgcolor ="#ffffff" cellpadding="3">  
		<tr><th bgcolor ="#40E0D0">種類</th>
			<td COLSPAN=2 bgcolor ="#ffffcc">
<%			sRDType = TRIM(Request.QueryString("sRDType" & CStr(idx)))
			If sRDType	= "" or sRDType	= null	then %>
				<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC">搬入　
<%			Else														
				Select case  sRDType									
					case	"REC"	%>
						<INPUT TYPE="radio" NAME="RDType<%=CStr(idx)%>" VALUE="REC" Checked>搬入　
<%				End Select												
			End If							%>
		</tr>
		<tr><th bgcolor ="#40E0D0" ROWSPAN=2>搬入時</th>
<%			If	sRDType = "" OR sRDType = "DEL" Then 							%>
				<td bgcolor=#cccc99>
					コンテナ番号　(必須)<INPUT TYPE="text" NAME="ContNoRec<%=CStr(idx)%>" SIZE="18" MAXLENGTH="12"><br>
				    ブッキング番号(必須)<INPUT TYPE="text" NAME="BKNo<%=CStr(idx)%>" SIZE="28" MAXLENGTH="20"><br>
					コンテナサイズ(必須)<SELECT NAME="ContSizeRec<%=CStr(idx)%>" size=0>
									<OPTION VALUE="BL" selected>
									<OPTION VALUE="20" >20
									<OPTION VALUE="40" >40</OPTION>
								</SELECT></td>
				<td bgcolor ="#ffffcc">
					<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>"> 搬出を載せない(選択)<br>
					<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>"> 夕積のみ載せる(選択)<br>
					<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>"> 20/40兼用シャーシ(選択)
				</td>
				</tr>
				<tr>
				<td colspan=3 bgcolor ="#ffffcc">
				(共通)既知の場合・・・コンテナ搬入元　
					<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				</td>
				</tr>

<%			Else	
				sContNoRec = TRIM(Request.QueryString("sContNoRec" & CStr(idx)))
				sBKNo = UCASE(TRIM(Request.QueryString("sBKNo" & CStr(idx))))
				sContSizeRec = UCASE(TRIM(Request.QueryString("sContSizeRec" & CStr(idx))))
				bChkA = Request.QueryString("bChkA" & CStr(idx))
				bChkB = Request.QueryString("bChkB" & CStr(idx))
				bChkC = Request.QueryString("bChkC" & CStr(idx))
				sReceiveFrom = Leftb(TRIM(Request.QueryString("sReceiveFrom" & CStr(idx))),30)
%>
				<td bgcolor=#cccc99>
					コンテナ番号　(必須)<INPUT TYPE="text" NAME="ContNoRec<%=CStr(idx)%>" Value="<%=sContNoRec%>" SIZE="18" MAXLENGTH="12"><br>
				    ブッキング番号(必須)<INPUT TYPE="text" NAME="BKNo<%=CStr(idx)%>" Value="<%=sBKNo%>" SIZE="28" MAXLENGTH="20"><br>
					コンテナサイズ(必須)<SELECT NAME="ContSizeRec<%=CStr(idx)%>" size=0>
<%					Select Case	sContSizeRec			
						Case	"20"	%>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" selected>20
								<OPTION VALUE="40" >40</OPTION>
							</SELECT></td>
<%						Case	"40"	%>
								<OPTION VALUE="BL" >
								<OPTION VALUE="20" >20
								<OPTION VALUE="40" selected>40</OPTION>
							</SELECT></td>
<%					End Select			%>
				<td bgcolor ="#ffffcc">
<%					If bChkA = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>" Checked> 搬出を載せない(選択)<br>
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkA<%=CStr(idx)%>"> 搬出を載せない(選択)<br>
<%					End If					%>

<%					If bChkB = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>" Checked> 夕積のみ載せる(選択)<br>
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkB<%=CStr(idx)%>"> 夕積のみ載せる(選択)<br>
<%					End If					%>
					
<%					If bChkC = "True" Then	%>
						<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>" Checked> 20/40兼用シャーシ(選択)
<%					Else					%>
						<INPUT TYPE=checkbox NAME="checkC<%=CStr(idx)%>"> 20/40兼用シャーシ(選択)
<%					End If					%>
				</td>
				</tr>
				<tr>
				<td colspan=3 bgcolor ="#ffffcc">
<%				If sReceiveFrom <> "" Then	%>
				(共通)既知の場合・・・コンテナ搬入元　<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" Value="<%=sReceiveFrom%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
<%				Else%>
				(共通)既知の場合・・・コンテナ搬入元　<INPUT NAME="ReceiveFrom<%=CStr(idx)%>" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
<%				End If	%>
				</td>
				</tr>
<%			End If	%>


	</table>
	
	<table border=0 width="700" bgcolor ="#ffffff">
		<tr><td align=center><font color="#ff0000"><small>
		（注意）コンテナ搬入元はダイヤ決定までに入力がない場合予約キャンセルとなります
		</small></font></td></tr>
	</table>

	<table border=0 width="700" bgcolor ="#ffffff"><tr align=right><td>
		<input type="submit" value="　全体送信　" id=submit4 name=submit4>
		<input type="button" value="予約<%=CStr(idx)%>ｸﾘｱ" id=submit4 name=submit4 onclick="return ClickSend<%=CStr(idx)%>(this)">
		</td></tr>
	</table>

<%	next	%>
</center>

<p>

<br>

<center>
<table border=0>
	</form>

    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="　　中止　　" id=submit4 name=submit4></td>
	</form>
</table>
</center>

</body>     
</html>     
