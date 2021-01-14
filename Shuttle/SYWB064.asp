<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<html>

<head>
<title>空バン予約変更画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* 移動のクリック */
function ClickMov() {
	/* チェックなし */
	return true;
}
/* 交換のクリック */
function ClickChg() {

	if (document.CHG.CHGOPE.value == "") {
		window.alert("交換対象を入力してください。");
		return false;
	}
	if (document.CHG.CHGOPE.value.length != 5 && document.CHG.CHGOPE.value.length != 4) {
		window.alert("交換対象を正しく入力してください。");
		return false;
	}length 
	return true;
}
/* 搬出先のクリック */
function ClickDel() {

		if (!ChkChara(document.UPD.DeliverTo.value)) {
			window.alert("コンテナ搬出先は半角ローマ字で入力して下さい。");
			return false;
		}
}
function ClickRec() {

		if (!ChkChara(document.UPD.ReceiveFrom.value)) {
			window.alert("コンテナ搬入元は半角ローマ字で入力して下さい。");
			return false;
		}
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

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sHH, sHHName, sOpeNo, sAppTerminal
	Dim conn, rsd
	Dim sShtStart, sShtEnd, iSTime, iETime
	Dim iCnt, i
	Dim iTimeCnt, TimeSlot(40), TimeName(40)
	Dim sVPBookNo, sRecDel, sContSize
	Dim sDeliverTo, sReceiveFrom

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD, 9, 2)
	sYMD = Left(sYMD, 8)
	sHHName = TRIM(Request.QueryString("NAME"))

	'作業番号取得
	sOpeNo = TRIM(Request.QueryString("OPENO"))

	'搬出入先取得
	sAppTerminal = TRIM(Request.QueryString("TNAME"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'シャトル運行時間取得
	sShtStart = GetEnv(conn, rsd, "ShtStart")
	sShtEnd   = GetEnv(conn, rsd, "ShtEnd")
	iSTime = CLng(Left(sShtStart, 2))
	iETime = CLng(Left(sShtEnd, 2))
	if Right(sShtEnd, 2) = "00" Then
		iETime = iETime - 1
	End If

	'シャトル運行時間帯計算
	iCnt = 0

	'時間帯の計算
	''午前時間
	For i = iSTime To 11
		TimeSlot(iCnt) = Right("0" & CStr(i), 2)
		If i = iSTime Then
			TimeName(iCnt) = GetTimeSlot(i, CLng(Right(sShtStart, 2)), "S")
		Else
			TimeName(iCnt) = GetTimeSlot(i, "00", "S")
		End If
		iCnt = iCnt + 1
	Next
	''午前指定
	TimeSlot(iCnt) = "12"
	TimeName(iCnt) = "午前"
	iCnt = iCnt + 1
	''午後時間
	For i = 13 To iETime
		TimeSlot(iCnt) = Right("0" & CStr(i), 2)
		If i = iETime Then
			TimeName(iCnt) = GetTimeSlot(i + 1, CLng(Right(sShtEnd, 2)), "E")
		Else
			TimeName(iCnt) = GetTimeSlot(i + 1, "00", "E")
		End If
		iCnt = iCnt + 1
	Next
	''午後指定
	TimeSlot(iCnt) = "A"
	TimeName(iCnt) = "午後"
	iCnt = iCnt + 1
	''夕積指定
	TimeSlot(iCnt) = "B"
	TimeName(iCnt) = "夕積"
	iCnt = iCnt + 1

	iTimeCnt = iCnt		'時間帯数

	'ＤＢ取得
	Call GetAppInfoOpeNo(conn, rsd, int(sOpeNo))
	If Not rsd.EOF Then
		sVPBookNo = Trim(rsd("VPBookNo"))
		sRecDel = "空バン"
		sContSize = Trim(rsd("ContSize")) & "ft"  'コンテナサイズ

'''		sReceiveFrom = Trim(rsd("ReceiveFrom"))			'搬入元
		sDeliverTo = Trim(rsd("DeliverTo"))			'搬出先
	end if
	rsd.Close

	'ＤＢ切断
	conn.Close
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title35.gif" width="236" height="34"><p>
<table border="1">   
	<tr ALIGN=middle>
		<td width="120" bgcolor ="#e8ffe8">対象</td>
		<td width="380" ><%=ChgYMDStr2(sYMD)%>　<%=sHHName%>　<%=sOpeNo%><br>
			<%=sVPBookNo%>　<%=sRecDel%>　<%=sContSize%>　<%=sAppTerminal%></td>
	</tr>
</table>
</center>
<br>

<font face="ＭＳ ゴシック">

<center>

<form  METHOD="post" NAME="UPD" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=UPD&OPENO=<%=sOpeNo%>" onSubmit="return ClickDel()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">搬出先変更</td></tr>
	<td></td>
	
		<td>
			<INPUT NAME="DeliverTo" Value="<%=sDeliverTo%>"	SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled">
			<input type="submit" value="　実行　" id=submit0 name=submit0>
		</td>
	</tr>
</table>
</form>

<form  METHOD="post" NAME="DEL" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=DEL&OPENO=<%=sOpeNo%>">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">削除</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<input type="submit" value="　削除　" id=submit1 name=submit1>
		</td>
	</tr>
</table>
</form>
</center>

<center>
<form  METHOD="post" NAME="MOV" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=MOV&OPENO=<%=sOpeNo%>" onSubmit="return ClickMov()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">移動</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>移動先を指定してください</td>
	</tr>
	<tr>
		<td></td>
		<td><FONT COLOR="4169E1"><SMALL>（午前、午後、夕積予約も可能です）</SMALL></FONT></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<SELECT NAME="SELECT">
<%
	For i = 0 To iTimeCnt - 1
%>
				<OPTION VALUE=<%=TimeSlot(i)%> ><%=TimeName(i)%>
<%
	Next
%>
			</SELECT>
			<input type="submit" value="　実行　" id=submit2 name=submit2>
		</td>
	</tr>
</table>
</form>
</center>

<center>
<form  METHOD="post" NAME="CHG" ACTION="SYWB065.asp?YMD=<%=sYMD%>&CMD=CHG&OPENO=<%=sOpeNo%>" onSubmit="return ClickChg()">
<table border="0" width="600"  >   
	<tr ALIGN=middle><td width="100" bgcolor ="#000080"><FONT COLOR="#ffffff">交換</td>
		<td></td>
	</tr>
	<tr>
		<td></td>
		<td>交換相手の予約番号を指定してください。</td>
	</tr>
	<tr>
		<td></td>
		<td><FONT COLOR="4169E1"><SMALL>（午前、午後、夕積予約の相手も指定できます）</SMALL></FONT></td>
	</tr>
	<tr>
		<td></td>
		<td>
			<input type="text" NAME="CHGOPE" SIZE="9" MAXLENGTH="5">
			<input type="submit" value="　実行　" id=submit3 name=submit3>
		</td>
	</tr>
</table>
</form>
</center>

</table>
</center>

<center>
    <form  METHOD="post"  NAME="CANCEL" ACTION="SYWB013.asp?TDATE=<%=sYMD%>" >
		<td><input type="submit" value="　中止　" id=submit6 name=submit6></td>
	</form>
</center>

</body>
</html>
