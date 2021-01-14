<%@ LANGUAGE="VBScript" %>
<%
Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="SYWB031A.inc"-->
<!--#include file="SYWB031B.inc"-->
<!--#include file="SYWB031C.inc"-->
<!--#include file="SYWB031D.inc"-->
<html>

<head>
<title>搬出入予約変更結果画面</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
/* 戻るのクリック */
function ClickBack() {
	history.back()
}

</SCRIPT>
</head>

<body>
<%
	Dim sYMD, sOpeNo, sCmd, sHH, sChgOpeNo, sCmdName, sDelFlag, sStatus
	Dim sRecDel, sSend, sVPLast
	Dim conn, rsd
	Dim sErrMsg

	'指定日付取得
	sYMD = TRIM(Request.QueryString("YMD"))

	'作業番号取得
	sOpeNo = TRIM(Request.QueryString("OPENO"))

	'実行コマンド取得
	sCmd = TRIM(Request.QueryString("CMD"))
	If     sCmd = "DEL" Then	'削除
		sCmdName = "削除"
	ElseIf sCmd = "MOV" Then	'移動
		sCmdName = "移動"
	ElseIf sCmd = "CHG" Then	'交換
		sCmdName = "交換"
	Else						'変更
		If TRIM(Request.Form("DeliverTo")) <> "" Then
			sCmdName = "搬出先変更"
		Else
			sCmdName = "搬入元変更"
		End If
	End If

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'申請情報の取得（指定作業番号）
	Call GetAppInfoOpeNo(conn, rsd, CLng(sOpeNo))
	sHH      = Trim(rsd("Term"))
	sDelFlag = Trim(rsd("DelFlag"))
	sStatus  = Trim(rsd("Status"))
	rsd.Close

	'01/12/05 空バン有積み指定有無取得
	sVPLast = GetEnv(conn, rsd, "VPLastFlag")

	'予約日、予約時間帯が有効かチェック
	sErrMsg = ""
	If sDelFlag <> "Y" And sStatus <> "03" Then
		Call CheckAppWorkDate(conn, rsd, sYMD, sHH, sErrMsg) 
	End If
	sHH = ""
	If sErrMsg = "" Then

		sChgOpeNo = ""
		sErrMsg = ""

		If     sCmd = "DEL" Then	'削除
			Call UpdOpeDel(conn, rsd, sOpeNo, sErrMsg)
		ElseIf sCmd = "MOV" Then	'移動
			sHH = TRIM(Request.Form("SELECT"))
			If sVPLast = "N"  AND sHH = "B" Then	'01/12/05
				sErrMsg = "空バン予約の夕積指定への移動はできません"
			Else
				Call UpdOpeMov(conn, rsd, sOpeNo, sYMD, sHH, sErrMsg)
			End If
		ElseIf sCmd = "CHG" Then	'交換
			sChgOpeNo = TRIM(Request.Form("CHGOPE"))
			Call UpdOpeChg(conn, rsd, sOpeNo, sYMD, sChgOpeNo, sErrMsg)
		Else						'変更
			If TRIM(Request.Form("DeliverTo")) <> "" Then
				sSend = TRIM(Request.Form("DeliverTo"))
			Else
				sSend = TRIM(Request.Form("ReceiveFrom"))
			End If
			Call UpdOpeUpd(conn, rsd, sOpeNo, sSend, sErrMsg)
		End If
	End If

	'ＤＢ切断
	conn.Close
%>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>
<center>
<p><img border="0" src="image/title35.gif" width="236" height="34"><p>
</center>
<br><br>

<font face="ＭＳ ゴシック">
<center>
	<table border="1" width="300" >
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">動作</td><td><%=sCmdName%></td></tr>
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">予約番号</td><td><%=sOpeNo%></td></tr>
<%
	If sChgOpeNo <> "" Then
%>
		<tr ALIGN=middle><td width="130" bgcolor ="#e8ffe8">交換相手</td><td><%=sChgOpeNo%></td></tr>
<%
	End If
%>
	</table>

	<br><br>
	<B><U><font color=#ff0000>
<%
	If sErrMsg <> "" Then
%>
	結果：<%=sErrMsg%><br>
<%
	Else
%>
	結果：OK<br>
<%
	End If
%>
	</font></U></B>
<br><br>
<%
	If sErrMsg <> "" Then
%>
		<input type="button" value="　戻る　" onClick="ClickBack()">
<%
	End If
%>
</center>
<br><br>
<FORM NAME="SEND">
	<INPUT TYPE=hidden NAME="YMD" VALUE=<%=sYMD%>>
</FORM>
<%
	If sErrMsg = "" Then
%>
<SCRIPT LANGUAGE="JavaScript">
	location.replace("SYWB013.asp?TDATE=" + document.SEND.YMD.value);
</SCRIPT>
<%
	End If
%>
</body>     
</html>     
