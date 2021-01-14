<!--<%@ LANGUAGE="VBScript" %>-->
<%
'Option Explicit
%>
<!--#include file="Common.inc"-->
<!--#include file="Common2.inc"-->
<!--#include file="Sywb060.inc"-->
<html>
<head>
<title>空バン予約画面(ＶＰブッキング)</title>
<SCRIPT LANGUAGE="JavaScript">
<!---
//--->
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
<%
	Dim sYMD, sYMD_OLD, Idx
	Dim conn, rsd
	Dim sName
	Dim sUsrID, sGrpID, sUsrName, sGrpName, sOperator
	Dim sVPBook, sErrMsg, sVPLast

	Dim sContType_M(14), sTerminal_M(49), sMaterial_M(9)

	'指定日付取得(次画面引継)
	sYMD_OLD = TRIM(Request.QueryString("YMD"))
	sHH = Mid(sYMD_OLD, 9, 2)
	sYMD = Left(sYMD_OLD, 8)
	'作業時間帯(次画面引継)
	sName = TRIM(Request.QueryString("NAME"))
	'選択ＣＹ／ＶＰ
	sVPBook = TRIM(Request.Form("VPBookNo"))

	'ＤＢ接続
	Call ConnectSvr(conn, rsd)

	'ユーザ情報の取得
	Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)
	If sGrpID = "" Then
		Response.Write "ユーザが登録されていません。(" & sUsrID & ")"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.End
	End If

	'01/12/05 空バン有積み指定有無取得
	sVPLast = GetEnv(conn, rsd, "VPLastFlag")

	If sVPLast = "N" and sName = "夕積指定" Then
		sErrMsg = "空バン予約の夕積指定はできません"
	Else

		'ユーザ情報の取得
		Call GetUserInfo(conn, rsd, sUsrID, sGrpID, sUsrName, sGrpName, sOperator)

		'コンテナタイプ取得
		Call GetContType(conn, rsd, sContType_M)

		'コンテナ材質取得
		Call GetMaterial(conn, rsd, sMaterial_M)

		'対象ＶＰ取得
		Call GetTerminal(conn, rsd, sTerminal_M)

		'ブッキング予約チェック処理１（作業日のチェック）
		Call VPBookCheck1(conn, rsd, sUsrID, sGrpID, _
				sYMD, sHH, sVPBook,	sErrMsg)

		If sErrMsg = "" Then
		'ＶＰブッキングレコード等の読み込みを行う
			Call GetVPBooking1(conn, rsd, sVPBook, sErrMsg)		
		End If
	End If

	If sErrMsg = "" then
		Idx = 1
		rsd.MoveFirst
		Do Until rsd.EOF
%>
			function ChkGo<%=Idx%>() {
				deliverto=document.form0.DeliverTo.value
				if ( !ChkChara(deliverto) ) {
					window.alert("コンテナ搬出先は半角ローマ字で入力してください。");
					return;
				}
				str="SYWB061.asp?YMD=<%=sYMD_OLD%>&NAME=<%=sName%>&VSL=<%=trim(rsd("VslCode"))%>&VOY=<%=trim(rsd("Voyage"))%>&LINE=<%=trim(rsd("LineCode"))%>&BOOK=<%=trim(rsd("BookNo"))%>&TERM=<%=trim(rsd("Terminal"))%>&SIZE=<%=trim(rsd("ContSize"))%>&TYPE=<%=trim(rsd("ContType"))%>&HIGHT=<%=trim(rsd("ContHeight"))%>&MATERIAL=<%=trim(rsd("Material"))%>"
				if ( confirm('予約しますか？') )
				{
					location.href=str + "&DELIVERTO=" + deliverto;
				}
			}
<%
			rsd.MoveNext
			Idx = Idx + 1
		Loop
		rsd.MoveFirst
	End If
%>
</SCRIPT>
</head>

<body>
<img border="0" src="image/title01.gif" width="311" height="42">
<br><br>

<%
'データがない場合の処理

	if sErrMsg <>  "" then
%>		<center>
			<p><img border="0" src="image/title34.gif" width="236" height="34"><p>
			<table border="1">   
				<tr ALIGN=middle>
					<td width="120" bgcolor ="#e8ffe8">作業時間</td>
					<td width="360" ><%=ChgYMDStr2(sYMD)%>　<%=sName%></td>
				</tr>
			</table>
			<br>
			<table border="1" width="500"  >
				<tr><td width="160" bgcolor="#cccc99">ブッキング番号</td>
					<td><%=sVPBook%></td>
				</tr>
			</table><br><%
			Response.Write "<center><FONT color=Red><U>" & "（結果）：不可　" & sErrMsg & "</U></FONT></center>"
			%><br>
		</center>

		<br>     
		<center>
			<table border=0>
			    <form id=form1 name=form1>
			    <input type="button" value="　戻る　" onclick="history.back()"  id=button1 name=button1>
				</form>
			</table>
		</center>
<%	Else	%>
		<center>
		<p><img border="0" src="image/title33.gif" width="236" height="34"><p>
		<table border="0">   
			<tr ALIGN=middle>
				<td><font size=5><u><%=ChgYMDStr2(sYMD)%>　<%=sName%>　空バン便予約</u></font></td>
			</tr>
			<tr></tr><tr></tr><tr></tr><tr></tr><tr></tr>
			<tr ALIGN=middle>
				<th><font size=4>ブッキング番号・・・<%=trim(rsd("BookNo"))%></font></th>
			</tr>
			<tr></tr><tr></tr><tr></tr>
		<table border="0" bgcolor ="#FFFFBB" width="420">   
			<tr ALIGN=middle>
				<td>コンテナ搬出先を入力後目的のものを選択してください</td>
			</tr>
			<tr ALIGN=middle>
			    <form id=form0 name=form0>
				<td>
				コンテナ搬出先：　
				<INPUT NAME="DeliverTo" SIZE="50" MAXLENGTH="30" STYLE="ime-mode:disabled"><br>
				(半角ローマ字で入力してください)</td>
				</form>
			</tr>
		</table>

		</table>
		<br>

		<table border="1">   
			<tr ALIGN=middle>
				<td width="50" BGCOLOR=#7FFFD4></td>
				<th width="150" BGCOLOR=#7FFFD4>対象バンプール</th>
				<th width="50" BGCOLOR=#7FFFD4>サイズ</th>
				<th width="100" BGCOLOR=#7FFFD4>タイプ</th>
				<th width="50" BGCOLOR=#7FFFD4>高さ</th>
				<th width="100" BGCOLOR=#7FFFD4>材質</th>
				<th width="50" BGCOLOR=#7FFFD4>本船</th>
				<th width="50" BGCOLOR=#7FFFD4>次航</th>
				<th width="50" BGCOLOR=#7FFFD4>船社</th>
			</tr>

<%
			Idx = 1
			rsd.MoveFirst
			Do Until rsd.EOF	%>
					<tr ALIGN=middle>
						<td><font size=4>
<!---						<A href="SYWB061.asp?YMD=<%=sYMD_OLD%>&
												NAME=<%=sName%>&
												VSL=<%=trim(rsd("VslCode"))%>&
												VOY=<%=trim(rsd("Voyage"))%>&
												LINE=<%=trim(rsd("LineCode"))%>&
												BOOK=<%=trim(rsd("BookNo"))%>&
												TERM=<%=trim(rsd("Terminal"))%>&
												SIZE=<%=trim(rsd("ContSize"))%>&
												TYPE=<%=trim(rsd("ContType"))%>&
												HIGHT=<%=trim(rsd("ContHeight"))%>&
												MATERIAL=<%=trim(rsd("Material"))%>" onclick="JavaScript:return confirm('予約しますか？')"><%=Idx%></a>
--->
<A href="JavaScript:ChkGo<%=Idx%>();"><%=Idx%></a>

						</font></td>
						<td><%=SetTerminal(rsd("Terminal"), sTerminal_M)%></td>
						<td><%=rsd("ContSize")%></td>
						<td><%=SetContType(rsd("ContType"), sContType_M)%></td>
						<td><%=rsd("ContHeight")%></td>
						<td><%=SetMaterial(rsd("Material"), sMaterial_M)%></td>
						<td><%=rsd("VslCode")%></td>
						<td><%=rsd("Voyage")%></td>
						<td><%=rsd("LineCode")%></td>
					</tr>
<%				rsd.MoveNext
				Idx = Idx + 1
			Loop
%>				
		</table>
		<br>
		<center>
			目的のものを選択してください（左の番号をクリックして下さい）
		    <form id=form1 name=form1>
		    <input type="button" value="　中止　" onclick="history.back()"  id=button1 name=button1>
			</form>
		</center>

<%
	End If 
	conn.Close
%>

</body>
</html>
