<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
     'セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {

	return (ChkSend(document.con.CntnrNo.value, 
				document.con.Year.value, 
				document.con.Month.value, 
				document.con.Day.value));
}
// 入力チェック
function ChkSend(sCntnrNo, sYear, sMonth, sDay ) {
	if (sCntnrNo == "") {	/* コンテナNo.未入力チェック */
		window.alert("コンテナNo.が未入力です。");
		return false;
	}

	if (sYear == "" ||  sMonth == "" || sDay == "") {
		window.alert("CY搬入日が未入力です。");
		return false;
	}

	if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* 年のチェック */
		window.alert("CY搬入日の年の入力が不正です。");
		return false;
	}
	if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* 月のチェック */
		window.alert("CY搬入日の月の入力が不正です。");
		return false;
	}
	if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* 日のチェック */
		window.alert("CY搬入日の日の入力が不正です。");
		return false;
	}

	if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
	   (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
		window.alert("CY搬入日の日の入力が不正です。");
		return false;
	}

	return true;
}
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/kaika2t.gif" width="506" height="73"></td>
					<td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
				</tr>
				<tr>
					<td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
<table border=0 cellpadding=0 cellspacing=0><tr><td align=left>
				<table>
					<tr> 
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>コンテナ情報入力</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
<center>
				<table>
					<tr>
						<td>下記の項目を入力の上、送信ボタンをクリックして下さい。</td>
					</tr>
				</table>
				<FORM NAME="con" METHOD="post" action="nyuryoku-ex-syori.asp" onSubmit="return ClickSend()">
								<table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
											コンテナNo.</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=170>
												<input type="text" name="CntnrNo" size="20" maxlength="12">
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
												<font size=1 color="#2288ff">[ 半角英数 ]</font>
											</td>
										  </tr>
										</table>
											
										</td>
									</tr>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
  											CY搬入日</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=170>
												<input type=text name="Year" value="<%=Year(Now)%>" size=4 maxlength="4">年
												<input type=text name="Month" value="<%=Month(Now)%>" size=2 maxlength="2">月
												<input type=text name="Day" size=2 maxlength="2">日　
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
												<font size=1 color="#2288ff">[ 半角数値 ]</font>
											</td>
										  </tr>
										</table>
											&nbsp;&nbsp;&nbsp;<font size=-1>（例） 2002年2月25日</font>
										</td>
									</tr>
								</table>
								<br>
								<input type=submit value=" 送  信 " name="リセット">
				</form>
</center>
				<table>
					<tr> 
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>CSVファイル転送</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
<center>
				<table border="0" cellspacing="1" cellpadding="2">
					<tr> 
						<td> 
							<p>情報をファイル転送する場合はここをクリック</p>
						</td>
						<td>…</td>
						<td><a href="nyuryoku-ex-csv.asp">CSVファイル転送</a></td>
					</tr>
					<tr> 
						<td>CSVファイル転送についての説明はここをクリック</td>
						<td>…</td>
						<td><a href="help09.asp">ヘルプ</a></td>
					</tr>
				</table>
</center>
</td></tr></table>

				<br>
          　		<br>
			</center>
            <br>
		</td>
	</tr>
	<tr>
		<td valign="bottom"> 
<%
    DispMenuBar
%>
		</td>
	</tr>
</table>
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "nyuryoku-kaika.asp"
%>
</body>
</html>
<%
    ' 海貨CY搬入日指示
    WriteLog fs, "4003","海貨入力CY搬入日", "00",","
%>