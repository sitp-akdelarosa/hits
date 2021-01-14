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

	return (ChkSend(document.con.ContNo.value, 
					document.con.BLNo.value, 
					document.con.Year.value, 
					document.con.Month.value, 
					document.con.Day.value, 
					document.con.Hour.value, 
					document.con.Min.value));
}
// 入力チェック
function ChkSend(sContNo, sBLNo, sYear, sMonth, sDay, sHour, sMin ) {
	if (sContNo == "" && sBLNo == "") {	/* コンテナNo.またはBL番号未入力チェック */
		window.alert("コンテナNo.またはBL番号が未入力です。");
		return false;
	}

	if (sContNo != "" && sBLNo != "") {	/* コンテナNo.またはBL番号未入力チェック */
		window.alert("コンテナNo.またはBL番号のどちらかを入力して下さい。");
		return false;
	}

	if (sYear == "" ||  sMonth == "" || sDay == "" || sHour == "" || sMin == "") {
		window.alert("届け時刻が未入力です。");
		return false;
	}

	if (!(sYear > 0 || sYear <= 0)|| sYear < 1990 || sYear > 2100 ) {	/* 年のチェック */
		window.alert("届け時刻の年の入力が不正です。");
		return false;
	}
	if (!(sMonth > 0 || sMonth <= 0)|| sMonth < 1 || sMonth > 12 ) {	/* 月のチェック */
		window.alert("届け時刻の月の入力が不正です。");
		return false;
	}
	if (!(sDay > 0 || sDay <= 0)|| sDay < 1 || sDay > 31  ) {		/* 日のチェック */
		window.alert("届け時刻の日の入力が不正です。");
		return false;
	}

	if (!(sHour > 0 || sHour <= 0)|| sHour < 0 || sHour > 24  ) {	/* 時のチェック */
		window.alert("届け時刻の時の入力が不正です。");
		return false;
	}

	if (!(sMin > 0 || sMin <= 0)|| sMin < 0 || sMin > 59  ) {		/* 分のチェック */
		window.alert("届け時刻の分の入力が不正です。");
		return false;
	}

	if (sDay<=0 || sDay>30+((sMonth==4||sMonth==6||sMonth==9||sMonth==11)?0:1) || 
	   (sMonth==2&&sDay>28+(((sYear%4==0&&sYear%100!=0)||sYear%400==0)?1:0)) ){
		window.alert("届け時刻の日の入力が不正です。");
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
					<td rowspan=2><img src="gif/kaika3t.gif" width="506" height="73"></td>
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
						<td align=left>下記の項目を入力の上、送信ボタンをクリックして下さい。<BR>
						コンテナNo.またはBLNo.と、倉庫届け時刻（荷主指示）は必須入力です。
						</td>
					</tr>
				</table>
				<FORM NAME="con" METHOD="post" action="nyuryoku-im-syori.asp" onSubmit="return ClickSend()">
					<table border=0 cellpadding=0>
						<tr> 
							<td align="center"> 
								<table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff" width=100%>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
											コンテナNo.</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=200>
												<input type="text" name="ContNo" size="22" maxlength="12">
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#2288ff">[ 半角英数 ]</font>
											</td>
										  </tr>
										</table>
											
										</td>
									</tr>
									<tr> 
										<td align="center" colspan="2">または、</td>
									</tr>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
											BL No.</font></b></td>
										<td> 
										<table border=0 cellpadding=0 cellspacing=0>
										  <tr>
											<td width=200>
												<input type="text" name="BLNo" size="22" maxlength="20">
											</td>
											<td align=left valign=middle nowrap>
												<font size=1 color="#2288ff">[ 半角英数 ]</font>
											</td>
										  </tr>
										</table>
											
										</td>
									</tr>
								</table>
<BR>
								<table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff" width=100%>
									<tr> 
										<td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
  											倉庫届け時刻</font></b></td>
										<td> 
											<input type=text name="Year" size=4 maxlength="4">年
											<input type=text name="Month" size=2 maxlength="2">月
											<input type=text name="Day" size=2 maxlength="2">日　
											<input type=text name="Hour" size=2 maxlength="2">時
											<input type=text name="Min" size=2 maxlength="2">分
											<table border=0 cellpadding=0 cellspacing=0>
											  <tr>
												<td width=200>
													&nbsp;&nbsp;&nbsp;<font size=-1>（例） 2002年2月25日 15時30分</font>
												</td>
												<td align=left valign=middle nowrap>
													<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
													<font size=1 color="#2288ff">[ 半角数値 ]</font>
												</td>
											  </tr>
											</table>
										</td>
									</tr>
								</table>
								<br>
									<input type=submit value=" 送  信 " name="リセット">
							</td>
						</tr>
					</table>
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
						<td><a href="nyuryoku-im-csv.asp">CSVファイル転送</a></td>
					</tr>
					<tr> 
						<td>CSVファイル転送についての説明はここをクリック</td>
						<td>…</td>
						<td><a href="help10.asp">ヘルプ</a></td>
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
    ' 海貨入力項目選択
    WriteLog fs, "4004","海貨入力実入り倉庫到着時刻","00", ","
%>