<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
     'セッションのチェック
    CheckLogin "nyuryoku-ki.asp"

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
	if (ChkSend(document.con.ContNo.value, 
				document.con.SealNo.value, 
				document.con.Jyuryo.value,
				document.con.SoJyuryo.value)) { 
		return true;
	}
	return false;
}
// 入力チェック
function ChkSend(sContNo, sSealNo, sJyuryo, sSoJyuryo) {

	if (sContNo == "") {	/* コンテナNo.未入力チェック */
			window.alert("コンテナNo.が未入力です。");
			return false;
	}

	if (sSealNo == "" && sJyuryo == "" && sSoJyuryo == "") {	/* シールNo.・重量未入力チェック */
			window.alert("詳細情報を入力して下さい。");
			return false;
	}
	return true;
}

// 数値チェック
function checknum(etext)
{
	if (etext.value == "")
		return false;

	if (isNaN(etext.value)) {
		alert("数値を入力して下さい。");
		etext.focus();
		etext.select();
		return false;
	}

	fTemp=parseFloat(etext.value)
    if (fTemp>99.9) {
		alert("99.9Ton以下の数値を入力して下さい。");
		etext.focus();
		etext.select();
		return false;
	}

	return true;
}

<!--
function gotoURL(){
    var gotoUrl=document.con.select.options[document.con.select.selectedIndex].value
    document.location.href=gotoUrl 
}
//-->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
	<td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaika1-2t.gif" width="506" height="73"></td>
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
		  <FORM NAME="con" METHOD="post" action="nyuryoku-ki-syori.asp" onSubmit="return ClickSend()">
                <table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">
                      コンテナNo.</font></b></td>
                    <td> 
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="ContNo" size="20" maxlength="12">
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
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">シールNo.</font></b></td>
                    <td> 
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="SealNo" size="20" maxlength="15">
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ 半角英数 ]</font>
							</td>
						  </tr>
						</table>
                      
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">貨物重量</font></b></td>
                    <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="Jyuryo" size="6"  maxlength="4" onblur="checknum(document.con.Jyuryo)">（t）
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ 半角数値 ]</font>
							</td>
						  </tr>
						</table>
                      
						&nbsp;&nbsp;&nbsp;<font size="-1">小数点以下1桁まで有効&nbsp;&nbsp;（例）10.2</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">総重量</font></b></td>
                    <td>
						<table border=0 cellpadding=0 cellspacing=0>
						  <tr>
							<td width=170>
								<input type="text" name="SoJyuryo" size="6"  maxlength="4" onblur="checknum(document.con.SoJyuryo)">（t）
							</td>
							<td align=left valign=middle nowrap>
								<font size=1 color="#2288ff">[ 半角数値 ]</font>
							</td>
						  </tr>
						</table>
                      
						&nbsp;&nbsp;&nbsp;<font size="-1">小数点以下1桁まで有効&nbsp;&nbsp;（例）10.2</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">リーファー</font></b></td>
                    <td>
                      <input type=checkbox name="rf"><font size=-1>リーファーの場合はチェックして下さい。</font>
                    </td>
                  </tr>
                  <tr> 
                    <td background="gif/tableback.gif" bgcolor="#000099" nowrap align="center"><b><font color="#FFFFFF">危険物</font></b></td>
                    <td>
                      <input type=checkbox name="dg"><font size=-1>危険物の場合はチェックして下さい。<sup>（※）</sup></font>
                    </td>
                  </tr>
                </table>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<font size=-1>（※） 消防法に関わる危険物の場合のみチェックして下さい。</font>
                <br><BR>
                <input type=submit value=" 送  信 " name="リセット">
        </form>
</center>

          <br>
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
              <td><a href="nyuryoku-kcsv.asp">CSVファイル転送</a></td>
            </tr>
            <tr> 
              <td>CSVファイル転送についての説明はここをクリック</td>
              <td>…</td>
              <td><a href="help08.asp">ヘルプ</a></td>
            </tr>
          </table>
</center>
          <br>
          　<br>
</td></tr></table>
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
    ' 海貨入力シールNo.、重量入力
    WriteLog fs, "4002","海貨入力シールNo.・重量入力", "00",","
%>
