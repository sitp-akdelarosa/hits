<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript"><!--
function FancBack()
{
        window.history.back();
}

function LinkSelect(form, sel)
{
        adrs = sel.options[sel.selectedIndex].value;
        if (adrs != "-" ) parent.location.href = adrs;
}
// -->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="gif/infot.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<!-- commented by seiko-denki 2003.07.18

			<FORM action=''>

				<SELECT NAME='link' onchange='LinkSelect(this.form, this)'>
					<OPTION VALUE='#'>Contents
					<option value='../index.asp'>TOP</option>
					<option value='#'>コンテナ情報照会 </option>
					<option value='../userchk.asp?link=expentry.asp'>├ 輸出コンテナ情報照会 </option>
					<option value='../userchk.asp?link=impentry.asp'>└ 輸入コンテナ情報照会 </option>
					<option value='#'>各社入力画面</option>
					<option value='../userchk.asp?link=nyuryoku-in1.asp'>├ 船社/ターミナル入力 </option>
					<option value='../userchk.asp?link=nyuryoku-kaika.asp'>├ 海貨入力 </option>
					<option value='../userchk.asp?link=nyuryoku-te.asp'>├ ターミナル入力 </option>
					<option value='../userchk.asp?link=rikuun1.asp'>└ 陸運入力</option>
					<option value='../userchk.asp?link=sokuji.asp'> 即時搬出システム </option>
					<option value='../userchk.asp?link=hits.asp'>ストックヤード利用システム</option>
					<option value='../userchk.asp?link=terminal.asp'>ゲート前映像・混雑状況照会 </option>
					<option value='../userchk.asp?link=request.asp'>利用者アンケート・Ｑ＆Ａ</option>
				</SELECT>
			</FORM>
End of comment by seiko-denki 2003.07.18 -->

          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.18
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right"> <font color="#333333" size="-1">
              Top &gt; 利用上のお願い</font> </td>
		  </tr>
		</table>
End of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
        <table width=550>
          <tr>
            <td>
              <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap><b><font color="#000000">利用に際してのお願い</font></b></td>
                  <td><img src="gif/hr.gif" width="320" height="3"></td>
                </tr>
              </table>
			  <ul>
			  <li>HiTS V2の利用に際しては、利用方法を理解のうえ利用される方自身の責任のもとにご活用下さい。<br>
		
			  <li>安全のため自動車運転中に携帯電話での利用はしないで下さい。<p><br>
			  </ul>
			  <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                    <td nowrap><b><font color="#000000">免責事項</font></b>&nbsp;&nbsp;</td>
                  <td><img src="gif/hr.gif" width="400" height="3"></td>
                </tr>
              </table>
			  <ul>
			  <li>利用者が当システムを利用すること、または、利用できなかったことに関連して生ずる一切の損害、トラブルに関していかなる責任も負いかねますのでご承知下さい。
			  </ul>
              
            </td>
   </tr>
  </table>
 <!---------->
  </center>
    </td>
 </tr>
 <tr>
    <td valign="bottom"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
	      <td valign="bottom" align="right"><a href="index.html"><img src="gif/b-home.gif" border="0" width="270" height="23" usemap="#map"></a></td>
        </tr>
        <tr>
          <td bgcolor="000099" height="10"><img src="gif/1.gif" ></td>
  </tr>
</table>
 </td>
 </tr>
 </table>
<!-------------画面終わり--------------------------->
<map name="map"> 
  <area shape="poly" coords="20,0,152,0,134,22,0,22" href="JavaScript:FancBack()">
  <area shape="poly" coords="154,0,136,22,284,22,284,0" href="http://www.hits-h.com/index.asp">
</map>
</body>
</html>