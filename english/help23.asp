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
// Added and Commented by seiko-denki 2003.07.18
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=330");
  CodeWin.focus();
}
// End of Addition by seiko-denki 2003.07.18
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
          <td rowspan=2><img src="gif/helpt2.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
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
				<%=strRoute%> &gt; ヘルプ
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.18 -->
		<BR>
		<BR>
		<BR>
        <table>
          <tr>
            <td> 
              <table>
                <tr> 
                  <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap><b><font color="#000000">ブッキング情報照会結果出力</font></b></td>
                  <td><img src="gif/hr.gif" width="400" height="3"></td>
                </tr>
              </table>

              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ａ．CSVファイル出力とは？</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">画面に表示されているコンテナのすべての情報をCSVファイルとしてお手持ちのパソコンに保存することができます。<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>ｂ．CSVファイルとは？</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>情報がカンマ「,」区切りで羅列されたテキストファイルです。<br>
                      <dd> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5">
                                <tr> 
                                  <td bgcolor="#FFFFFF" nowrap>
									Booking No.,船社,船名,Voyage No.,仕向港,空コン搬出場所,サイズ,タイプ<br>
                                    ABS544423, ABCDE LINE, WAN CHAN 211, 12345, HONGKONG, 香椎VP, 20, RF<br>
                                    ABS544423, ABCDE LINE, WAN CHAN 211, 12345, PARIS, 香椎VP, 40, DG<br>
                                    <br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table><br>
                      <dt>このファイルをWindows付属のメモ帳で開くと上の例のようにわかりにくいままですが、たとえばEXCELのような表計算ソフトで開くと下のようにわかりやすい表示となります。 
                      <dd>
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                          <td> 
                                  
                              <table border="1" bgcolor="#FFFFFF" >
                                <tr valign="top"> 
                                  <td nowrap>Booking No.</td>
                                  <td nowrap>船社 </td>
                                  <td nowrap>船名</td>
                                  <td nowrap>Voyage No.</td>
                                  <td nowrap>仕向港</td>
                                  <td nowrap>空コン<BR>搬出場所</td>
                                  <td nowrap>サイズ</td>
                                  <td nowrap>タイプ</td>
                                </tr>
                                <tr valign="top"> 
                                  <td nowrap>ABS544423</td>
                                  <td nowrap>ABCDE LINE</td>
                                  <td nowrap>WAN CHAN 211</td>
                                  <td nowrap>12345</td>
                                  <td nowrap>HONGKONG</td>
                                  <td>香椎VP </td>
                                  <td>20</td>
                                  <td>RF</td>
                                </tr>
                                <tr valign="top"> 
                                  <td>ABS544423</td>
                                  <td>ABCDE LINE</td>
                                  <td>WAN CHAN 211</td>
                                  <td>12345</td>
                                  <td>PARIS</td>
                                  <td>香椎VP</td>
                                  <td>40</td>
                                  <td>DG</td>
                                </tr>
                              </table>
                          </td>
                        </tr>
                      </table><br>
                      <dt>CSVファイルは表計算ソフトに限らず、さまざまなデータベースソフトでも読み込むことが可能です。 <br><br>
                    </dl>
				   
				   </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ｃ．本画面で出力されるCSVファイルの内容
                    </font><font color="#666666"> </font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>画面に表示されている１つのコンテナについて次の項目を出力します。<br>
                      <dd> 
                        <table border="1" cellspacing="1" cellpadding="5" width=500>
                          <tr> 
                            <td bgcolor="#FFFFFF">
								Booking No., 船社, 船名, VoyageNo., 仕向港, 空コン搬出場所,
								サイズ, タイプ, 高さ, 予約本数, 搬出済本数, コンテナNo.
							</td>
                          </tr>
                        </table><br>
                      <dt>上のCSVファイルの例のように１行目が項目名で２行目が値となります。<BR>
                      
                    </dl>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>ｄ．CSVファイル出力の方法</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dt> 画面上の『CSVファイル出力』ボタンを押すことで保存先と保存ファイル名を指定する画面が表示されます。<br>
                    
                    <dd> 
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                          <td> 
                                  <form>
                                    <input type=button value=" CSVファイル出力" name="ボタン">
                                  </form>
                            
                          </td>
                        </tr>
                      </table><br>
                    <dt>保存先と保存ファイル名はともに自由ですが、ファイル名の拡張子は通常、「.csv」とします。
                    <dd><font color="#FF0033">【例】</font>C:\MyDocument内に abcdef.csv  というファイル名で保存します。<br><br>
                  </td>
                </tr>
              </table>



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