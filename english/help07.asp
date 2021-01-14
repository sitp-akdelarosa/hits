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
          <td rowspan=2><img src="gif/helpt.gif" width="506" height="73"></td>
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
                  <td nowrap><b>本船動静入力</b></td>
                  <td><img src="gif/hr.gif" width="540"></td>
                </tr>
              </table>
<center>
              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>ａ．CSVファイル転送とは？</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">入力したい情報が多い場合、何度も入力するのは面倒です。<br>
                    そこで、本システムでは情報を羅列したファイルを作り、そのファイルを転送することでまとめて入力する機能を用意しています。<br>
                    本システムに転送できるファイルの形式は「CSVファイル」といわれる一般的なものです。<br>
                    この「CSVファイル」を作成し転送を行う手順を以下に説明します。<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b>ｂ．必要なアプリケーション</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt>CSVファイルの作成はWindows付属のメモ帳で可能です。あるいは、EXCELで作成してCSVファイル形式で保存することも可能です。<br>
                    </dl>
				   
				   </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ｃ．CSVファイルの作成</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <dl> 
                        <dt>前述のアプリケーションを使って、コールサイン、Voyage No．、港名・・・の順にひとつひとつの値をカンマ「,」で区切りながら１行に１セットの情報を記述します。<br>

 <dt>&nbsp;&nbsp;<font color=#ff0000>【注意】</font>
 <dd>同じ本船（同一のコールサインとVoyage No.）の行は続けて記述してください。
<dd>同じ本船の同一港に対するデータは１行で記述してください。（例えば、着岸Long Scheduleと着岸予定時刻を入力する場合は２行
に分けず１行で記述してください）<p>
                          <table border="1" cellspacing="1" cellpadding="5" width=500>
                            <tr> 
                              <td bgcolor="#FFFFFF" nowrap><font size="1">
										A1284, B3567, JPTYO, 2002/3/12/14/50, 2002/3/12/15/00, 2002/3/12/16/05, 2002/3/12, 2002/3/12
									<br>
										F8976, D7909, JPTYO, 2002/3/18/03/00, 2002/3/18/03/08, 2002/3/18/4/30, 2002/3/18, 2002/3/18
									</font>
								</td>
                            </tr>
                          </table>
                          <br>
                        <dt>1行分の項目の詳細仕様<BR>
                        <dd>
                          <table width="100" border="1" cellspacing="0" cellpadding="2" bgcolor="#FFFFFF">
                            <tr bgcolor="#99aaFF" align="center"> 
                              <td nowrap><b><font color="#333333">項目</font></b></td>
                              <td nowrap><b><font color="#333333">例</font></b></td>
                              <td nowrap><b><font color="#333333">入力仕様</font></b></td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>コールサイン</td>
                              <td nowrap>A1284</td>
                              <td nowrap>半角大文字英数字7桁以内</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>Voyage No.</td>
                              <td nowrap>B3567</td>
                              <td nowrap>半角大文字英数字12桁以内<br>
                                特殊記号含む（'-'、'/'など）</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>港名 </td>
                              <td nowrap>JPTYO</td>
                              <td nowrap>UNLOコード（半角大文字英字数字５桁）</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>着岸予定時刻<br>
                                （年月日時分）</td>
                              <td nowrap>2002/3/12/14/5 </td>
                              <td nowrap>・年：数字4桁<br>
                                ・その他：数字2桁('01'と'1'の両方の表現に対応)<br>
                                ・以上を半角スラッシュ「/」で区切る。<br>
                                ・値が無い場合はスラッシュだけを残す(「//」)</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>着岸完了時刻</td>
                              <td nowrap>（同様の形式）</td>
                              <td nowrap>同上</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>離岸完了時刻 </td>
                              <td nowrap>（同様の形式）</td>
                              <td nowrap>同上</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>着岸Long Schedule<BR>（年月日）</td>
                              <td nowrap>2002/3/12</td>
                              <td nowrap>同上</td>
                            </tr>
                            <tr valign="top"> 
                              <td nowrap>離岸Long Schedule<BR>（年月日）</td>
                              <td nowrap>2002/3/12</td>
                              <td nowrap>同上</td>
                            </tr>
                          </table>
                          <br>
                        <dt>ファイル名は何でもかまいませんが、拡張子は通常「.csv」とします。保存先も自由です 
                        <dd><font color="#FF0033">【例】</font>C:\MyDocument内に abcdef.csv  というファイル名で保存します。 
                      </dl>
                  </td>
                </tr>
                <tr> 
                    <td colspan="2" bgcolor="#99ccFF"><b>ｄ．CSVファイルの転送</b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                    <td width="575"> 
                      <ul>
                        <li>画面上のCSVファイル転送をクリックすると次のようなCSVファイルを指定する画面が表示されます。<br>
                          <table border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                              <td> 
                                <table border="1" cellspacing="1" cellpadding="5">
                                  <tr> 
                                    <td bgcolor="#FFFFFF" align="center"> 
                                      <form>
                                        <table border="1" cellspacing="0" cellpadding="2">
                                          <tr> 
                                            <td bgcolor="#000099" nowrap> <font color="#FFFFFF"><b>CSVファイル名</b></font> 
                                            </td>
                                            <td nowrap> 
                                              <input name=csvfile size=30 accept="text/css">
                                            </td>
                                            <td nowrap> 
                                              <input type=button value="参照..." name="ボタン">
                                            </td>
                                          </tr>
                                        </table>
                                        <input type=button value=" 送  信 " name="ボタン">
                                      </form>
                                    </td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table><br>
                        <li>空欄に作成したCSVファイルのフルパスを記述します。 <br>
                          <font color="#FF0033">【例】</font>作成例の場合は「C:\MyDocument\abcdef.csv」と記述します。<br>
                        <li>手入力するのが面倒な場合は、［参照...］ボタンを押すとファイルを選択する画面が出ますので、保存先のフォルダとファイルを順に選択していくことでファイル名が自動的に入力されます。<br>
                        <li>最後に［送信］ボタンを押します。<br>
                        <li>検索結果は通常の画面で表示されます。
                          <p> 
                          <table border="1" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#FF9933" nowrap valign="top">注意</td>
                              <td bgcolor="#FFFFFF">ファイルの作成が規則どおりできていないとシステムは内容を読み出すことができずエラーを表示します。その場合は、ファイルの内容を見直し、修正した後再度送信を行ってください。</td>
                            </tr>
                          </table>
                      </ul>
                    </td>
                </tr>
              </table>

</center>

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
          <td bgcolor="000099" height="10"><img src="gif/1.gif"></td>
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