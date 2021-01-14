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
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td rowspan=2><img src="../gif/helpt.gif" width="506" height="73"></td>
	      <td height="25" bgcolor="000099" align="right"><img src="../gif/logo_hits_ver2.gif" width="300" height="25"></td>
  </tr>
  <tr>
	<td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strRoute
'	strRoute = Session.Contents("route")
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.07
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%> &gt; ヘルプ
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.07 -->
		<BR>
		<BR>
		<BR>
        <table>
          <tr>
            <td align="center"> 
              <table>
                <tr> 
                  <td><img src="../gif/botan.gif" width="17" height="17" vspace="4"></td>
                  <td nowrap> <b><font color="#000000">輸入コンテナ情報照会キー入力</font></b>&nbsp;&nbsp;</td>
                  <td><img src="../gif/hr.gif"></td>
                </tr>
              </table>

              <table border="0" cellspacing="2" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ａ．CSVファイル転送とは？</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">参照したいコンテナNo.やBL No．が多い場合、何度も入力して検索を実行するのは面倒です。<br>
                    そこで、本システムでは参照したいコンテナNo.、または、BL No．を羅列したファイルを作り、そのファイルを転送してまとめて検索を行う機能を用意しています。<br>
                    本システムに転送できるファイルの形式は「CSVファイル」といわれる一般的なものです。<br>
                    この「CSVファイル」を作成し転送を行う手順を以下に説明します。<br>
                    &nbsp; </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ｂ．必要なアプリケーション</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575">CSVファイルの作成はWindows付属のメモ帳で可能です。あるいは、EXCELで作成してCSVファイル形式で保存することも可能です。<br>
                    &nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ｃ．CSVファイルの作成 
                    </font><font color="#666666"> </font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dl> 
                      <dt><b>（１）複数のコンテナNo.で参照したい場合</b> 
                      <dd>前述のアプリケーションを使って１行に１個のコンテナNo.を記述し、目的のコンテナNo.の数だけ行を作ります。<br>
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5" width=300>
                                <tr> 
                                  <td bgcolor="#FFFFFF">KYGU2234455<BR>
                                    GFDU2556379<BR>
                                    FGYU9882567<br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      <dd>ファイル名は何でもかまいませんが、拡張子は通常「.csv」とします。保存先も自由です。<br>
                        <font color="#FF0033">【例】</font>C:\MyDocument内に abcdef.csv  というファイル名で保存します。<br>
                        <br>
                      <dt><b>（２）複数のBL No．で参照したい場合</b><br>
                      <dd>コンテナNo.の場合と同様に１行に１個のBL No．を記述し、目的のBL No．の数だけ行を作ります。 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td valign="top" nowrap><font color="#FF0033">【例】</font></td>
                            <td> 
                              <table border="1" cellspacing="1" cellpadding="5" width=300>
                                <tr> 
                                  <td bgcolor="#FFFFFF">BL12546<BR>
                                    BL88976<br>
                                    <br>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                        <BR>
                        ファイル名の規則も同様です。<BR>
                        <table border="1" cellspacing="0" cellpadding="3">
                          <tr> 
                            <td bgcolor="#FF9933">注意</td>
                            <td bgcolor="#FFFFFF">１つのCSVファイルの中にコンテナNo.とBL No．を混在させることはできません。 
                            </td>
                          </tr>
                        </table>
                        <br>
                    </dl>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" bgcolor="#99ccFF"><b><font color="#000000">ｄ．CSVファイルの転送</font></b></td>
                </tr>
                <tr> 
                  <td width="15"> </td>
                  <td width="575"> 
                    <dt> 画面上のCSVファイル転送をクリックすると次のようなCSVファイルを指定する画面が表示されます。<br>
                    <dd> 
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
                      </table>
                    <dt> 
                      <ul>
                        <li>空欄に作成したCSVファイルのフルパスを記述します。<br>
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
                      <br>
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
	      <td valign="bottom" align="right"><a href="index.html"><img src="../gif/b-home.gif" border="0" width="270" height="23" usemap="#map"></a></td>
        </tr>
        <tr>
          <td bgcolor="000099" height="10"><img src="../gif/1.gif"></td>
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