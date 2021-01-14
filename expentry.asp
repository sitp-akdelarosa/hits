<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 輸出コンテナ照会
    WriteLog fs, "1001","輸出コンテナ照会","00", ","
%>

<html>
<head>
<title></title>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから照会画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/expentryt.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48"> 
<%
' Added and Commented by seiko-denki 2003.07.07
	DisplayCodeListButton
'    DispMenu
'	Dim strScriptName,strRoute
'	strScriptName = Request.ServerVariables("SCRIPT_NAME")
'	strRoute = SetRoute(strScriptName)
'	Session.Contents("route") = strRoute
' End of Addition by seiko-denki 2003.07.07
%>
          </td>
        </tr>
      </table>
      <center>
<!-- commented by seiko-denki 2003.07.17
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
end of comment by seiko-denki 2003.07.17 -->
		<BR>
		<BR>
		<BR>
<table border=0><tr><td align=left>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>キー入力の場合</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<center>
      <table width="480">
        <tr>
          <td colspan="4">参照したいコンテナNo.または、Booking No.を半角で入力し、『輸出照会』ボタンをクリックして下さい。
              複数入力する場合には","で区切って入力して下さい。<br>
          </td>
        </tr>
        <tr>
          <td width="20">&nbsp;</td>
          <td> 
           複数コンテナNo.入力 例）FYTU2334999,HYKU9882272,DYTU3998821</td>
        </tr>
      </table>
      <form action="expcntnr.asp">
              <table border="1" cellspacing="1" cellpadding="3" bgcolor="#ffffff">
                <tr>
                  <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>コンテナNo.</b></font></td>
                  <td nowrap>
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td>
							<input type=text name=cntnrno size=20 maxlength="100">
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
                  <td align="center" colspan="2">または
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="#000099"><font color="#FFFFFF"><b>Booking No.</b></font></td>
                  <td nowrap> 
					<table border=0 cellpadding=0 cellspacing=0>
					  <tr>
						<td>
							<input type="text" name=booking size=20 maxlength="100">
						</td>
						<td align=left valign=middle nowrap>
							<font size=1 color="#ee2200">[ 必須入力 ]</font><BR>
							<font size=1 color="#2288ff">[ 半角英数 ]</font>
						</td>
					  </tr>
					</table>
                  </td>
                </tr>
              </table>
              <br>
              <input type=submit value="   輸出照会   "><br><br>
				何も入力しないで『輸出照会』ボタンを押すと照会結果の
				サンプルを表示します。
      </form>
<!			ブッキング情報照会はここをクリック ･･･ <a href="bookentry.asp"><!ブッキング情報照会</a>
			<BR><BR>

</center>
      <table>
        <tr> 
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSVファイル入力の場合</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
<center>
      <table border="0" cellspacing="1" cellpadding="2">
        <tr> 
          <td> 
            <p>検索条件をファイル転送する場合はここをクリック</p>
          </td>
          <td>…</td>
          <td><a href="expcsv.asp">CSVファイル転送</a></td>
        </tr>
        <tr> 
          <td>CSVファイル転送についての説明はここをクリック</td>
          <td>…</td>
          <td><a href="help01.asp">ヘルプ</a></td>
        </tr>
      </table>
        <br>&nbsp;<br>
</center>
            <table>
              <tr> 
                <td>　</td>
              </tr>
            </table>		  
<center>
              <BR>
              <BR>
            </center>
      </center>

</td></tr></table>

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
<!-------------照会画面終わり--------------------------->
<%
    DispMenuBarBack "../index.asp" 'http://www.hits-h.com/index.asp
%>
</body>
</html>
