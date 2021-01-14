<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin ""nyuryoku-in1.asp"
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
<!--
function gotoURL(){
    var gotoUrl=document.con.select.options[document.con.select.selectedIndex].value
    document.location.href=gotoUrl 
}
//-->
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから船社入力項目選択画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/nyuryoku-s.gif" width="506" height="73"></td>
          <td height="25" bgcolor="000099" align="right"><img src="gif/logo_hits_ver2.gif" width="300" height="25"></td>
        </tr>
        <tr>
          <td align="right" width="100%" height="48">
<%
' Added and Commented by seiko-denki 2003.07.18
	DisplayCodeListButton
'    DispMenu
' End of Addition by seiko-denki 2003.07.18
%>
          </td>
        </tr>
      </table>
      <br>　
      <br>　
      <br>　
      <br>　
      <center>
      <table>
        <tr>
          <td>本船動静入力かファイル転送かを選択の上、送信ボタンをクリックして下さい。</td>
        </tr>
      </table>
      <FORM NAME="con">
        <table border="1" cellspacing="2" cellpadding="3">
          <tr> 
            <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>入力項目選択</b></font></td>
            <td nowrap> 
              <p>
                <select name="select" size="2">
                  <option value="nyuryoku-in1.asp">本船動静入力</option>
                  <option value="nyuryoku-csv.asp">ファイル転送</option>
                </select>
            　</p>
            </td>
          </tr>
        </table>
        <br>　<br>
        <INPUT TYPE=BUTTON VALUE=" 送  信 " onClick="gotoURL()">
      </form>
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

<!-------------船社入力項目選択画面終わり--------------------------->

</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 船社入力項目選択
    WriteLog fs, "船社入力項目選択", ""
%>
