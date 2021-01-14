<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	'
	'	【海貨入力】	作業選択画面
	'
%>


<%
    ' セッションのチェック
    CheckLogin "ms-kaika.asp"
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
<!-------------ここから処理選択画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
    <td valign=top>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td rowspan=2><img src="gif/kaikat.gif" width="506" height="73"></td>
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
      <center>
      <table>
        <tr> 
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>海貨入力</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
      <br>
      <form name=con>
        <table border="1" cellspacing="2" cellpadding="3">
          <tr> 
            <td bgcolor="#000099" nowrap><font color="#FFFFFF"><b>入力選択</b></font></td>
            <td nowrap>
              <SELECT NAME="select" size="3">
                <option value="ms-kaika-expinfo.asp">（輸出）貨物情報入力</option>
                <option value="ms-kaika-expcontinfo.asp">（輸出）コンテナ情報入力</option>
                <option value="ms-kaika-impcontinfo.asp">（輸入）コンテナ情報入力</option>
              </select>
            </td>
          </tr>
        </table>
        <br>　<br>
        <INPUT TYPE=BUTTON VALUE=" 送  信 " onClick="gotoURL()">
      </form>
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
<!-------------処理選択画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 海貨入力項目選択
    WriteLog fs, "海貨入力項目選択", ""
%>
