<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
  <tr>
  <td valign=top>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td rowspan=2><img src="gif/bookingt.gif" width="506" height="73"></td>
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
		<table width=95% cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td align="right">
			  <font color="#333333" size="-1">
				<%=strRoute%>
			  </font>
			</td>
		  </tr>
		</table>
		<BR>
		<BR>
		<BR>
      <table>
        <tr>
          <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
          <td nowrap><b>CSVファイル転送</b></td>
          <td><img src="gif/hr.gif"></td>
        </tr>
      </table>
    <table>
      <tr>
        <td nowrap>
            検索条件を入力したCSVファイルを選択し、『送信』ボタンを押して下さい。
		</td>
	  </tr>
    </table>
    <form action="bookcsvin.asp" method="POST" enctype="multipart/form-data">
      <table border="1" cellspacing="2" cellpadding="3" bgcolor="#FFFFFF">
        <tr>
          <td bgcolor="#000099" nowrap>
            <font color="#FFFFFF"><b>CSVファイル名</b></font>
          </td>
          <td nowrap> 
            <input type=file name=csvfile size=50 accept="text/css">
          </td>
        </tr>
      </table>
		<BR>
      <input type=submit value=" 送  信 ">
    </form>
    <br>
    <br>
    <table border="0" cellspacing="1" cellpadding="2">
      <tr> 
        <td>CSVファイル転送についての説明はここをクリック</td>
        <td>…</td>
        <td><a href="help22.asp">ヘルプ</a></td>
      </tr>
    </table>
    </center>
  </td></tr> 
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
    DispMenuBarBack "bookentry.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    WriteLog fs, "1012","ブッキング情報照会-CSVファイル転送","00", ","
%>
