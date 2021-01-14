<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' セッションのチェック
    CheckLogin "nyuryoku-kaika.asp"
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
	  
<table>
 <tr>
            <td> 
              <center>
                <BR>
                <BR>
              </center>
			<br>　<br>

          <table>
            <tr> 
              <td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
              
                  <td nowrap><b>荷主管理番号で情報を照会</b></td>
              <td><img src="gif/hr.gif"></td>
            </tr>
          </table>
		  <center>		  
 荷主管理番号で情報を照会する場合は、下記の目的の作業を選択して下さい。<BR><BR>
                <table border="0" cellspacing="1" cellpadding="2">
                  <tr> 
                    
                    <td> 
                      <table border=0 cellspacing=1 cellpadding=4>
						<tr><td>
                        ・<a href="ms-kaika-expinfo.asp">（輸出）貨物情報入力 </a> 
						</td></tr>
						<tr><td>
                        ・<a href="ms-kaika-expcontinfo.asp">（輸出）コンテナ情報入力</a> 
						</td></tr>
						<tr><td>
                        ・<a href="ms-kaika-impcontinfo.asp">（輸入）コンテナ情報入力</a> 
                      	</td></tr>
					  </table>
                    </td>
                  </tr>
                </table></center>
  </td>
  </tr>
 </table>  
 </center>
</td>
 </tr>
 <tr>
    <td valign="bottom"> 
<!-------------処理選択画面終わり--------------------------->
<%
    DispMenuBar
%>
</td>
  </tr>
</table>
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>

</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")

    ' 海貨入力項目選択
    WriteLog fs, "4001","海貨入力", "00", ","
%>
