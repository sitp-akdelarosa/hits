<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
''    ' セッションのチェック
    CheckLogin "nyuryoku-im.asp"
%>
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
					<td rowspan=2><img src="gif/csvt.gif" width="506" height="73"></td>
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
						<td><img src="gif/botan.gif" width="17" height="17" vspace="4"></td>
						<td nowrap><b>実入り倉庫到着時刻（指示済）</b></td>
						<td><img src="gif/hr.gif"></td>
					</tr>
				</table>
				<table>
					<tr>
						<td nowrap>
							<center><font color="#000066" size="+1">【実入り倉庫到着時刻用ファイル転送画面】</font></center><br>
							入力したい実入り倉庫届け時刻情報を含むCSVファイルを選択し、送信ボタンを
								クリック<br>して下さい。</td></tr>
				</table>
				<form action="nyuryoku-im-csvin.asp?kind=cntnr" enctype="multipart/form-data" method="post"> 
					<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
						<tr> 
							<td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>コンテナNo.を羅列したCSVファイルの転送の場合</b></font>
							</td>
						</tr>
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
					<br>
				</form>
				<br><br>
				<form action="nyuryoku-im-csvin.asp?kind=bl" enctype="multipart/form-data" method="post" id=form1 name=form1> 
					<table border="1" cellspacing="2" cellpadding="3" bgcolor="#ffffff">
						<tr> 
							<td bgcolor="#000099" nowrap align="center" background="gif/tableback.gif" colspan="2"><font color="#FFFFFF"><b>BL No.を羅列したCSVファイルの転送の場合</b></font>
							</td>
						</tr>
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
	  				<input type=submit value=" 送  信 " id=submit1 name=submit1>
					<br>
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
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "nyuryoku-im.asp"
%>
</body>
</html>

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' 実入り倉庫届け時刻指示用ファイル転送画面
    WriteLog fs, "4007","海貨入力実入り倉庫到着時刻-CSVファイル転送","00",","
%>
