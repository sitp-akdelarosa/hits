<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%

    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemObject")

    ' BASP21 コンポーネントの作成
    Set bsp=Server.CreateObject("basp21")

'    パラメータ
'      sSvrName		: SMTPサーバー名
'      sMailto		: 宛先（宛先名<宛先メールアドレス>）
'      sMailfrom	: 送信者（送信者名<送信者メールアドレス>）
'      sSubj		: 件名
'      sBody		: 本文

	sSvrName = "mail.cont-info.com"
	sMailto = "request-syori.asp<cont-info@cont-info.com>"
'	if Request.form("mail") ="" then
		sMailfrom = "アンケート<NoName@cont-info.com>"
'	else
'		sMailfrom = Request.form("mail")
'	end if
	sSubj = "【利用者アンケート】"
	sBody = "【利用者アンケート】" & vbcrlf & vbcrlf & _ 
	        "会社名," & Request.form("name1") & vbcrlf & _ 
			"ご住所,〒" & Request.form("posta1") & "-" & Request.form("posta1") & vbtab & _
			              Request.form("ken") & Request.form("address1") & vbcrlf & _  
			"ご担当者名," & Request.form("tantouname") & vbcrlf & _
			"E-mail," & Request.form("mail") & vbcrlf & _
			"システムの情報について１," & Request.form("radiobutton") & vbcrlf & _
			"２," & Request.form("add") & vbcrlf & _
			"３," & Request.form("change")
	rc = bsp.SendMail(sSvrName,sMailto,sMailfrom, sSubj,sBody,"")

	if trim(rc) = "" then
		strError = "正常に送信されました。"
	    WriteLog fs, "9001", "利用者アンケート・Q&A", "10", "," & "送信の結果:0(成功)"
	else	
		strError = "正常に送信できませんでした。" 
	    WriteLog fs, "9001", "利用者アンケート・Q&A", "10", "," & "送信の結果:1(失敗)" & rc
	end if	

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
<!-------------ここからエラー画面--------------------------->
<table width="100%" border="0" cellspacing="0" cellpadding="0" height=100%>
	<tr>
		<td valign=top>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td rowspan=2><img src="gif/requestt.gif" width="506" height="73"></td>
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
						<td nowrap>
							<dl> 
								<dt><font color="#000066" size="+1">【利用者アンケート送信画面】</font><br>
								<dd>
<%
    ' エラーメッセージの表示
    DispErrorMessage strError 
%>
							</dl>
						</td>
					</tr>
				</table>
				<form>
					<br><br>
					<input type="button" value=" 戻  る " onclick="history.back()">
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
<!-------------エラー画面終わり--------------------------->
<%
    DispMenuBarBack "request.asp"
%>
</body>
</html>

