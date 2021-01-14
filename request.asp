<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
    ' File System Object の生成
    Set fs=Server.CreateObject("Scripting.FileSystemobject")
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<META name="GENERATOR" content="IBM HomePage Builder 2001 V5.0.0 for Windows">
<SCRIPT Language="JavaScript">
<%
    DispMenuJava
%>
function ClickSend() {

	if (document.con.name1.value == "") {	
		window.alert("会社名が未入力です。");
		return false;
	}

	if (document.con.address1.value == "") {	
		window.alert("住所が未入力です。");
		return false;
	}

}

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録画面--------------------------->

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
   					<form NAME="con" action="request-syori.asp" method=post onSubmit="return ClickSend()">
      <table width="500" cellpadding="0">
					<tr>
						<td bordercolor="#FFFFFF">HiTS V3をご利用頂きありがとうございます。
            <BR>
<BR>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;■ 質問等がございましたら<a href="mailto:mrhits@hits-h.com">E-mail</a>でお問合せ下さい。 <br>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;■ Ｑ＆Ａ集のページが<a href="qa/index.html">こちら</a>からご参照になれます。<BR>
<BR>
							なお、お急ぎの場合は下記にお問い合わせ下さい。<BR><BR>

							<table border=0 cellpadding=1 cellspacing=1>
							  <tr>
								<td colspan=5 align=left>
								&nbsp;【 電話でのお問い合わせ先 】
								</td>
							  </tr>

							  <tr><td colspan=5 height=2></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4>・博多港物流ITシステムの運用に関する事<BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2>博多港ふ頭株式会社</td>
								<td align=left nowrap>担当：木本</td>
							  </tr>
							  <tr>
								<td width=15><BR></td>
								<td nowrap align=left valign=top>TEL 092-663-3021<BR>
								</td>
								<td><BR></td>
							  </tr>

							  <tr><td colspan=5 height=5></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4>・博多港物流ITシステムの開発に関する事<BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2>福岡市港湾局　港湾振興部　物流企画課</td>
								<!-- <td align=left nowrap>担当：谷口</td> -->
							  </tr>
							  <tr>
								<td width=15><BR></td>
								<td nowrap align=left valign=top>TEL 092-282-7108<BR>
                  </td>
								<td><BR></td>
							  </tr>

							  <tr><td colspan=5 height=5></td></tr>

							  <tr>
								<td width=20 rowspan=3><BR></td>
								<td nowrap align=left valign=top colspan=4><BR>
								</td>
							  </tr>
							  <tr>
								<td width=15 rowspan=2><BR></td>
								<td nowrap align=left valign=top colspan=2></td>
								<td align=left nowrap></td>
							  </tr>
							  <tr>
								<td width=15></td>
								<td nowrap align=left valign=top></td>
								<td></td>
							  </tr>

							</table>
						</td>
					</tr>
				</table> 
<BR>
					
<%
    DispMenuBar
%>
		</FORM></CENTER></td>
	</tr>
</table>
<!-------------登録画面終わり--------------------------->
<%
    DispMenuBarBack "http://www.hits-h.com/index.asp"
%>
</body>
</html>
<%

    WriteLog fs, "9001", "利用者アンケート・Q&A", "00", ","
%>
