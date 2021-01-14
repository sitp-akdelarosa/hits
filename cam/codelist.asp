<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
Dim strUser
strUser = Request.QueryString("user")
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="../gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録コード一覧画面--------------------------->

<center>

<BR>

<font size=4><b>登録コード一覧</b></font>

<BR><BR>

<table border=0>
  <tr>
	<td align=left valign=middle nowrap>
		本システムに登録されているコードと<BR>名称の一覧を表示します。
	</td>
  </tr>
</table>

<BR>

<% If strUser="" Then %>

	<table border=1 cellpadding=3 cellspacing=1 bgcolor="#ffffff">
		<tr>
			<td align=center nowrap>
				<font color="#ff3300"><b>ログインしていない時は表示できません。</b></font>
			</td>
		</tr>
	</table>
	<BR>
<% Else %>

目的のコードを選択して下さい。

<BR><BR>

<table border=0>
  <tr>
	<td align=left valign=middle nowrap>
		１． <a href="codelist-detail.asp?kind=1">荷主コード</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		２． <a href="codelist-detail.asp?kind=2">海貨コード</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		３． <a href="codelist-detail.asp?kind=3">陸運業者コード</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		４． <a href="codelist-detail.asp?kind=4">船社コード</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		５． <a href="codelist-detail.asp?kind=5">港運コード</a>
	</td>
  </tr>
  <tr>
	<td align=left valign=middle nowrap>
		６． <a href="codelist-detail.asp?kind=6">船社からの船名照会</a>
	</td>
  </tr>
</table>

<BR>

<table border=0 width=85%>
  <tr>
	<td align=left valign=middle>
		画面に表示されるコードをマウスで選択してコピー（Ctrl + C）することで、キー入力枠のところに張り付け（Ctrl + V）できます。
	</td>
  </tr>
</table>

<% End If %>

<form>
	<input type=button value=" close " onClick="JavaScript:window.close()">
</form>

</center>
</body>
</html>

<%
%>
