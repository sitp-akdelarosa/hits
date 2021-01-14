<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	Dim iKind,sql,sTblName,sCodeName,sKindName,sCode,sFull,sAbrev
	ConnectSvr conn, rsd

	iKind = Request.QueryString("kind")

	Select Case iKind
		Case 1
			sTblName  = "mShipper"
			sCodeName = "Shipper"
			sKindName = "荷主"
		Case 2
			sTblName  = "mForwarder"
			sCodeName = "Forwarder"
			sKindName = "海貨"
		Case 3
			sTblName  = "mTrucker"
			sCodeName = "Trucked"
			sKindName = "陸運業者"
		Case 4
			sTblName  = "mShipLine"
			sCodeName = "ShipLine"
			sKindName = "船社"
		Case 5
			sTblName  = "mOperator"
			sCodeName = "OpeCode"
			sKindName = "港運"
		Case 6
			sTblName  = "mShipLine"
			sCodeName = "ShipLine"
			sKindName = "船社"
	End Select

	sql = "SELECT FullName,NameAbrev," & sCodeName & " FROM " & sTblName
	rsd.Open sql, conn, 0, 1, 1
%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<SCRIPT Language="JavaScript">

</SCRIPT>
</head>
<body bgcolor="DEE1FF" text="#000000" link="#3300FF" background="gif/back.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-------------ここから登録コード一覧画面--------------------------->

<center>

<BR>

<font size=4><b><%=sKindName%>コード一覧</b></font>

<BR><BR>

<table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
  <tr align="center" bgcolor="#FFCC33"> 
	<td align=center valign=middle height=15 nowrap>名称</td>
	<td align=center valign=middle height=15 nowrap>コード</td>
	<td align=center valign=middle height=15 nowrap>略称</td>
  </tr>

<%
	    Do While Not rsd.EOF

			sCode  = Trim(rsd(sCodeName))
			sFull  = Trim(rsd("FullName"))
			sAbrev = Trim(rsd("NameAbrev"))
%>

  <tr>
	<td align=left valign=middle nowrap>
<% If iKind=6 Then %>
		<a href="codelist-vessel.asp?vsl=<%=sCode%>"><%=sFull%><BR></a>
<% Else %>
		<%=sFull%><BR>
<% End If %>
	</td>
	<td align=left valign=middle nowrap>
		<%=sCode%><BR>
	</td>
	<td align=left valign=middle nowrap>
		<%=sAbrev%><BR>
	</td>
  </tr>

<%
	        rsd.MoveNext
	    Loop

	rsd.Close
%>

</table>

<BR><BR>

<form>
	<input type=button value="  戻る  " onClick="JavaScript:window.history.back()">
	<input type=button value=" close " onClick="JavaScript:window.close()">
</form>

</center>
</body>
</html>

<%
%>
