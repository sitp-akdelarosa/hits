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
			sKindName = "�׎�"
		Case 2
			sTblName  = "mForwarder"
			sCodeName = "Forwarder"
			sKindName = "�C��"
		Case 3
			sTblName  = "mTrucker"
			sCodeName = "Trucked"
			sKindName = "���^�Ǝ�"
		Case 4
			sTblName  = "mShipLine"
			sCodeName = "ShipLine"
			sKindName = "�D��"
		Case 5
			sTblName  = "mOperator"
			sCodeName = "OpeCode"
			sKindName = "�`�^"
		Case 6
			sTblName  = "mShipLine"
			sCodeName = "ShipLine"
			sKindName = "�D��"
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
<!-------------��������o�^�R�[�h�ꗗ���--------------------------->

<center>

<BR>

<font size=4><b><%=sKindName%>�R�[�h�ꗗ</b></font>

<BR><BR>

<table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
  <tr align="center" bgcolor="#FFCC33"> 
	<td align=center valign=middle height=15 nowrap>����</td>
	<td align=center valign=middle height=15 nowrap>�R�[�h</td>
	<td align=center valign=middle height=15 nowrap>����</td>
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
	<input type=button value="  �߂�  " onClick="JavaScript:window.history.back()">
	<input type=button value=" close " onClick="JavaScript:window.close()">
</form>

</center>
</body>
</html>

<%
%>
