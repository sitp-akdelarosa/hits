<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	Dim sql,sVsl,iCount
	ConnectSvr conn, rsd

	sVsl = Request.QueryString("vsl")
	iCount = 0

	sql = "SELECT VslCode,FullName,NameAbrev FROM mVessel WHERE ShipLine='" & sVsl & "'"
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

<font size=4><b>�D���R�[�h�ꗗ</b></font>

<BR><BR>

<table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
  <tr align="center" bgcolor="#FFCC33"> 
	<td align=center valign=middle height=15 nowrap>�D��</td>
	<td align=center valign=middle height=15 nowrap>�R�[�h</td>
	<td align=center valign=middle height=15 nowrap>����</td>
  </tr>

<%
	    Do While Not rsd.EOF

			sCode  = Trim(rsd("VslCode"))
			sFull  = Trim(rsd("FullName"))
			sAbrev = Trim(rsd("NameAbrev"))
%>

  <tr>
	<td align=left valign=middle nowrap>
		<%=sFull%><BR>
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
			iCount = iCount + 1
	    Loop

	rsd.Close

	If iCount=0 Then
%>

	<tr>
	  <td colspan=3 align=center valign=middle nowrap>
		�\���o����f�[�^������܂���
	  </td>
	</tr>

<% End If %>

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
