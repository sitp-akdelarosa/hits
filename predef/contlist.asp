<%@Language="VBScript" %>

<!--#include file="Common.inc"-->

<%
	Dim sql
	Dim v_BookNo
	dim ObjConn, ObjRS	
	
	On Error Resume Next
	
	ConnDBH ObjConn, ObjRS	
		
	v_BookNo = Request.QueryString("BookNo")
	 
	if v_BookNo <> "" then
		sql = "SELECT DISTINCT Con.ContNo"
		sql = sql & " FROM ExportCont AS EXC" 
		sql = sql & " LEFT JOIN BookingAssign AS SPB ON EXC.BookNo=SPB.BookNo"
		sql = sql & " LEFT JOIN Container AS Con ON EXC.ContNo=Con.ContNo AND EXC.VoyCtrl=Con.VoyCtrl AND EXC.VslCode=Con.VslCode"
		sql = sql & " WHERE EXC.BookNo = '" & Trim(v_BookNo) & "'"
		sql = sql & " ORDER BY Con.ContNo"
	end if
	
	ObjRS.Open sql, ObjConn
	
	
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

<font size=4><b>�R���e�i�ꗗ</b></font>

<BR><BR>

<table border="1" cellspacing="1" cellpadding="3" bgcolor="#FFFFFF">
  <tr align="center" bgcolor="#FFCC33"> 
	<td align=center valign=middle height=15 nowrap>�R���e�i�ԍ�</td>	
  </tr>

<%
	    Do While Not ObjRS.EOF

			sCode  = Trim(ObjRS("ContNo"))
%>

  <tr>
	<td align=left valign=middle nowrap>
		<%=sCode%><BR>
	</td>
  </tr>

<%
	        ObjRS.MoveNext
	    Loop
	ObjRS.close
	
	DisConnDBH ObjConn, ObjRS	'DB�ؒf
%>

</table>

<BR><BR>

<form>	
	<input type=button value=" close " onClick="JavaScript:window.close()">
</form>

</center>
</body>
</html>

<%
%>
