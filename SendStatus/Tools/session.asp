<html>
<body>
<table bordercolor="silver" border="1" cellpadding="2" cellspacing="0">
<%	for each name in session.contents %>
<tr>
	<td><%=name%></td>
	<td><font color="blue"><%=session(name)%><br></font></td>
</tr>
<%	next %>
</table>
</body>
</html>
