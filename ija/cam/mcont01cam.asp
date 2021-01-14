<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log弌椡
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "2401", "実懷-僐儞僥僫斣崋徠夛乮拞墰傆摢乯", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb梡僞僌傪曇廤
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry name="p1" key="cont_e" format="*A" title="僐儞僥僫斣崋徠夛">
		<action type="accept" task="go" dest="#p2">
		<center>
		亂狠门斣崋徠夛亃<br><br>
		愭摢塸帤4寘:
	</entry>
	
	<entry name="p2" key="cont_s" format="*N">
		<action type="accept" task="go" dest="mcont02cam.asp?cont_e=$cont_e&cont_s=$cont_s">
		<center>
		亂狠门斣崋徠夛亃<br><br>
		悢帤晹暘7寘:
	</entry>

	</hdml>
<%
Else
	' EzWeb埲奜偺僞僌傪曇廤
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("僐儞僥僫斣崋徠夛")%>
	</head>
	<body>
	<center>
	亂狠门斣崋徠夛亃
	<hr>
	<form action="mcont02cam.asp" method="get">
		狠门斣崋擖椡<br>
		<table border="0">
			<tr><td>
				塸帤4寘:
				<input type="text" name="cont_e" maxlength="4" <%=GetTextSizeMode(4, "A")%>><br>
			</td></tr>
			<tr><td>
				悢帤:
				<input type="text" name="cont_s" maxlength="8" <%=GetTextSizeMode(8, "N")%>><br>
			</td></tr>
		</table>
		<input type="submit" value="寛掕">
	</form>
	<hr>

	<!--<br><a href="../mcont01.asp">崄捙丒ICCT徠夛</a>-->

	</body>
	</html>
<%
End If
%>
