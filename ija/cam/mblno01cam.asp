<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common_cam.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()
' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "2403", "携帯-BL番号照会（中央ふ頭）", "00" , sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry key="blno" title="BL番号照会">
		<action type="accept" task="go" dest="mblno02cam.asp?blno=$blno">
		<center>
		【BL番号照会】<br><br>
		BL番号：
	</entry>

	</hdml>
<%
Else
	' EzWeb以外のタグを編集
%>
	<html>
	
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("BL番号照会")%>
	</head>
	
	<body>
	<center>
	【BL番号照会】
	<hr>
	<form action="mblno02cam.asp" METHOD="get">
		BL番号入力<br>
		<input type="text" name="BLno" maxlength="20" <%=GetTextSizeMode(20, "A")%>><br>
		<br>
		<input type="submit" value="決定" >
	</form>
	<hr>

	<!--<br><a href="../mblno01.asp">香椎・ICCT照会</a>-->

	</body>
	
	</html>
<%
End If
%>
