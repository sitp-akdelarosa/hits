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
WriteLogM oFS, "Unknown", "2401", "携帯-コンテナ番号照会（中央ふ頭）", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry name="p1" key="cont_e" format="*A" title="コンテナ番号照会">
		<action type="accept" task="go" dest="#p2">
		<center>
		【ｺﾝﾃﾅ番号照会】<br><br>
		先頭英字4桁:
	</entry>
	
	<entry name="p2" key="cont_s" format="*N">
		<action type="accept" task="go" dest="mcont02cam.asp?cont_e=$cont_e&cont_s=$cont_s">
		<center>
		【ｺﾝﾃﾅ番号照会】<br><br>
		数字部分7桁:
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
		<%=GetTitleTag("コンテナ番号照会")%>
	</head>
	<body>
	<center>
	【ｺﾝﾃﾅ番号照会】
	<hr>
	<form action="mcont02cam.asp" method="get">
		ｺﾝﾃﾅ番号入力<br>
		<table border="0">
			<tr><td>
				英字4桁:
				<input type="text" name="cont_e" maxlength="4" <%=GetTextSizeMode(4, "A")%>><br>
			</td></tr>
			<tr><td>
				数字:
				<input type="text" name="cont_s" maxlength="8" <%=GetTextSizeMode(8, "N")%>><br>
			</td></tr>
		</table>
		<input type="submit" value="決定">
	</form>
	<hr>

	<!--<br><a href="../mcont01.asp">香椎・ICCT照会</a>-->

	</body>
	</html>
<%
End If
%>
