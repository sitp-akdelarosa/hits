<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log出力
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "6200", "携帯-ログイン", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb用タグを編集
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry key="userid" format="*N" title="運行情報入力">
		<action type="accept" task="go" dest="mrung01.asp?UserID=$userid">
		<center>
		【完了時刻入力】<br><br>
		ユーザーID：
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
		<%=GetTitleTag("完了時刻入力")%>
	</head>
	
	<body>
	<center>
	【完了時刻入力】
	<hr>
	<form action="mrung01.asp" method="get">
		ユーザーID<br>
		<input type="text" name="UserID" maxlength="6" <%=GetTextSizeMode(6, "N")%>>
		<br><br>
		<input type="submit" value="決定" >
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
