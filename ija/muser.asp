<%@ LANGUAGE="VBScript" %>
<%Option Explicit%>
<!--#include file="common.inc"-->
<!--#include file="mcommon.inc"-->
<%
Dim sPhoneType
sPhoneType = GetPhoneType()

' Log�o��
Dim oFs
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
WriteLogM oFS, "Unknown", "6200", "�g��-���O�C��", "00",sPhoneType, ","
Set oFS = Nothing

If sPhoneType = "E" Then
	' EzWeb�p�^�O��ҏW
	Response.ContentType = "text/x-hdml; charset=Shift_JIS hdml"
%>
	<hdml version="3.0" public="true" markable="true">
	
	<entry key="userid" format="*N" title="�^�s������">
		<action type="accept" task="go" dest="mrung01.asp?UserID=$userid">
		<center>
		�y�����������́z<br><br>
		���[�U�[ID�F
	</entry>
	
	</hdml>
<%
Else
	' EzWeb�ȊO�̃^�O��ҏW
%>
	<html>
	<head>
		<meta http-equiv="Content-Language" content="ja">
		<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
		<%=GetTitleTag("������������")%>
	</head>
	
	<body>
	<center>
	�y�����������́z
	<hr>
	<form action="mrung01.asp" method="get">
		���[�U�[ID<br>
		<input type="text" name="UserID" maxlength="6" <%=GetTextSizeMode(6, "N")%>>
		<br><br>
		<input type="submit" value="����" >
	</form>
	<hr>
	</body>
	</html>
<%
End If
%>
